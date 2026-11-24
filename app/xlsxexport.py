"""Выгрузка в Excel: таблицы и текст готового DOCX — по листу на страницу PDF.

Конвертер уже восстановил структуру таблиц (объединённые ячейки, строки по заливке,
подзаголовки групп), поэтому Excel строится из DOCX, а не из PDF заново:
    * таблица → ячейки с объединениями, границами, заливкой, жирным и выравниванием;
    * абзац → строка; табуляция делит строку на соседние ячейки (содержание, «Ду215  40  16»);
    * колонки листа и таблицы рядом (pdf2docx кладёт их во вложенную таблицу-макет) — рядом;
    * числа («50», «3,0», «105,0») — числами, с тем же числом знаков после запятой;
    * картинки, колонтитулы и штампы — не переносятся: это выгрузка данных.
"""
from __future__ import annotations

import copy
import logging
import re
from pathlib import Path

log = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PT_PER_CHAR = 5.6            # ширина символа Calibri 11 в пунктах — для ширины колонок Excel
MAX_COL_CHARS = 60
MIN_COL_CHARS = 4
LINE_FACTOR = 1.3            # высота строки текста = кегль × LINE_FACTOR
NUMBER = re.compile(r"^-?(0|[1-9]\d{0,14})([.,]\d+)?$")


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _val(el, tag: str, attr: str = "val"):
    if el is None:
        return None
    child = el.find(w(tag))
    return None if child is None else child.get(w(attr))


# --------------------------------------------------------------------------- текст и формат

def paragraph_text(p) -> str:
    """Текст абзаца; табуляция и разрыв строки сохраняются как \\t и \\n."""
    out = []
    for el in p.iter(w("t"), w("tab"), w("br"), w("cr")):
        if el.tag == w("t"):
            out.append(el.text or "")
        elif el.tag == w("tab"):
            out.append("\t")
        elif el.get(w("type")) not in ("page", "column"):
            out.append("\n")
    return "".join(out)


def _runs_with_text(p):
    for r in p.iter(w("r")):
        if any((t.text or "").strip() for t in r.iter(w("t"))):
            yield r


def paragraph_bold(p) -> bool:
    runs = list(_runs_with_text(p))
    if not runs:
        return False
    for r in runs:
        b = r.find(f"{w('rPr')}/{w('b')}")
        if b is None or b.get(w("val")) in ("0", "false", "off"):
            return False
    return True


def paragraph_size(p) -> float | None:
    for r in _runs_with_text(p):
        sz = _val(r.find(w("rPr")), "sz")
        if sz and sz.isdigit():
            return int(sz) / 2
    return None


def paragraph_font(p) -> str | None:
    for r in _runs_with_text(p):
        fonts = r.find(f"{w('rPr')}/{w('rFonts')}")
        if fonts is not None:
            name = fonts.get(w("ascii")) or fonts.get(w("hAnsi")) or fonts.get(w("cs"))
            if name:
                return name
    return None


def tab_stops(p) -> list[tuple[float, str]]:
    """Позиции табуляции абзаца (pt от левого поля) и их выравнивание; отступ абзаца — первым."""
    ppr = p.find(w("pPr"))
    ind = ppr.find(w("ind")) if ppr is not None else None
    left = 0.0
    if ind is not None:
        try:
            left = float(ind.get(w("left")) or ind.get(w("start")) or 0) / 20
        except ValueError:
            left = 0.0
    stops = []
    tabs = ppr.find(w("tabs")) if ppr is not None else None
    for t in (tabs if tabs is not None else []):
        kind = t.get(w("val")) or "left"
        if kind == "clear":
            continue
        try:
            stops.append((float(t.get(w("pos")) or 0) / 20, kind))
        except ValueError:
            continue
    return [(left, "left")] + sorted(stops)


def paragraph_align(p) -> str | None:
    jc = _val(p.find(w("pPr")), "jc")
    return {"center": "center", "right": "right", "end": "right", "both": "justify"}.get(jc or "")


def to_value(text: str):
    """Число, если это целиком число (запятая или точка — десятичный разделитель), иначе текст.
    Возвращает (значение, формат числа)."""
    s = text.strip()
    if not NUMBER.match(s):
        return text, None
    frac = re.split(r"[.,]", s)[1] if re.search(r"[.,]", s) else ""
    value = float(s.replace(",", ".")) if frac else int(s)
    return value, ("0." + "0" * len(frac)) if frac else "0"


# --------------------------------------------------------------------------- лист

LONG_TEXT = 60               # символов: длинный абзац переносится по словам в ширину таблиц листа
TEXT_WIDTH = 90              # символов: ширина текста на листе без таблиц


class _Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.widths: dict[int, float] = {}
        self.heights: dict[int, float] = {}
        self.long: list[tuple[int, int, str, float]] = []    # (строка, колонка, текст, кегль)
        self.bars: list[tuple[int, int, float]] = []         # (строка, колонка, ширина) — таблицы в 1 колонку
        self.tabbed: list[tuple[int, int, list, dict]] = []  # строки с табуляцией: раскладка после ширин
        self.keys: dict[tuple[int, int], tuple] = {}         # оформление записанных клеток
        self._styles: dict[tuple, object] = {}
        self.used = False

    def need_width(self, col: int, chars: float) -> None:
        chars = max(MIN_COL_CHARS, min(MAX_COL_CHARS, chars))
        self.widths[col] = max(self.widths.get(col, 0.0), chars)

    def need_height(self, row: int, pt: float) -> None:
        self.heights[row] = max(self.heights.get(row, 0.0), pt)

    def style(self, cell, key: tuple) -> None:
        """Оформление клетки по ключу (шрифт, кегль, жирный, выравнивание, перенос, заливка,
        границы, формат числа). Одинаковые наборы — копией готового стиля: openpyxl иначе заново
        сравнивает и хэширует объекты стилей для каждой клетки."""
        cached = self._styles.get(key)
        if cached is not None:
            cell._style = copy.copy(cached)
            return
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

        font, size, bold, h, v, wrap, fill, sides, fmt = key
        cell.font = Font(name=font, bold=bold, size=size)
        if h or v or wrap:
            cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
        if fill:
            cell.fill = PatternFill("solid", start_color=fill, end_color=fill)
        if sides:
            cell.border = Border(**{side: Side(style=kind, color=color) for side, kind, color in sides})
        if fmt:
            cell.number_format = fmt
        self._styles[key] = copy.copy(cell._style)

    def put(self, row: int, col: int, text: str, *, bold=False, size=None, font=None, align=None,
            valign=None, fill=None, wrap=False):
        cell = self.ws.cell(row=row, column=col)
        value, fmt = to_value(text) if "\n" not in text else (text, None)
        cell.value = value
        key = (font or "Calibri", size or 11, bool(bold), align, valign, bool(wrap), fill, (), fmt)
        self.style(cell, key)
        self.keys[(row, col)] = key
        self.used = True
        return cell

    def _column_at(self, col: int, x_pt: float) -> int:
        """Колонка, в которую попадает точка x (pt от левого края колонки col)."""
        c, edge = col, 0.0
        while c < col + 60:
            edge += self.widths.get(c, 8.43) * PT_PER_CHAR
            if x_pt < edge - 2:
                return c
            c += 1
        return c

    def _place_tabbed(self) -> None:
        """Части строки с табуляцией — в колонки по позициям табуляторов (содержание: номер
        страницы справа; «Минимальная толщина стенки → Таблица 1» над правым краем таблицы)."""
        from openpyxl.styles import Alignment

        for row, col, parts, style in self.tabbed:
            cols, prev = [], col - 1
            for _text, x, _kind in parts:
                c = max(self._column_at(col, x), prev + 1)
                cols.append(c)
                prev = c
            for k, ((text, _x, kind), c) in enumerate(zip(parts, cols)):
                cell = self.put(row, c, text, **style)
                last = cols[k + 1] - 1 if k + 1 < len(cols) else c
                if k + 1 == len(cols) and ("\n" in text or len(text) > LONG_TEXT):
                    total = self.widths.get(c, 8.43)       # хвост строки — на несколько колонок вправо
                    while total < 40 and last < c + 10:
                        last += 1
                        total += self.widths.get(last, 8.43)
                if "\n" in text or (last > c and len(text) > self.widths.get(c, 8.43)):
                    if last > c:
                        self.ws.merge_cells(start_row=row, start_column=c, end_row=row, end_column=last)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    width = sum(self.widths.get(i, 8.43) for i in range(c, last + 1))
                    lines = sum(max(1, -(-len(part) * 1.1 // max(width, 1))) for part in text.split("\n"))
                    self.need_height(row, lines * (style.get("size") or 11) * LINE_FACTOR)
                elif kind == "right" and k:
                    cell.alignment = Alignment(horizontal="right")

    def finish(self) -> None:
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter

        self._place_tabbed()

        # длинные абзацы: перенос по словам в ширину таблиц листа (или в широкую колонку A)
        last = max(self.widths, default=0)
        for row, col, text, size in self.long:
            cols = [c for c in range(col, last + 1)]
            width = sum(self.widths.get(c, 8.43) for c in cols)
            if width < 40:
                cols, width = [col], max(self.widths.get(col, 0.0), TEXT_WIDTH)
                self.widths[col] = width
            if len(cols) > 1 and not any(self.ws.cell(row=row, column=c).value for c in cols[1:]):
                self.ws.merge_cells(start_row=row, start_column=cols[0], end_row=row, end_column=cols[-1])
            cell = self.ws.cell(row=row, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            lines = sum(max(1, -(-len(part) * 1.1 // max(width, 1))) for part in text.split("\n"))
            self.need_height(row, lines * size * LINE_FACTOR)
        # таблица в одну колонку (плашка «Монтажный размер, мм») не задаёт ширину колонки,
        # а объединяет ячейки на свою ширину
        for row, col, target in self.bars:
            c, total = col, self.widths.get(col, 8.43)
            while total < 0.9 * target and c < col + 30:
                c += 1
                total += self.widths.get(c, 8.43)
            if c > col and not any(self.ws.cell(row=row, column=k).value for k in range(col + 1, c + 1)):
                src = self.ws.cell(row=row, column=col)
                for k in range(col + 1, c + 1):
                    dst = self.ws.cell(row=row, column=k)
                    dst.fill, dst.border = src.fill.copy(), src.border.copy()
                self.ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=c)
        for col, chars in self.widths.items():
            self.ws.column_dimensions[get_column_letter(col)].width = round(chars, 1)
        for row, pt in self.heights.items():
            if pt > 15.5:
                self.ws.row_dimensions[row].height = round(pt, 1)


def _render_elems(sheet: _Sheet, elems, row: int, col: int) -> tuple[int, int]:
    """Абзацы и таблицы друг под другом с (row, col). Возвращает (занято строк, колонок)."""
    r, width, blank = row, 1, False
    for el in elems:
        if el.tag == w("p"):
            text = paragraph_text(el)
            if not text.strip():
                blank = r > row        # подряд идущие пустые абзацы — одна пустая строка
                continue
            if blank:
                r += 1
                blank = False
            pieces = text.split("\t")
            bold, size, font = paragraph_bold(el), paragraph_size(el), paragraph_font(el)
            if len(pieces) > 1:
                # части — к позициям табуляторов (следующий за концом текста), табуляция в начале
                # строки — отступ; нет табулятора дальше — шаг по умолчанию 36 pt
                stops = tab_stops(el)
                placed, x, end = [], stops[0][0], stops[0][0]
                for k, piece in enumerate(pieces):
                    kind = "left"
                    if k and placed:
                        nxt = [st for st in stops[1:] if st[0] > end + 1]
                        x, kind = nxt[0] if nxt else ((end // 36 + 1) * 36, "left")
                    if piece.strip():
                        placed.append((piece.strip(), x, kind))
                        end = x if kind == "right" else x + len(piece.strip()) * (size or 11) * 0.5
                sheet.tabbed.append((r, col, placed, {"bold": bold, "size": size, "font": font}))
                sheet.used = True
                width = max(width, len(placed))
            else:
                part = text.strip()
                multiline = "\n" in part
                # без выравнивания по центру: в узкой колонке центрированный текст обрезается слева
                sheet.put(r, col, part, bold=bold, size=size, font=font, wrap=multiline)
                if len(part) > LONG_TEXT or multiline:
                    sheet.long.append((r, col, part, size or 11))
            r += 1
        elif el.tag == w("tbl"):
            if blank:
                r += 1
                blank = False
            h, wd = _render_table(sheet, el, r, col)
            r += h
            width = max(width, wd)
    return r - row, width


def _has_nested(tbl) -> bool:
    return any(tc.find(w("tbl")) is not None for tr in tbl.findall(w("tr")) for tc in tr.findall(w("tc")))


def _visible_borders(tbl) -> bool:
    """Есть ли у ячеек самой таблицы (не вложенных) видимые границы."""
    for tr in tbl.findall(w("tr")):
        for tc in tr.findall(w("tc")):
            if _cell_style(tc)[1]:
                return True
    return False


def _render_table(sheet: _Sheet, tbl, row: int, col: int) -> tuple[int, int]:
    # таблица-макет pdf2docx — без видимых границ, блоки во вложенных таблицах; у обычной таблицы
    # вложенная таблица в ячейке (заливка меньше ячейки) — просто текст ячейки
    if _has_nested(tbl) and not _visible_borders(tbl):
        return _render_layout(sheet, tbl, row, col)
    return _render_grid(sheet, tbl, row, col)


def _render_layout(sheet: _Sheet, tbl, row: int, col: int) -> tuple[int, int]:
    """Таблица-макет pdf2docx (блоки рядом): содержимое ячеек — рядом, через пустую колонку."""
    r, width = row, 1
    for tr in tbl.findall(w("tr")):
        c, height = col, 0
        for tc in tr.findall(w("tc")):
            content = [el for el in tc if el.tag in (w("p"), w("tbl"))]
            h, wd = _render_elems(sheet, content, r, c)
            if h:
                c += wd + 1
                height = max(height, h)
        width = max(width, c - col - 1)
        r += height
    return r - row, width


def _cell_style(tc):
    tcpr = tc.find(w("tcPr"))
    fill = None
    shd = tcpr.find(w("shd")) if tcpr is not None else None
    if shd is None and tc.find(w("tbl")) is not None:      # заливка во вложенной таблице
        shd = next(tc.find(w("tbl")).iter(w("shd")), None)
    if shd is not None:
        f = (shd.get(w("fill")) or "").lstrip("#").upper()
        if re.fullmatch(r"[0-9A-F]{6}", f) and f != "FFFFFF":
            fill = f
    borders = {}
    tb = tcpr.find(w("tcBorders")) if tcpr is not None else None
    if tb is not None:
        for side, names in (("top", ("top",)), ("bottom", ("bottom",)),
                            ("left", ("left", "start")), ("right", ("right", "end"))):
            for name in names:
                b = tb.find(w(name))
                if b is None:
                    continue
                kind = b.get(w("val"))
                try:
                    sz = float(b.get(w("sz")) or 4)     # pdf2docx пишет «4.0»
                except ValueError:
                    sz = 4.0
                if kind in (None, "nil", "none") or sz <= 0:
                    continue
                color = (b.get(w("color")) or "").lstrip("#").upper()
                borders[side] = ("medium" if sz >= 12 else "thin",
                                 color if re.fullmatch(r"[0-9A-F]{6}", color) else "000000")
                break
    valign = {"center": "center", "bottom": "bottom"}.get(_val(tcpr, "vAlign") or "", "top")
    return fill, borders, valign


def _paragraph_spans(tc) -> list[tuple[float, float, str]]:
    """(верх, низ, текст) абзацев ячейки от её верха — по интервалам и высоте строк, которые
    pdf2docx ставит по PDF (строки PDF сохранены разрывами, поэтому перенос не оценивается)."""
    from .docxfix import paragraph_height

    out, y = [], 0.0
    for p in tc.findall(w("p")):
        sp = p.find(f"{w('pPr')}/{w('spacing')}")
        before = after = 0.0
        if sp is not None:
            try:
                before = float(sp.get(w("before"), "0")) / 20
                after = float(sp.get(w("after"), "0")) / 20
            except ValueError:
                pass
        h = paragraph_height(p)
        text = paragraph_text(p).replace("\t", " ").strip()
        if text:
            out.append((y + before, y + h - after, text))
        y += h
    return out


def _split_row(tr):
    """Строка таблицы без линий между записями (спецификация ГОСТ: позиции идут столбиком в одной
    высокой ячейке) — несколько строк Excel. Запись — абзацы разных колонок, перекрывающиеся по
    высоте (позиция в 4 строки, наименование, обозначение, «компл.», «4» — одна запись).
    Возвращает [{колонка сетки: текст}] или None, если делить нечего."""
    cells, gi = [], int(_val(tr.find(w("trPr")), "gridBefore") or 0)
    for tc in tr.findall(w("tc")):
        tcpr = tc.find(w("tcPr"))
        if tcpr is not None and tcpr.find(w("vMerge")) is not None or tc.find(w("tbl")) is not None:
            return None
        cells.append((gi, _paragraph_spans(tc)))
        gi += int(_val(tcpr, "gridSpan") or 1)
    if sum(1 for _, spans in cells if len(spans) >= 2) < 2:
        return None
    marks = sorted((top, bottom, g, t) for g, spans in cells for top, bottom, t in spans)
    records: list[dict[int, list[str]]] = []
    reach = None
    for top, bottom, g, t in marks:
        if reach is None or top >= reach - 1.0:
            records.append({})
            reach = bottom
        else:
            reach = max(reach, bottom)
        records[-1].setdefault(g, []).append(t)
    if len(records) < 2:
        return None
    return [{g: "\n".join(ts) for g, ts in rec.items()} for rec in records]


def _render_grid(sheet: _Sheet, tbl, row: int, col: int) -> tuple[int, int]:
    grid = [int(g.get(w("w")) or 0) for g in tbl.findall(f"{w('tblGrid')}/{w('gridCol')}")]
    rows = tbl.findall(w("tr"))
    ncols = max(len(grid), 1)
    merges: list[list[int]] = []          # [r0, c0, r1, c1]
    open_v: dict[int, list[int]] = {}     # колонка сетки → объединение по вертикали
    styled: list[tuple[list[int], dict]] = []
    r = row
    for tr in rows:
        trh = tr.find(f"{w('trPr')}/{w('trHeight')}")
        row_pt = int(trh.get(w("val")) or 0) / 20 if trh is not None else 0.0
        lines = _split_row(tr)
        count = len(lines) if lines else 1
        for j in range(count):
            gi = int(_val(tr.find(w("trPr")), "gridBefore") or 0)
            lines_pt = 0.0
            for tc in tr.findall(w("tc")):
                tcpr = tc.find(w("tcPr"))
                span = int(_val(tcpr, "gridSpan") or 1)
                vm = tcpr.find(w("vMerge")) if tcpr is not None else None
                vmode = None if vm is None else (vm.get(w("val")) or "continue")
                c0, c1 = col + gi, col + gi + span - 1
                fill, borders, valign = _cell_style(tc)
                if lines:      # строка формы: верхняя граница — у первой, нижняя — у последней
                    borders = {k: v for k, v in borders.items()
                               if (k != "top" or j == 0) and (k != "bottom" or j == count - 1)}
                if vmode == "continue" and gi in open_v:
                    open_v[gi][2] = r          # продолжение объединения по вертикали
                    gi += span
                    continue
                for g in range(gi, gi + span):   # новая ячейка закрывает объединения над собой
                    if g in open_v:
                        merges.append(open_v.pop(g))
                paras = list(tc.iter(w("p")))          # и абзацы вложенных таблиц
                first = next((p for p in paras if paragraph_text(p).strip()), None)
                if lines:
                    text = lines[j].get(gi, "")         # запись формы
                else:
                    text = "\n".join(t for t in (paragraph_text(p).replace("\t", " ").strip() for p in paras) if t)
                size = paragraph_size(first) if first is not None else None
                sheet.put(r, c0, text, bold=first is not None and paragraph_bold(first), size=size,
                          font=paragraph_font(first) if first is not None else None,
                          align=paragraph_align(first) if first is not None else None, valign=valign,
                          fill=fill, wrap=True)
                rng = [r, c0, r, c1]
                if vmode == "restart":
                    open_v[gi] = rng
                elif span > 1:
                    merges.append(rng)
                if text and vmode != "restart":
                    width_pt = sum(grid[gi:gi + span]) / 20 or 60
                    est = sum(max(1, -(-len(line) * (size or 11) * 0.5 // max(width_pt, 1)))
                              for line in text.split("\n"))
                    lines_pt = max(lines_pt, est * (size or 11) * LINE_FACTOR)
                styled.append((rng, {"fill": fill, "borders": borders}))
                gi += span
            sheet.need_height(r, lines_pt if lines else max(row_pt, lines_pt))
            r += 1
    merges.extend(open_v.values())
    for rng in merges:
        if rng[0] != rng[2] or rng[1] != rng[3]:
            sheet.ws.merge_cells(start_row=rng[0], start_column=rng[1], end_row=rng[2], end_column=rng[3])
    # границы и заливка — каждой клетке объединения (так их показывает Excel)
    spans = {(m[0], m[1]): m for m in merges}
    for rng, style in styled:
        r0, c0, r1, c1 = spans.get((rng[0], rng[1]), rng)
        b = style["borders"]
        for rr in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1):
                sides = tuple((side, kind, color) for side, (kind, color) in sorted(b.items())
                              if {"top": rr == r0, "bottom": rr == r1, "left": cc == c0, "right": cc == c1}[side])
                base = sheet.keys.get((rr, cc), ("Calibri", 11, False, None, None, False, None, (), None))
                key = base[:6] + (style["fill"] or base[6], sides, base[8])
                if key != base:
                    sheet.style(sheet.ws.cell(row=rr, column=cc), key)
    if len(grid) == 1:
        sheet.bars.extend((k, col, grid[0] / 20 / PT_PER_CHAR) for k in range(row, r))
    else:
        for i, twips in enumerate(grid):
            sheet.need_width(col + i, twips / 20 / PT_PER_CHAR)
    return r - row, ncols


# --------------------------------------------------------------------------- документ

def _sheet_title(n: int, used: set[str]) -> str:
    title = f"Стр. {n}"
    k = 2
    while title in used:
        title = f"Стр. {n} ({k})"
        k += 1
    used.add(title)
    return title


def docx_to_xlsx(docx_path: Path, xlsx_path: Path, page_numbers: list[int] | None = None) -> int:
    """Пишет книгу Excel по DOCX конвертера. Возвращает число листов (0 — нечего выгружать,
    файл не создаётся)."""
    import docx
    from openpyxl import Workbook

    from .docxfix import _page_groups

    document = docx.Document(str(docx_path))
    pages = _page_groups(document)
    if page_numbers is None or len(page_numbers) != len(pages):
        page_numbers = list(range(1, len(pages) + 1))
    wb = Workbook()
    wb.remove(wb.active)
    titles: set[str] = set()
    count = 0
    for number, (_section, parts) in zip(page_numbers, pages):
        ws = wb.create_sheet(_sheet_title(number, titles))
        sheet = _Sheet(ws)
        r = 1
        for part in parts:
            if part.cols > 1:
                c, height = 1, 0
                for chunk in part.chunks:
                    h, wd = _render_elems(sheet, chunk, r, c)
                    if h:
                        c += wd + 1
                        height = max(height, h)
                r += height
            else:
                h, _ = _render_elems(sheet, part.elems, r, 1)
                r += h
        if not sheet.used:
            wb.remove(ws)
            titles.discard(ws.title)
            continue
        sheet.finish()
        count += 1
    if count:
        wb.save(str(xlsx_path))
    return count
