"""Исправления поверх pdf2docx.

1. Нормализация сырого текста страницы (до разбора макета):
   * склейка фрагментов одной строки, которые MuPDF разрезал из-за больших пробелов
     (выравнивание по ширине) — иначе pdf2docx принимает абзацы за таблицы;
   * склейка маркера списка (•, 1., а)) с текстом пункта;
   * пробел между соседними фрагментами строки, если в PDF его не было («1.Анализ» → «1. Анализ»);
   * символы из Private Use Area шрифтов Symbol/Wingdings → обычный Unicode (•, ▪, ✓ …).
2. Проверка «потоковых» (безрамочных) таблиц: если найденная таблица похожа на разорванный текст,
   страница разбирается повторно без поиска таких таблиц.
"""
from __future__ import annotations

import copy
import io
import logging
import pickle
import re

log = logging.getLogger(__name__)

# --------------------------------------------------------------------------- символы

_SYMBOL_PUA = {
    0xF0B7: "•", 0xF0A7: "▪", 0xF02D: "−", 0xF0B0: "°", 0xF0B1: "±", 0xF0B4: "×", 0xF0B8: "÷",
    0xF0A3: "≤", 0xF0B3: "≥", 0xF0B9: "≠", 0xF0BB: "≈", 0xF0AE: "→", 0xF0AC: "←", 0xF0DE: "⇒",
    0xF0D7: "⋅", 0xF0E0: "◊", 0xF06F: "○", 0xF0A8: "□",
}
_WINGDINGS_PUA = {
    0xF06C: "●", 0xF06E: "■", 0xF071: "❑", 0xF075: "◆", 0xF076: "❖", 0xF0A7: "▪", 0xF0A8: "◻",
    0xF0D8: "➢", 0xF0E8: "➔", 0xF0FC: "✓", 0xF0FB: "✗", 0xF0FE: "☑", 0xF09F: "•", 0xF0B7: "•",
    0xF0F0: "⇨", 0xF0E0: "➢", 0xF077: "⬥", 0xF0B2: "✧",
}

_LIST_MARKERS = set("•◦▪▫■□●○‣⁃–—−-*·➢✓❖◆➔⇨❑►▸")   # «−» — минус из шрифта Symbol (списки Word)
_LIST_MARKER_RE = re.compile(r"^(\(?\d{1,3}[.)]|\d{1,2}\.\d{1,2}\.?|[a-zа-яё][.)]|[IVXLC]{1,5}\.)$", re.I)


def _fix_symbol_chars(span: dict) -> None:
    font = (span.get("font") or "").lower()
    if "wingding" in font:
        table = _WINGDINGS_PUA
    elif "symbol" in font:
        table = _SYMBOL_PUA
    else:
        table = None
    changed = False
    for ch in span.get("chars", ()):
        c = ch.get("c", "")
        if len(c) == 1 and 0xE000 <= ord(c) <= 0xF8FF:
            repl = (table or {}).get(ord(c)) or _SYMBOL_PUA.get(ord(c)) or _WINGDINGS_PUA.get(ord(c))
            if repl:
                ch["c"] = repl
                changed = True
    if changed and table is not None:
        # шрифт-символ заменяем обычным, чтобы Word не показывал «квадратики»
        span["font"] = "Arial"


# --------------------------------------------------------------------------- строки

def _horizontal(line: dict) -> bool:
    d = line.get("dir", (1, 0))
    return abs(d[0] - 1) < 1e-3 and abs(d[1]) < 1e-3 and line.get("wmode", 0) == 0


def _chars(line: dict):
    for s in line["spans"]:
        yield from s.get("chars", ())


def _first_char(line: dict) -> dict | None:
    for s in line["spans"]:
        if s.get("chars"):
            return s["chars"][0]
    return None


def _last_char(line: dict) -> dict | None:
    for s in reversed(line["spans"]):
        if s.get("chars"):
            return s["chars"][-1]
    return None


def _line_size(line: dict) -> float:
    return max((s.get("size", 0) for s in line["spans"]), default=0) or 10.0


def _line_text(line: dict) -> str:
    return "".join(c.get("c", "") for c in _chars(line))


def _is_list_marker(text: str) -> bool:
    t = text.strip()
    return bool(t) and (t in _LIST_MARKERS or bool(_LIST_MARKER_RE.match(t)))


def _space_char(after: dict, gap: float, size: float, fill: bool = False) -> dict:
    """Синтетический пробел после символа after. fill=True — пробел закрывает весь промежуток,
    иначе pdf2docx разрежет строку по большому промежутку и пробел потеряется."""
    x0 = after["bbox"][2]
    w = max(0.1, gap if fill else min(gap, size * 0.25))
    return {"c": " ", "origin": (x0, after["origin"][1]),
            "bbox": (x0, after["bbox"][1], x0 + w, after["bbox"][3]), "synthetic": True}


MARKER_TAB_GAP = 2.5      # промежуток после маркера списка больше 2,5 пробела — это табуляция Word
_SPACE_EM = 0.28          # ширина пробела Arial/Times в долях кегля


def _tab_char(after: dict, text_x: float) -> dict:
    """Синтетическая табуляция от символа after до начала текста text_x (позиция табуляции)."""
    x0 = after["bbox"][2]
    return {"c": "\t", "origin": (x0, after["origin"][1]),
            "bbox": (x0, after["bbox"][1], max(x0 + 0.1, text_x), after["bbox"][3]), "synthetic": True}


def marker_gap_is_tab(gap: float, size: float) -> bool:
    """«−<TAB>текст» в списке Word: от маркера до текста заметно больше обычного пробела."""
    return gap > MARKER_TAB_GAP * _SPACE_EM * size


def _strip_trailing_spaces(line: dict) -> dict | None:
    """Убрать пробелы в конце строки; вернуть последний оставшийся символ."""
    for span in reversed(line["spans"]):
        chars = span.get("chars") or []
        while chars and chars[-1].get("c", "").isspace():
            chars.pop()
        if chars:
            return chars[-1]
    return None


def _union(a, b):
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


class Barriers:
    """Вертикальные линии страницы (границы ячеек): фрагменты по разные стороны — разные ячейки."""

    def __init__(self, lines=()):
        self._lines = sorted(lines)               # (x, y0, y1)
        self._xs = [ln[0] for ln in self._lines]

    def __bool__(self) -> bool:
        return bool(self._lines)

    def between(self, x0: float, x1: float, y0: float, y1: float) -> bool:
        """Есть ли линия с x в (x0, x1), перекрывающая по высоте половину строки [y0, y1]."""
        from bisect import bisect_left, bisect_right
        need = 0.5 * max(y1 - y0, 0.1)
        for k in range(bisect_left(self._xs, x0), bisect_right(self._xs, x1)):
            _, a, b = self._lines[k]
            if min(b, y1) - max(a, y0) >= need:
                return True
        return False


NO_BARRIERS = Barriers()
BARRIER_THICKNESS = 3.0     # pt: Word рисует линии таблиц тонкими заливками


def page_barriers(page) -> Barriers:
    """Вертикальные линии и тонкие вертикальные заливки страницы PyMuPDF."""
    try:
        drawings = page.get_cdrawings()
    except Exception:
        return NO_BARRIERS
    out = []
    for d in drawings:
        x0, y0, x1, y1 = d.get("rect", (0, 0, 0, 0))
        if d.get("fill") is not None and d.get("color") is None:
            if x1 - x0 <= BARRIER_THICKNESS and y1 - y0 > x1 - x0:
                out.append(((x0 + x1) / 2, y0, y1))
            continue
        for it in d.get("items", ()):
            if it[0] == "l":
                (ax, ay), (bx, by) = tuple(it[1]), tuple(it[2])
                if abs(ax - bx) <= 0.5 and abs(ay - by) > 1.0:
                    out.append(((ax + bx) / 2, min(ay, by), max(ay, by)))
            elif it[0] == "re":
                rx0, ry0, rx1, ry1 = tuple(it[1])
                if ry1 - ry0 > 1.0:
                    out.extend(((rx0, ry0, ry1), (rx1, ry0, ry1)))
    return Barriers(out)


def _try_merge(a: dict, b: dict, barriers: Barriers = NO_BARRIERS) -> bool:
    """Склеить строку b в конец строки a, если это продолжение той же строки текста."""
    if not (_horizontal(a) and _horizontal(b)):
        return False
    ca_last, cb_first = _last_char(a), _first_char(b)
    if ca_last is None or cb_first is None:
        return False
    size = max(_line_size(a), _line_size(b))
    if abs(ca_last["origin"][1] - cb_first["origin"][1]) > 0.2 * size:
        return False
    gap = b["bbox"][0] - a["bbox"][2]
    if gap < -0.5 or gap > 15 * size:
        return False
    if barriers and barriers.between(a["bbox"][2] - 0.5, b["bbox"][0] + 0.5,
                                     min(a["bbox"][1], b["bbox"][1]), max(a["bbox"][3], b["bbox"][3])):
        return False      # между фрагментами линия таблицы: это соседние ячейки
    has_space = ca_last["c"].isspace() or cb_first["c"].isspace()
    marker = _is_list_marker(_line_text(a))
    if marker and not cb_first["c"].isspace():
        # «−» и текст пункта — разные строки PDF; промежуток между ними в Word был табуляцией
        head = _last_char(a) if not ca_last["c"].isspace() else None
        tail = head or _strip_trailing_spaces(a)
        if tail is not None and marker_gap_is_tab(b["bbox"][0] - tail["bbox"][2], size) \
                and b["bbox"][0] - tail["bbox"][2] < 6 * size:
            last_span = next(sp for sp in reversed(a["spans"]) if sp.get("chars"))
            last_span["chars"].append(_tab_char(tail, b["bbox"][0]))
            a["spans"].extend(b["spans"])
            a["bbox"] = _union(a["bbox"], b["bbox"])
            return True
    if not has_space:
        if not (marker and gap < 4 * size):
            return False
        a["spans"][-1]["chars"].append(_space_char(ca_last, gap, size, fill=True))
    a["spans"].extend(b["spans"])
    a["bbox"] = _union(a["bbox"], b["bbox"])
    return True


def _absorb_blank_spans(line: dict) -> None:
    """Фрагменты из одних пробелов pdf2docx отбрасывает — переносим их в предыдущий фрагмент."""
    out = []
    for span in line["spans"]:
        chars = span.get("chars") or []
        if out and chars and all(c.get("c", "").isspace() for c in chars):
            prev = out[-1]
            prev["chars"].extend({**c, "c": " "} for c in chars[:1])
            b = prev["bbox"]
            prev["bbox"] = (b[0], b[1], max(b[2], chars[-1]["bbox"][2]), b[3])
            continue
        out.append(span)
    line["spans"] = out


def _insert_missing_spaces(line: dict) -> None:
    spans = [s for s in line["spans"] if s.get("chars")]
    head = ""
    for prev, nxt in zip(spans, spans[1:]):
        head += "".join(c.get("c", "") for c in prev["chars"])
        c1, c2 = prev["chars"][-1], nxt["chars"][0]
        if c1["c"].isspace() or c2["c"].isspace():
            continue
        size = max(prev.get("size", 10), nxt.get("size", 10))
        gap = c2["bbox"][0] - c1["bbox"][2]
        if gap > 0.2 * size:
            marker = _is_list_marker(head) and gap < 4 * size
            if marker and marker_gap_is_tab(gap, size):
                prev["chars"].append(_tab_char(c1, c2["bbox"][0]))
                continue
            prev["chars"].append(_space_char(c1, gap, size, fill=marker))


_NO_SPACE_AFTER = set("-\u00ad\u2010\u2011/")   # дефисы: перенос слова, пробел не нужен
_NEVER_WORD_START = set("ьъы")                 # с этих букв не начинается ни одно слово


def _is_cjk(c: str) -> bool:
    o = ord(c)
    return 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF


def _add_line_end_space(line: dict, nxt: dict) -> None:
    """pdf2docx склеивает строки абзаца без пробела, если строка в PDF не кончается пробелом."""
    last, first = _last_char(line), _first_char(nxt)
    if last is None or first is None:
        return
    c, n = last["c"], first["c"]
    if not c or c.isspace() or n.isspace() or c in _NO_SPACE_AFTER or _is_cjk(c) or _is_cjk(n[:1] or " "):
        return
    if n in _NEVER_WORD_START and c.isalpha():
        return  # слово разорвано без дефиса (узкая ячейка): «Удовлетворённост|ь»
    size = _line_size(line)
    if first["origin"][1] - last["origin"][1] < 0.5 * size:  # следующая строка не ниже — не перенос
        return
    line["spans"][-1]["chars"].append(_space_char(last, size * 0.25, size))


def normalize_text_blocks(blocks: list[dict], barriers: Barriers = NO_BARRIERS) -> list[dict]:
    for block in blocks:
        lines = block.get("lines")
        if not lines:
            continue
        for line in lines:
            for span in line["spans"]:
                _fix_symbol_chars(span)
        merged = [lines[0]]
        for ln in lines[1:]:
            if not _try_merge(merged[-1], ln, barriers):
                merged.append(ln)
        for ln in merged:
            _absorb_blank_spans(ln)
            _insert_missing_spaces(ln)
        for ln, nxt in zip(merged, merged[1:]):
            _add_line_end_space(ln, nxt)
        block["lines"] = merged
        bb = merged[0]["bbox"]
        for ln in merged[1:]:
            bb = _union(bb, ln["bbox"])
        block["bbox"] = bb
    return blocks


# --------------------------------------------------------------------------- патчи pdf2docx

_installed = False


def install_patches() -> None:
    """Подменяет извлечение текста в pdf2docx на версию с нормализацией (идемпотентно)."""
    global _installed
    if _installed:
        return
    from pdf2docx.page.RawPageFitz import RawPageFitz

    original = RawPageFitz._preprocess_text

    def _preprocess_text(self, **settings):
        blocks = original(self, **settings)
        try:
            from .regionfix import filter_text
            blocks = filter_text(self.page_engine, blocks)  # проходы «тело/колонтитул» (app/furniture.py)
        except Exception:
            log.exception("region filter failed")
        try:
            return normalize_text_blocks(blocks, page_barriers(self.page_engine))
        except Exception:  # нормализация не должна ломать конвертацию
            log.exception("text normalization failed")
            return blocks

    RawPageFitz._preprocess_text = _preprocess_text
    _patch_paragraph_split()
    try:
        _patch_marker_tabs()
    except Exception:
        log.exception("marker tab patch failed")
    try:
        _patch_merged_cell_align()
    except Exception:
        log.exception("merged cell alignment patch failed")
    try:
        _patch_char_scaling()
    except Exception:
        log.exception("char scaling patch failed")
    try:
        _patch_two_line_indent()
    except Exception:
        log.exception("two-line indent patch failed")
    try:
        from . import vectorfix
        vectorfix.install()
    except Exception:  # без патча графики конвертация всё равно работает
        log.exception("vector graphics patch failed")
    try:
        from . import regionfix
        regionfix.install()
    except Exception:
        log.exception("region filter patch failed")
    try:
        from . import imagetrim
        imagetrim.install()
    except Exception:
        log.exception("image rotation patch failed")
    try:
        from . import speedups
        speedups.install()
    except Exception:  # без ускорения результат тот же, только медленнее
        log.exception("speedup patch failed")
    _installed = True


def _row_starts_list_item(text: str) -> bool:
    text = text.lstrip()
    if not text:
        return False
    if text[0] in _LIST_MARKERS and (len(text) == 1 or text[1].isspace()):
        return True
    first = text.split(None, 1)[0]
    return bool(_LIST_MARKER_RE.match(first)) and len(text) > len(first)


LIST_PREV_END = (":", ";", ".", "!", "?", ",")


_DASHES = set("–—−-")


def list_item_follows(prev_text: str, prev_width: float, full_width: float, text: str = "–") -> bool:
    """Строка с маркером списка в начале — новый пункт списка?

    Точка-маркер «•» или маркер с табуляцией после него (список Word) — всегда новый пункт. Тире
    бывает и посреди фразы: «…четыре нитки «ленточный сгуститель / – ленточный фильтр-пресс» со
    всем…», поэтому строка с тире — новый пункт, только если предыдущая строка закончилась (знак
    препинания в конце или не дошла до правого края)."""
    t = text.lstrip()
    if len(t) > 1 and t[1] == "\t":
        return True
    if t[:1] not in _DASHES:
        return True
    return prev_text.rstrip().endswith(LIST_PREV_END) or prev_width < 0.9 * full_width


# строка оглавления: «… название ........ 16» — на ней пункт оглавления заканчивается
TOC_ENTRY_END_RE = re.compile(r"(?:[.·…]\s*){4,}\d{1,4}\s*$")
SIZE_JUMP = 0.2


def toc_entry_ends(text: str) -> bool:
    return bool(TOC_ENTRY_END_RE.search(text))


def size_jump(prev_size: float, size: float) -> bool:
    """Кегль соседних строк отличается на 20 % и больше — это разные абзацы (заголовок
    «Содержание» 15 pt над пунктами 12 pt), а не перенос внутри одного."""
    return prev_size > 0 and size > 0 and abs(size - prev_size) >= SIZE_JUMP * max(prev_size, size)


def _row_size(row) -> float:
    """Кегль большинства букв строки (картинки и пробелы не считаются)."""
    weight: dict[float, int] = {}
    for ln in row:
        for sp in getattr(ln, "spans", ()):
            text = getattr(sp, "text", "") or ""
            size = getattr(sp, "size", 0.0) or 0.0
            n = len(text.strip())
            if n and size:
                key = round(size, 1)
                weight[key] = weight.get(key, 0) + n
    return max(weight, key=weight.get) if weight else 0.0


def marker_tab_stops(block_x0: float, left_space: float, tab_ends: list[float]) -> list[float]:
    """Позиции табуляций (pt от поля) для «−<TAB>текст»: там, где текст начинался в PDF."""
    out: list[float] = []
    for x in tab_ends:
        pos = round(left_space + (x - block_x0), 1)
        if pos > 0 and all(abs(pos - o) > 1.0 for o in out):
            out.append(pos)
    return out


def _patch_marker_tabs() -> None:
    """Табуляция после маркера списка доходит до того же места, что в PDF (своя позиция табуляции)."""
    from docx.shared import Pt
    from pdf2docx.text.TextBlock import TextBlock

    original = TextBlock.make_docx

    def make_docx(self, p):
        res = original(self, p)
        try:
            ends = [c.bbox[2] for ln in self.lines for sp in getattr(ln, "spans", ())
                    for c in getattr(sp, "chars", ()) if getattr(c, "c", "") == "\t"]
            if ends:
                have = [ts.position.pt for ts in p.paragraph_format.tab_stops]
                for pos in marker_tab_stops(self.bbox[0], self.left_space, ends):
                    if all(abs(pos - h) > 1.0 for h in have):
                        p.paragraph_format.tab_stops.add_tab_stop(Pt(pos))
        except Exception:
            log.debug("marker tab stop skipped", exc_info=True)
        return res

    TextBlock.make_docx = make_docx


CENTER_MIN_GAP = 12.0     # pt: меньше — обычные поля ячейки, а не выравнивание по центру


def cell_vertical_align(cell_box, content_box) -> str | None:
    """'center' / 'bottom', если текст объединённой по вертикали ячейки стоит посередине / внизу.

    pdf2docx ставит такой текст отступом перед абзацем (114 pt в «Где проводится: Лаборатория…» на
    три строки таблицы). Word и LibreOffice отдают этот отступ первой строке объединения — она
    вырастает, и таблица уходит на следующую страницу. Выравнивание ячейки даёт тот же вид без
    этого."""
    top = content_box[1] - cell_box[1]
    bottom = cell_box[3] - content_box[3]
    if top < CENTER_MIN_GAP:
        return None
    if abs(top - bottom) <= max(6.0, 0.25 * (top + bottom)):
        return "center"
    if bottom < 0.25 * top and bottom < CENTER_MIN_GAP:
        return "bottom"
    return None


def _drop_blank_blocks(cell) -> None:
    """Абзац из одних пробелов в ячейке (хвосты строк соседней графы с выравниванием по ширине)
    ничего не несёт, а с огромным правым отступом Word/LibreOffice разбивают его на десяток строк —
    строка таблицы вырастает, и лист переезжает на следующую страницу."""
    blocks = list(cell.blocks)
    keep = [b for b in blocks if not (getattr(b, "is_text_block", False) and not b.text.strip()
                                      and not any(getattr(sp, "image", None) is not None
                                                  for ln in b.lines for sp in ln.spans))]
    if len(keep) != len(blocks):
        cell.blocks.reset(keep)


def _patch_merged_cell_align() -> None:
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.shared import Pt
    from pdf2docx.table.Cell import Cell

    original = Cell.make_docx
    modes = {"center": WD_CELL_VERTICAL_ALIGNMENT.CENTER, "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}

    def make_docx(self, table, indexes):
        try:
            _drop_blank_blocks(self)
        except Exception:
            log.debug("blank cell blocks kept", exc_info=True)
        original(self, table, indexes)
        try:
            if not bool(self) or not self.blocks or self.merged_cells[0] < 2:
                return
            boxes = [b.bbox for b in self.blocks]
            content = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                       max(b[2] for b in boxes), max(b[3] for b in boxes))
            mode = cell_vertical_align(tuple(self.bbox), content)
            if mode is None:
                return
            docx_cell = table.cell(*indexes)
            docx_cell.vertical_alignment = modes[mode]
            if docx_cell.paragraphs:
                docx_cell.paragraphs[0].paragraph_format.space_before = Pt(0)
        except Exception:
            log.debug("merged cell alignment skipped", exc_info=True)

    Cell.make_docx = make_docx


def _patch_paragraph_split() -> None:
    """Включает в pdf2docx разбиение абзаца перед пунктом списка (в оригинале отключено: `if 0 and …`)."""
    from pdf2docx.common import constants
    from pdf2docx.text.Lines import Lines

    def split_vertically_by_text(self, line_break_free_space_ratio: float, new_paragraph_free_space_ratio: float):
        rows = self.group_by_physical_rows()
        for row in rows:
            row.sort_in_line_order()
        num = len(rows)
        if num == 1:
            return rows
        W = max(row[-1].bbox[2] - row[0].bbox[0] for row in rows[1:])
        H = sum(row[0].bbox[3] - row[0].bbox[1] for row in rows) / num
        res = []
        lines = Lines()
        punc = tuple(constants.SENTENCE_END_PUNC)
        start_of_para = end_of_para = False
        start_of_sen = end_of_sen = False
        prev_w, prev_text, prev_size = 0.0, "", 0.0
        for i, row in enumerate(rows):
            end_of_sen = row[-1].text.strip().endswith(punc)
            w = row[-1].bbox[2] - row[0].bbox[0]
            text = "".join(ln.text for ln in row)
            size = _row_size(row)
            if i > 0 and size_jump(prev_size, size):
                start_of_para = True
            elif i > 0 and _row_starts_list_item(text) and list_item_follows(prev_text, prev_w, W, text):
                start_of_para = True
            elif end_of_sen and w / W <= 1.0 - line_break_free_space_ratio:
                end_of_para = True
            elif start_of_sen and (W - w) / H >= new_paragraph_free_space_ratio:
                start_of_para = True
            if toc_entry_ends(text):
                end_of_para = True
            if start_of_para and lines:
                res.append(lines)
                lines = Lines()
            lines.extend(row)
            if end_of_para:
                res.append(lines)
                lines = Lines()
            start_of_sen = end_of_sen
            start_of_para = end_of_para = False
            prev_w, prev_text, prev_size = w, text, size
        if lines:
            res.append(lines)
        return res

    Lines.split_vertically_by_text = split_vertically_by_text

    # В оригинале пробел между соседними строками ставится только для ASCII (`c.encode().isalnum()`),
    # поэтому русские слова из разных блоков склеивались: «2026 г.Срок оплаты».
    from pdf2docx.text.TextSpan import TextSpan

    word_punct = set(".,;:!?)]}»”\"'%…№") | set("([{«„“")

    def wordish(c: str) -> bool:
        return bool(c) and not _is_cjk(c[-1]) and (c[-1].isalnum() or c[-1] in word_punct)

    def adjust_last_word(self, delete_end_line_hyphen: bool):
        for i, line in enumerate(self._instances[:-1]):
            end_span, start_span = line.spans[-1], self._instances[i + 1].spans[0]
            if not (isinstance(end_span, TextSpan) and isinstance(start_span, TextSpan)):
                continue
            if not end_span.chars or not start_span.chars:
                continue
            end_char, next_char = end_span.chars[-1], start_span.chars[0]
            if delete_end_line_hyphen and end_char.c.endswith("-") and next_char.c[:1].islower():
                end_char.c = ""
            nc = next_char.c[:1]
            if wordish(end_char.c) and wordish(nc) and not (nc in _NEVER_WORD_START and end_char.c[-1:].isalpha()):
                end_char.c += " "

    Lines.adjust_last_word = adjust_last_word


# --------------------------------------------------------------------------- ложные таблицы

def _iter_tables(blocks):
    for b in blocks:
        if getattr(b, "is_table_block", False):
            yield b
            for row in b:
                for cell in row:
                    if getattr(cell, "blocks", None):
                        yield from _iter_tables(cell.blocks)


def _page_tables(page):
    for section in page.sections:
        for column in section:
            yield from _iter_tables(column.blocks)


def _flat_text(value) -> str:
    """Текст блока: у вложенной таблицы pdf2docx отдаёт список строк, а не строку."""
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple)):
        return " ".join(_flat_text(v) for v in value if v is not None)
    return ""


def _cell_text(cell) -> str:
    # раньше вложенная таблица в ячейке роняла разбор всей страницы (TypeError в join)
    return " ".join(_flat_text(getattr(b, "text", "")) for b in (cell.blocks or [])).strip()


def _suspicious_stream_table(table) -> bool:
    """Безрамочная «таблица», которая на деле — разорванный по словам абзац."""
    rows = list(table)
    if not rows:
        return False
    total = placeholders = 0
    texts = []
    for row in rows:
        for cell in row:
            total += 1
            bb = cell.bbox
            if bb.is_empty or bb.width <= 0:
                placeholders += 1
            elif t := _cell_text(cell):
                texts.append(t)
    if len(rows) == 1:
        # одна строка из нескольких ячеек по 1–2 слова — это слова одной строки текста;
        # блок 1×1 (плашка с фоном) — нормальная таблица
        return len(texts) >= 2 and all(len(t.split()) <= 2 for t in texts)
    return total > 0 and placeholders / total > 0.4


def page_has_false_tables(page) -> bool:
    """Откат нужен, только если подозрительных безрамочных таблиц на странице большинство."""
    stream = [t for t in _page_tables(page) if t.is_stream_table_block]
    bad = sum(1 for t in stream if _suspicious_stream_table(t))
    return bad > 0 and bad >= 0.5 * len(stream)


class _PagePickler(pickle.Pickler):
    """Снимок страницы: ссылки секций на саму страницу остаются ссылками на неё, а не на копию."""

    def __init__(self, file, page):
        super().__init__(file, protocol=pickle.HIGHEST_PROTOCOL)
        self._page = page

    def persistent_id(self, obj):
        return "page" if obj is self._page else None


class _PageUnpickler(pickle.Unpickler):
    def __init__(self, file, page):
        super().__init__(file)
        self._page = page

    def persistent_load(self, pid):
        return self._page


def _snapshot(page):
    """Состояние страницы до разбора. pickle в 3–4 раза быстрее deepcopy, а распаковывается
    только при откате (редко); если что-то не сериализуется — deepcopy."""
    try:
        buf = io.BytesIO()
        _PagePickler(buf, page).dump(page.__dict__)
        return buf.getvalue()
    except Exception:
        try:
            return copy.deepcopy(page.__dict__, {id(page): page})
        except Exception:
            return None


def _restore(page, snapshot) -> None:
    state = _PageUnpickler(io.BytesIO(snapshot), page).load() if isinstance(snapshot, bytes) else snapshot
    page.__dict__.clear()
    page.__dict__.update(state)


def parse_page_safely(page, settings: dict) -> bool:
    """Разбирает страницу; при ложных безрамочных таблицах — повторно без них.
    Возвращает True, если понадобился повторный разбор."""
    if not settings.get("parse_stream_table", True):
        page.parse(**settings)
        return False
    backup = _snapshot(page)
    page.parse(**settings)
    if backup is None or not page_has_false_tables(page):
        return False
    _restore(page, backup)
    page.parse(**{**settings, "parse_stream_table": False})
    return True


# --------------------------------------------------------------------------- сжатие символов

# Шрифты, ширины которых совпадают со встроенными шрифтами MuPDF (Helvetica = Arial,
# Times = Times New Roman, Courier = Courier New; у MuPDF в них есть и кириллица).
_METRIC_FAMILIES = (
    (("arial", "helvetica", "liberationsans", "liberation sans", "arimo"), ("helv", "hebo", "heit", "hebi")),
    (("timesnewroman", "times new roman", "times", "liberationserif", "liberation serif", "tinos"),
     ("tiro", "tibo", "tiit", "tibi")),
    (("couriernew", "courier new", "courier", "liberationmono", "liberation mono", "cousine"),
     ("cour", "cobo", "coit", "cobi")),
)
_REF_FONTS: dict = {}
MIN_SCALE_CHARS = 4


def _reference_font(name: str, bold: bool, italic: bool):
    """Встроенный шрифт MuPDF с теми же ширинами, что у шрифта PDF, или None."""
    low = re.sub(r"^[A-Z]{6}\+", "", name or "").lower().replace("-", "").replace(",", "")
    if "narrow" in low or "condensed" in low:
        return None
    for keys, variants in _METRIC_FAMILIES:
        if any(k in low for k in keys):
            bold = bold or "bold" in low
            italic = italic or "italic" in low or "oblique" in low
            code = variants[(1 if bold else 0) + (2 if italic else 0)]
            if code not in _REF_FONTS:
                import pymupdf
                _REF_FONTS[code] = pymupdf.Font(code)
            return _REF_FONTS[code]
    return None


def horizontal_scale(chars, ref_font, font_size: float) -> float | None:
    """Во сколько раз буквы в PDF уже (меньше 1) или шире естественной ширины шрифта.

    Word при выгрузке в PDF записывает «масштаб» символов как горизонтальное сжатие глифов; pdf2docx
    его теряет, и в Word текст выходит шире, чем в PDF: коды в узких графах таблиц переносятся,
    строки растут, лист уходит на следующую страницу. Сравниваются ширины букв без пробелов, поэтому
    растянутые выравниванием по ширине пробелы на результат не влияют.
    """
    glyphs = [c for c in chars if c.c.strip()]
    if len(glyphs) < MIN_SCALE_CHARS:
        return None
    # только горизонтальный текст
    if glyphs[-1].bbox[0] <= glyphs[0].bbox[0]:
        return None
    observed = sum(c.bbox[2] - c.bbox[0] for c in glyphs)
    natural = ref_font.text_length("".join(c.c for c in glyphs), fontsize=font_size)
    if natural <= 0 or observed <= 0:
        return None
    return observed / natural


def _set_run_scale(run, percent: int) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    rpr = run._r.get_or_add_rPr()
    for old in rpr.findall(qn("w:w")):
        rpr.remove(old)
    if percent == 100:
        return
    el = OxmlElement("w:w")
    el.set(qn("w:val"), str(percent))
    rpr.insert_element_before(el, "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
                              "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
                              "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")


def _patch_char_scaling() -> None:
    from pdf2docx.text.TextSpan import TextSpan

    original = TextSpan._set_text_format

    def _set_text_format(self, docx_run):
        original(self, docx_run)
        try:
            font_size = round(self.size * 2) / 2.0 or self.size
            percent = None
            ref = _reference_font(self.font, bool(self.flags & 2 ** 4), bool(self.flags & 2 ** 1))
            if ref is not None and self.chars:
                scale = horizontal_scale(self.chars, ref, font_size)
                if scale is not None and abs(scale - 1.0) >= 0.03:
                    percent = int(round(min(max(scale, 0.5), 2.0) * 100))
            if percent is None:
                # поправка pdf2docx на округление кегля — целым числом (дробное Word/LibreOffice не читают)
                scale = self.size / font_size
                percent = int(round(scale * 100)) if abs(scale - 1.0) >= 0.01 else 100
            _set_run_scale(docx_run, percent)
        except Exception:
            log.debug("char scaling skipped", exc_info=True)

    TextSpan._set_text_format = _set_text_format


# --------------------------------------------------------------------------- красная строка в двух строках

def two_line_indent(x0_first: float, x0_second: float, x1_first: float, x1_block: float,
                    right_threshold: float) -> tuple[float, bool] | None:
    """(отступ первой строки, первая строка во всю ширину) для абзаца из двух строк или None.

    pdf2docx ищет отступ первой строки только у абзацев от трёх строк: у двухстрочного абзаца
    с красной строкой левые края расходятся, и он считает абзац «невыровненным» — отступ
    теряется (жалоба на опросные листы: «не сохраняется форматирование, отступ»)."""
    indent = x0_first - x0_second
    if not (2.0 <= abs(indent) <= 100.0):
        return None
    return indent, x1_first >= x1_block - max(right_threshold, 1.0)


def _patch_two_line_indent() -> None:
    from pdf2docx.common.share import TextAlignment
    from pdf2docx.text.TextBlock import TextBlock

    original = TextBlock._parse_alignment

    def _parse_alignment(self, bbox, text_direction_param, line_separate_threshold,
                         lines_left_aligned_threshold, lines_right_aligned_threshold,
                         lines_center_aligned_threshold):
        res = original(self, bbox, text_direction_param, line_separate_threshold, lines_left_aligned_threshold,
                       lines_right_aligned_threshold, lines_center_aligned_threshold)
        try:
            idx0, idx1, f = text_direction_param
            if res != TextAlignment.NONE or f != 1:
                return res
            rows = self.lines.group_by_physical_rows()
            if len(rows) != 2 or any(len(r) != 1 for r in rows):
                return res
            found = two_line_indent(rows[0][0].bbox[idx0], rows[1][0].bbox[idx0], rows[0][-1].bbox[idx1],
                                    self.bbox[idx1], lines_right_aligned_threshold)
            if found is None:
                return res
            indent, full = found
            self.first_line_space = indent
            return TextAlignment.JUSTIFY if full else TextAlignment.LEFT
        except Exception:
            log.debug("two-line indent skipped", exc_info=True)
            return res

    TextBlock._parse_alignment = _parse_alignment
