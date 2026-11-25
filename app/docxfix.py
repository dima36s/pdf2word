"""Структура документа Word после pdf2docx: заголовки, содержание.

* Нумерованные заголовки («2 Характеристика…», «2.1 …», «2.2.1 …») и стандартные разделы
  («Содержание», «Введение», «Приложение А») получают стили «Заголовок 1–4»: в Word появляется
  панель навигации и автоматическое оглавление. Вид не меняется — стили заголовков очищены от
  своего шрифта, цвета и интервалов, остаётся прямое форматирование из PDF и уровень структуры.
* Заголовок, который pdf2docx разложил в безрамочную таблицу «номер | текст», снова становится
  абзацем (номер, табуляция, текст с выступом).
* Строки содержания «Текст ........ 12»: отточие из точек заменяется табуляцией по правому краю
  с заполнителем — номера страниц встают ровно у края при любом шрифте.
"""
from __future__ import annotations

import copy
import re

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


NUM = r"\d{1,2}(?:\.\d{1,2}){0,3}"
HEADING_RE = re.compile(rf"^\s*({NUM})\.?\s+(\S.*)$", re.S)
NUMBER_ONLY_RE = re.compile(rf"^\s*({NUM})\.?\s*$")
TOC_RE = re.compile(r"^(?P<title>.*?\S)\s*(?:[.·…]\s*){4,}(?P<page>\d{1,4})\s*$", re.S)
TOC_TAB_RE = re.compile(r"^(?P<title>.*?\S)[ \u00a0]*\t+(?P<page>\d{1,4})\s*$", re.S)
STRUCT_HEADINGS = {"содержание", "введение", "заключение", "обозначения и сокращения", "термины и определения",
                   "список использованных источников", "библиография", "перечень сокращений",
                   "сокращения", "аннотация", "реферат"}


# --------------------------------------------------------------------------- текст и формат

def _text_nodes(p):
    """w:t и табуляции в тексте абзаца (w:tab внутри w:r; w:tab в w:tabs — это позиции, не текст)."""
    for node in p.iter(w("t"), w("tab")):
        if node.tag == w("tab") and (node.getparent() is None or node.getparent().tag != w("r")):
            continue
        yield node


def _ptext(p) -> str:
    parts = []
    for node in _text_nodes(p):
        parts.append("\t" if node.tag == w("tab") else (node.text or ""))
    return "".join(parts)


def _on(el) -> bool:
    return el is not None and el.get(w("val"), "true") not in ("0", "false", "off")


def _run_style(r) -> tuple[bool, float | None]:
    rpr = r.find(w("rPr"))
    if rpr is None:
        return False, None
    sz = rpr.find(w("sz"))
    size = int(sz.get(w("val"))) / 2 if sz is not None and (sz.get(w("val")) or "").isdigit() else None
    return _on(rpr.find(w("b"))), size


def _format(p) -> tuple[float, float | None]:
    """(доля жирного текста, максимальный кегль) абзаца."""
    total = bold = 0
    sizes = []
    for r in p.iter(w("r")):
        text = "".join(t.text or "" for t in r.iter(w("t")))
        n = len(text.strip())
        if not n:
            continue
        b, size = _run_style(r)
        total += n
        bold += n if b else 0
        if size:
            sizes.append(size)
    return (bold / total if total else 0.0), (max(sizes) if sizes else None)


def _body_paragraphs(body):
    return [el for el in body if el.tag == w("p")]


def _median_size(body) -> float | None:
    weights: dict[float, int] = {}
    for p in _body_paragraphs(body):
        for r in p.iter(w("r")):
            n = len("".join(t.text or "" for t in r.iter(w("t"))).strip())
            _, size = _run_style(r)
            if n and size:
                weights[size] = weights.get(size, 0) + n
    if not weights:
        return None
    total, acc = sum(weights.values()), 0
    for size in sorted(weights):
        acc += weights[size]
        if acc >= total / 2:
            return size
    return None


# --------------------------------------------------------------------------- таблица «номер | заголовок»

def _visible_borders(tbl) -> bool:
    for b in tbl.iter(w("tcBorders"), w("tblBorders")):
        for edge in b:
            if edge.get(w("val")) not in (None, "nil", "none") and edge.get(w("sz"), "1") != "0":
                return True
    return False


def _twips(el, attr: str) -> int:
    try:
        return int(float(el.get(w(attr), "0")))
    except (TypeError, ValueError):
        return 0


def heading_tables_to_paragraphs(body) -> int:
    converted = 0
    for tbl in [el for el in body if el.tag == w("tbl")]:
        if _visible_borders(tbl):
            continue
        rows = tbl.findall(w("tr"))
        if not rows or len(rows) > 6:
            continue
        plan = []
        for tr in rows:
            cells = tr.findall(w("tc"))
            texts = [_ptext(tc).strip() for tc in cells]
            filled = [(k, t) for k, t in enumerate(texts) if t]
            if len(filled) != 2 or not NUMBER_ONLY_RE.match(filled[0][1]) or len(filled[1][1]) > 300:
                plan = None
                break
            plan.append((cells[filled[0][0]], cells[filled[1][0]], filled[0][1]))
        if not plan:
            continue
        grid = [_twips(g, "w") for g in tbl.iter(w("gridCol"))]
        tblpr = tbl.find(w("tblPr"))
        ind = tblpr.find(w("tblInd")) if tblpr is not None else None
        left0 = _twips(ind, "w") if ind is not None else 0
        num_w = grid[0] if grid else 567
        new_ps = []
        for num_cell, text_cell, number in plan:
            paras = [copy.deepcopy(p) for p in text_cell.findall(w("p")) if _ptext(p).strip()]
            if not paras:
                continue
            first = paras[0]
            src_run = next(num_cell.iter(w("r")), None)
            run = etree.Element(w("r"))
            if src_run is not None and src_run.find(w("rPr")) is not None:
                run.append(copy.deepcopy(src_run.find(w("rPr"))))
            t = etree.SubElement(run, w("t"))
            t.text = number
            etree.SubElement(run, w("tab"))
            ppr = first.find(w("pPr"))
            first.insert(0 if ppr is None else 1, run)
            for p in paras:
                _set_indent(p, left0 + num_w, num_w if p is first else 0, tab=left0 + num_w if p is first else None)
            new_ps.extend(paras)
        if not new_ps:
            continue
        for p in new_ps:
            tbl.addprevious(p)
        body.remove(tbl)
        converted += 1
    return converted


_PPR_AFTER_TABS = ("w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                   "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
                   "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                   "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
                   "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
_PPR_AFTER_IND = _PPR_AFTER_TABS[_PPR_AFTER_TABS.index("w:ind") + 1:]


def _ppr(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    return ppr


def _insert(ppr, el, successors) -> None:
    for tag in successors:
        prefix, local = tag.split(":")
        nxt = ppr.find(w(local))
        if nxt is not None:
            nxt.addprevious(el)
            return
    ppr.append(el)


def _set_indent(p, left: int, hanging: int, tab: int | None) -> None:
    ppr = _ppr(p)
    ind = ppr.find(w("ind"))
    if ind is None:
        ind = etree.Element(w("ind"))
        _insert(ppr, ind, _PPR_AFTER_IND)
    for attr in ("left", "start", "hanging", "firstLine", "firstLineChars", "hangingChars"):
        ind.attrib.pop(w(attr), None)
    ind.set(w("left"), str(left))
    if hanging:
        ind.set(w("hanging"), str(hanging))
    if tab is not None:
        _set_tabs(p, [("left", tab, None)])


def _set_tabs(p, stops) -> None:
    ppr = _ppr(p)
    old = ppr.find(w("tabs"))
    if old is not None:
        ppr.remove(old)
    tabs = etree.Element(w("tabs"))
    for val, pos, leader in stops:
        t = etree.SubElement(tabs, w("tab"))
        t.set(w("val"), val)
        if leader:
            t.set(w("leader"), leader)
        t.set(w("pos"), str(int(pos)))
    _insert(ppr, tabs, _PPR_AFTER_TABS)


# --------------------------------------------------------------------------- сетка таблиц

def _cell_span(tc) -> int:
    tcpr = tc.find(w("tcPr"))
    gs = tcpr.find(w("gridSpan")) if tcpr is not None else None
    try:
        return max(1, int(gs.get(w("val")))) if gs is not None else 1
    except (TypeError, ValueError):
        return 1


def _cell_width(tc) -> int | None:
    tcpr = tc.find(w("tcPr"))
    tcw = tcpr.find(w("tcW")) if tcpr is not None else None
    if tcw is None or tcw.get(w("type"), "dxa") != "dxa":
        return None
    try:
        return int(float(tcw.get(w("w"))))
    except (TypeError, ValueError):
        return None


def fix_table_grid(tbl) -> bool:
    """Ширины столбцов сетки (w:gridCol) по ширинам ячеек.

    pdf2docx задаёт ширину каждой ячейки, а сетку оставляет равномерной. Word смотрит на ячейки,
    LibreOffice и Р7-Офис — на сетку: таблица у них «расползается» в одинаковые столбцы.
    """
    grid = tbl.find(w("tblGrid"))
    if grid is None:
        return False
    cols = grid.findall(w("gridCol"))
    n = len(cols)
    # исходная сетка python-docx равномерная — её ширина и есть «заглушка» у ячеек, которым
    # pdf2docx ширину не задал (пустые части объединений)
    original = {_twips(c, "w") for c in cols}
    placeholder = original.pop() if len(original) == 1 else None

    def is_placeholder(width: int, span: int) -> bool:
        return placeholder is not None and abs(width - placeholder * span) <= 2 * span

    votes: list[dict[int, int]] = [{} for _ in range(n)]
    spans = []
    for tr in tbl.findall(w("tr")):
        c = 0
        for tc in tr.findall(w("tc")):
            span, width = _cell_span(tc), _cell_width(tc)
            if c + span > n:
                break
            tcpr = tc.find(w("tcPr"))
            vm = tcpr.find(w("vMerge")) if tcpr is not None else None
            # продолжение вертикального объединения: python-docx оставляет там ширину-заглушку
            continuation = vm is not None and vm.get(w("val"), "continue") != "restart"
            if width is not None and not continuation and not is_placeholder(width, span):
                if span == 1:
                    votes[c][width] = votes[c].get(width, 0) + 1
                else:
                    spans.append((c, span, width))
            c += span
    known: list[int | None] = [max(v.items(), key=lambda kv: kv[1])[0] if v else None for v in votes]
    # у объединённых ячеек python-docx оставляет сумму ширин-заглушек: для каждой группы
    # столбцов берём самую частую ширину (при равенстве — меньшую)
    groups: dict[tuple[int, int], dict[int, int]] = {}
    for c, span, width in spans:
        groups.setdefault((c, span), {}).setdefault(width, 0)
        groups[(c, span)][width] += 1
    spans = [(c, span, max(ws.items(), key=lambda kv: (kv[1], -kv[0]))[0]) for (c, span), ws in groups.items()]
    spans.sort(key=lambda t: t[1])
    progress = True
    while progress:
        progress = False
        for c, span, width in spans:
            missing = [k for k in range(c, c + span) if known[k] is None]
            if len(missing) == 1:
                rest = width - sum(known[k] for k in range(c, c + span) if known[k] is not None)
                known[missing[0]] = max(rest, 10)
                progress = True
    for c, span, width in spans:      # что осталось — поровну
        missing = [k for k in range(c, c + span) if known[k] is None]
        if missing:
            rest = width - sum(known[k] for k in range(c, c + span) if known[k] is not None)
            for k in missing:
                known[k] = max(rest // len(missing), 10)
    if any(v is None for v in known):
        return False
    for col, value in zip(cols, known):
        col.set(w("w"), str(int(value)))
    # pdf2docx пишет отступ таблицы дробным числом («11906.0») — строгие редакторы его игнорируют
    for ind in tbl.iter(w("tblInd")):
        try:
            ind.set(w("w"), str(int(round(float(ind.get(w("w"), "0"))))))
        except ValueError:
            pass
    return True


def dedupe_cell_margins(root) -> int:
    """pdf2docx при объединении ячеек добавляет w:tcMar на каждую исходную ячейку — по схеме он один,
    и Word может счесть такой файл повреждённым."""
    n = 0
    for tcpr in root.iter(w("tcPr")):
        mars = tcpr.findall(w("tcMar"))
        for extra in mars[1:]:
            tcpr.remove(extra)
            n += 1
    return n


def fix_table_grids(root) -> int:
    dedupe_cell_margins(root)
    n = sum(1 for tbl in root.iter(w("tbl")) if fix_table_grid(tbl))
    for tbl in root.iter(w("tbl")):
        drop_sliver_rows(tbl)
        fix_empty_cells(tbl)
    return n


SLIVER_TWIPS = 100   # 5 pt


def drop_sliver_rows(tbl) -> int:
    """Пустые строки-«щели» в несколько пунктов (между последней линией сетки и границей штампа).

    В Word строка не бывает ниже строки текста пустого абзаца, поэтому щель в 2 pt превращалась
    в строку ~10 pt и выталкивала конец таблицы на новую страницу.
    """
    rows = tbl.findall(w("tr"))
    removed = 0
    for tr in rows:
        if len(rows) - removed <= 1:
            break
        h = tr.find(f"{w('trPr')}/{w('trHeight')}")
        try:
            height = int(float(h.get(w("val")))) if h is not None else None
        except (TypeError, ValueError):
            height = None
        if height is None or height >= SLIVER_TWIPS:
            continue
        if any((t.text or "").strip() for t in tr.iter(w("t"))):
            continue
        if any(next(tr.iter(w(tag)), None) is not None for tag in ("drawing", "pict", "object")):
            continue
        if any(vm.get(w("val")) == "restart" for vm in tr.iter(w("vMerge"))):
            continue
        tbl.remove(tr)
        removed += 1
    return removed


def fix_empty_cells(tbl) -> int:
    """Пустые ячейки: знак абзаца получает кегль, преобладающий в таблице.

    Иначе пустой абзац берёт шрифт 11 pt по умолчанию и раздвигает строку высотой «не менее»
    (бланк спецификации с десятками пустых строк уходил на лишнюю страницу), а вписанный
    пользователем текст получал бы чужой размер.
    """
    sizes: dict[int, int] = {}
    for r in tbl.iter(w("r")):
        n = len("".join(t.text or "" for t in r.iter(w("t"))).strip())
        rpr = r.find(w("rPr"))
        sz = rpr.find(w("sz")) if rpr is not None else None
        if n and sz is not None and (sz.get(w("val")) or "").isdigit():
            sizes[int(sz.get(w("val")))] = sizes.get(int(sz.get(w("val"))), 0) + n
    if not sizes:
        return 0
    half_points = max(sizes.items(), key=lambda kv: kv[1])[0]
    fixed = 0
    for tc in tbl.iter(w("tc")):
        for p in tc.findall(w("p")):
            if p.find(w("r")) is not None:
                continue
            ppr = _ppr(p)
            rpr = ppr.find(w("rPr"))
            if rpr is None:
                rpr = etree.Element(w("rPr"))
                _insert(ppr, rpr, ("w:sectPr", "w:pPrChange"))
            for tag in ("sz", "szCs"):
                el = rpr.find(w(tag))
                if el is None:
                    el = etree.SubElement(rpr, w(tag))
                el.set(w("val"), str(half_points))
            fixed += 1
    return fixed


def fix_default_spacing(docx) -> None:
    """Интервалы по умолчанию — без отступа после абзаца и одинарные.

    Шаблон python-docx задаёт 10 pt после абзаца и множитель 1,15. pdf2docx прописывает интервалы
    своим абзацам явно, так что умолчание действует только на служебные пустые абзацы — прежде
    всего в пустых ячейках таблиц, где оно раздувало строки.
    """
    styles = docx.styles.element
    defaults = styles.find(w("docDefaults"))
    if defaults is None:
        return
    ppr_default = defaults.find(w("pPrDefault"))
    if ppr_default is None:
        ppr_default = etree.SubElement(defaults, w("pPrDefault"))
    ppr = ppr_default.find(w("pPr"))
    if ppr is None:
        ppr = etree.SubElement(ppr_default, w("pPr"))
    sp = ppr.find(w("spacing"))
    if sp is None:
        sp = etree.SubElement(ppr, w("spacing"))
    sp.set(w("after"), "0")
    sp.set(w("line"), "240")
    sp.set(w("lineRule"), "auto")


# --------------------------------------------------------------------------- заголовки

def heading_level(raw: str, bold_share: float, size: float | None, median: float | None) -> int | None:
    if TOC_RE.match(raw) or TOC_TAB_RE.match(raw):
        return None                      # строка содержания, а не заголовок
    text = " ".join(raw.split())
    if not text or len(text) > 250:
        return None
    emphasized = bold_share >= 0.8 or (size is not None and median is not None and size >= median + 1.5)
    if not emphasized:
        return None
    m = HEADING_RE.match(text)
    if m:
        rest = m.group(2)
        if rest[-1:] in ",;:" or not rest[:1].isalpha():
            return None
        return min(4, m.group(1).count(".") + 1)
    low = text.lower().rstrip(".")
    if low in STRUCT_HEADINGS or re.match(r"^приложение\s+[а-яa-z0-9]{1,2}\b", low):
        return 1
    return None


def _neutral_heading_styles(docx, levels) -> dict[int, str]:
    ids = {}
    for lvl in levels:
        try:
            st = docx.styles[f"Heading {lvl}"]
        except KeyError:
            continue
        ids[lvl] = st.style_id
        el = st.element
        rpr = el.find(w("rPr"))
        if rpr is not None:
            el.remove(rpr)
        ppr = el.find(w("pPr"))
        if ppr is not None:
            for tag in ("keepNext", "keepLines", "spacing", "ind", "jc", "pageBreakBefore"):
                node = ppr.find(w(tag))
                if node is not None:
                    ppr.remove(node)
    return ids


def mark_headings(docx) -> int:
    body = docx.element.body
    median = _median_size(body)
    found = []
    for p in _body_paragraphs(body):
        ppr = p.find(w("pPr"))
        if ppr is not None and ppr.find(w("pStyle")) is not None:
            continue
        share, size = _format(p)
        lvl = heading_level(_ptext(p), share, size, median)
        if lvl:
            found.append((p, lvl))
    if not found:
        return 0
    ids = _neutral_heading_styles(docx, sorted({lvl for _, lvl in found}))
    n = 0
    for p, lvl in found:
        if lvl not in ids:
            continue
        ppr = _ppr(p)
        ps = etree.Element(w("pStyle"))
        ps.set(w("val"), ids[lvl])
        ppr.insert(0, ps)
        n += 1
    return n


# --------------------------------------------------------------------------- содержание

def _section_widths(docx) -> list[int]:
    """Ширина текста каждого раздела, твипы."""
    out = []
    for s in docx.sections:
        try:
            out.append(int((s.page_width - s.left_margin - s.right_margin) / 635))   # EMU → твипы
        except Exception:
            out.append(9355)
    return out


def _replace_leader(p, start: int, end: int) -> None:
    """Удалить символы [start, end) текста абзаца и поставить на их место табуляцию."""
    pos = 0
    tab_done = False
    for node in list(_text_nodes(p)):
        text = "\t" if node.tag == w("tab") else (node.text or "")
        a, b = pos, pos + len(text)
        pos = b
        if b <= start or a >= end:
            continue
        lo, hi = max(start, a) - a, min(end, b) - a
        if node.tag == w("tab"):
            node.getparent().remove(node)
            continue
        before, after = text[:lo], text[hi:]
        node.text = before
        if before.endswith(" ") or after.startswith(" "):
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        anchor = node
        if not tab_done:
            tab = etree.Element(w("tab"))
            anchor.addnext(tab)
            anchor = tab
            tab_done = True
        if after:
            t2 = etree.Element(w("t"))
            t2.text = after
            anchor.addnext(t2)
        if not before:
            node.getparent().remove(node)


def fix_toc(docx) -> int:
    body = docx.element.body
    widths = _section_widths(docx)
    sec = 0
    candidates = []
    for el in body:
        if el.tag == w("p"):
            text = _ptext(el)
            m = TOC_RE.match(text)
            kind = "dots" if m else None
            if not m:
                m = TOC_TAB_RE.match(text)
                kind = "tab" if m else None
            candidates.append((el, m, kind, sec))
            ppr = el.find(w("pPr"))
            if ppr is not None and ppr.find(w("sectPr")) is not None:
                sec += 1
        else:
            candidates.append((el, None, None, sec))
    done = []
    leaders: dict[int, str | None] = {}
    for k, (p, m, kind, s) in enumerate(candidates):
        if m is None:
            continue
        if kind == "tab":
            # одиночная строка «текст<TAB>число» — не оглавление; нужно хотя бы три подряд
            near = [c for c in candidates[max(0, k - 2):k + 3] if c[1] is not None]
            if len(near) < 3:
                continue
        _replace_leader(p, len(m.group("title")), m.start("page"))
        done.append((k, p, s))
        leaders[id(p)] = "dot" if kind == "dots" else None      # без отточия в PDF — без него и в Word
    for group in _toc_groups(done):
        s = group[0][2]
        width = widths[min(s, len(widths) - 1)] if widths else 9355
        right = toc_right_indent([_twips(_ppr(p).find(w("ind")), "right") if _ppr(p).find(w("ind")) is not None
                                  else 0 for _, p, _ in group], width)
        labels = _align_toc_labels([p for _, p, _ in group])
        for _, p, _ in group:
            stop = max(0, width - right)
            stops = [("right", stop, leaders.get(id(p), "dot"))]
            if p in labels:
                stops.insert(0, ("left", labels[p], None))
            else:
                # левые табуляции pdf2docx (номер главы картинкой, текст пункта с выступом) остаются:
                # без них табуляция после номера уходила к номеру страницы — отточие перед текстом
                lead = _leading_tabs(p)
                stops = [("left", pos, None) for pos in _left_stops(p) if pos < stop - 300][:lead] + stops
            if right:
                _set_ind(p, right=right)
            _set_tabs(p, stops)
    return len(done)


def _leading_tabs(p) -> int:
    """Сколько табуляций стоит до текста пункта (после номера-картинки, номера главы)."""
    n = 0
    for el in p.iter(w("t"), w("tab")):
        if el.tag == w("tab"):
            if el.getparent().tag == w("tabs"):
                continue
            n += 1
        elif (el.text or "").strip():
            break
    return n


def _left_stops(p) -> list[int]:
    tabs = _ppr(p).find(w("tabs"))
    out = []
    for t in tabs if tabs is not None else ():
        if t.get(w("val")) == "left":
            try:
                out.append(int(float(t.get(w("pos")))))
            except (TypeError, ValueError):
                pass
    return sorted(out)


TOC_LABEL_RE = re.compile(r"^(?P<label>\d{1,3}(?:\.\d{1,3}){0,5}\.?)(?P<gap>[ \u00a0]+)\S")


def _toc_groups(done: list) -> list[list]:
    """Подряд идущие пункты оглавления одного раздела (между ними не больше двух других абзацев)."""
    groups: list[list] = []
    for item in done:
        if groups and item[2] == groups[-1][-1][2] and item[0] - groups[-1][-1][0] <= 3:
            groups[-1].append(item)
        else:
            groups.append([item])
    return groups


def toc_right_indent(rights: list[int], width: int) -> int:
    """Общий правый край пунктов оглавления.

    pdf2docx даёт правый отступ только многострочным пунктам (у однострочного он 0), поэтому номера
    страниц у одних пунктов стояли у края листа, у других — на 2–3 см левее. Край у всего оглавления
    один — берём наибольший разумный отступ."""
    usable = [r for r in rights if 0 < r < width * 0.4]
    return max(usable) if usable else 0


def _set_ind(p, **values) -> None:
    ppr = _ppr(p)
    ind = ppr.find(w("ind"))
    if ind is None:
        ind = etree.Element(w("ind"))
        _insert(ppr, ind, _PPR_AFTER_TABS[_PPR_AFTER_TABS.index("w:ind") + 1:])
    for key, val in values.items():
        if val is None:
            ind.attrib.pop(w(key), None)
        else:
            ind.set(w(key), str(int(val)))


def _align_toc_labels(paragraphs: list) -> dict:
    """«2.3 Отделение …»: номер пункта и текст — как в оглавлении Word, через табуляцию.

    Положение текста берётся у многострочного пункта того же уровня (у него pdf2docx определил
    выступ первой строки), однострочным пунктам ставится тот же выступ. Возвращает {абзац: позиция
    табуляции после номера, твипы}."""
    info = []
    for p in paragraphs:
        m = TOC_LABEL_RE.match(_ptext(p))
        if not m:
            continue
        ind = _ppr(p).find(w("ind"))
        left = _twips(ind, "left") if ind is not None else 0
        hang = _twips(ind, "hanging") if ind is not None else 0
        first = _twips(ind, "firstLine") if ind is not None else 0
        depth = m.group("label").rstrip(".").count(".") + 1
        info.append((p, m, depth, left - hang + first, left, hang))
    exemplars: dict[int, tuple[int, int]] = {}
    for p, m, depth, label_x, left, hang in info:
        if hang > 0:
            exemplars.setdefault(depth, (label_x, left))
    out = {}
    for p, m, depth, label_x, left, hang in info:
        if depth not in exemplars:
            continue
        ex_label, ex_text = exemplars[depth]
        if abs(label_x - ex_label) > 60 or ex_text <= ex_label:
            continue
        _set_ind(p, left=ex_text, hanging=ex_text - ex_label, firstLine=None)
        start = len(m.group("label"))
        _replace_leader(p, start, start + len(m.group("gap")))
        out[p] = ex_text
    return out


def shrink_section_breaks(body) -> int:
    """Пустой абзац с разрывом раздела после таблицы — высотой 1 pt.

    Когда таблица заполняет лист до штампа, служебный абзац высотой в строку не помещается и
    Word/LibreOffice начинают ради него новую страницу — после каждого листа спецификации
    появлялся пустой лист с рамкой и штампом.
    """
    n = 0
    children = list(body)
    page_ends = _page_end_breaks(children)
    for k, el in enumerate(children):
        if el.tag != w("p") or el.find(w("r")) is not None:
            continue
        ppr = el.find(w("pPr"))
        breaks = ppr is not None and ppr.find(w("sectPr")) is not None
        last = k + 1 < len(children) and children[k + 1].tag == w("sectPr")
        after_table = k > 0 and children[k - 1].tag == w("tbl")
        if not (breaks or last):
            continue
        if not (after_table or _is_tiny_neighbour(children, k) or k in page_ends or last):
            continue
        ppr = _ppr(el)
        sp = ppr.find(w("spacing"))
        if sp is None:
            sp = etree.Element(w("spacing"))
            _insert(ppr, sp, _PPR_AFTER_TABS[_PPR_AFTER_TABS.index("w:spacing") + 1:])
        for k2, v in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
            sp.set(w(k2), v)
        rpr = ppr.find(w("rPr"))
        if rpr is None:
            rpr = etree.Element(w("rPr"))
            _insert(ppr, rpr, ("w:sectPr", "w:pPrChange"))
        sz = rpr.find(w("sz"))
        if sz is None:
            sz = etree.SubElement(rpr, w("sz"))
        sz.set(w("val"), "2")
        n += 1
    return n


def _page_end_breaks(children) -> set[int]:
    """Индексы абзацев-разрывов, за которыми следующий раздел начинается с новой страницы.

    Такой абзац — последний на листе, после него ничего не видно; его высота в строку только
    выталкивает на лишнюю страницу плотно заполненный лист (штамп, нижняя таблица, имя файла).
    """
    marks = []
    for k, el in enumerate(children):
        if el.tag == w("sectPr"):
            marks.append((k, el))
        elif el.tag == w("p"):
            sp = el.find(f"{w('pPr')}/{w('sectPr')}")
            if sp is not None:
                marks.append((k, sp))
    out = set()
    for (k, _), (_, nxt) in zip(marks, marks[1:]):
        kind = nxt.find(w("type"))
        if kind is None or kind.get(w("val"), "nextPage") != "continuous":
            out.add(k)
    return out


def _is_tiny_neighbour(children, k: int) -> bool:
    """Перед абзацем — служебный абзац pdf2docx «после таблицы» (точная высота меньше 2 pt)."""
    if k == 0 or children[k - 1].tag != w("p") or children[k - 1].find(w("r")) is not None:
        return False
    sp = children[k - 1].find(f"{w('pPr')}/{w('spacing')}")
    try:
        return sp is not None and sp.get(w("lineRule")) == "exact" and int(sp.get(w("line"), "999")) <= 40
    except ValueError:
        return False


# --------------------------------------------------------------------------- один лист PDF — одна страница

LINE_FACTOR = 1.15          # высота строки одинарного интервала относительно кегля (оценка)
FIT_SAFETY = 6.0            # pt запаса на лист (проверено на листах A3 спецификаций в LibreOffice)
FIT_SAFETY_PER_ROW = 0.5    # pt на строку таблицы (толщина линий)


def _twip_attr(el, attr: str) -> float:
    if el is None:
        return 0.0
    try:
        return float(el.get(w(attr), "0")) / 20.0
    except (TypeError, ValueError):
        return 0.0


def _para_size(p) -> float:
    sizes = []
    for sz in p.iter(w("sz")):
        try:
            sizes.append(int(sz.get(w("val"))) / 2.0)
        except (TypeError, ValueError):
            pass
    return max(sizes) if sizes else 11.0


CHAR_WIDTH = 0.52           # средняя ширина буквы относительно кегля (Arial/Times, кириллица)


def _text_width(p) -> float:
    """Оценка ширины текста абзаца в одну строку, pt."""
    total = 0.0
    for r in p.iter(w("r")):
        text = "".join(t.text or "" for t in r.iter(w("t")))
        if not text:
            continue
        rpr = r.find(w("rPr"))
        size, scale = 11.0, 1.0
        if rpr is not None:
            sz, sc = rpr.find(w("sz")), rpr.find(w("w"))
            try:
                size = int(sz.get(w("val"))) / 2.0 if sz is not None else size
                scale = float(sc.get(w("val"))) / 100.0 if sc is not None else scale
            except (TypeError, ValueError):
                pass
        total += len(text) * size * CHAR_WIDTH * scale
    return total


def paragraph_height(p, width: float | None = None) -> float:
    """Оценка высоты абзаца, pt. width — ширина, доступная абзацу (текст переносится по ней)."""
    sp = p.find(f"{w('pPr')}/{w('spacing')}")
    before, after = _twip_attr(sp, "before"), _twip_attr(sp, "after")
    lines = 1 + sum(1 for br in p.iter(w("br")) if br.get(w("type")) in (None, "textWrapping"))
    if width:
        ind = p.find(f"{w('pPr')}/{w('ind')}")
        room = width - _twip_attr(ind, "left") - _twip_attr(ind, "right") - max(_twip_attr(ind, "firstLine"), 0.0)
        tw = _text_width(p)
        if tw > 0 and room > 20:
            lines = max(lines, int(tw // room) + 1)
    rule = sp.get(w("lineRule")) if sp is not None else None
    # картинка в строке (логотип) задаёт высоту строки, если интервал не точный
    pictures = 0.0
    for ext in p.iter(f"{{{WP}}}extent"):
        if ext.getparent() is not None and ext.getparent().tag == f"{{{WP}}}inline":
            try:
                pictures = max(pictures, int(ext.get("cy", "0")) / 12700.0)
            except ValueError:
                pass
    if rule == "exact":
        line = _twip_attr(sp, "line")
    elif rule == "atLeast":
        line = max(_twip_attr(sp, "line"), _para_size(p) * LINE_FACTOR)
    else:
        mult = (_twip_attr(sp, "line") * 20 / 240.0) if sp is not None and sp.get(w("line")) else 1.0
        line = _para_size(p) * LINE_FACTOR * mult
    if rule != "exact" and pictures:
        line = max(line, pictures)
    return before + lines * line + after


def _row_height(tr) -> tuple[float, float]:
    """(минимальная высота строки, оценка высоты содержимого), pt."""
    h = tr.find(f"{w('trPr')}/{w('trHeight')}")
    min_h = _twip_attr(h, "val")
    content = 0.0
    for tc in tr.findall(w("tc")):
        vm = tc.find(f"{w('tcPr')}/{w('vMerge')}")
        if vm is not None and vm.get(w("val"), "continue") != "restart":
            continue
        cw = _twip_attr(tc.find(f"{w('tcPr')}/{w('tcW')}"), "w") or None
        content = max(content, sum(paragraph_height(p, cw) for p in tc.findall(w("p"))))
    if h is not None and h.get(w("hRule")) == "exact":
        return min_h, min_h
    return min_h, content


def _thin_filler_rows(tbl) -> None:
    """Пустая строка-заполнитель ниже строки шрифта: пустому абзацу — точный интервал по строке,
    иначе строка вырастет до высоты строки шрифта (и лист не поместится)."""
    for tr in tbl.findall(w("tr")):
        min_h, content = _row_height(tr)
        if min_h <= 0 or content <= min_h + 0.5:
            continue
        if any((t.text or "").strip() for t in tr.iter(w("t"))):
            continue
        for tc in tr.findall(w("tc")):
            for p in tc.findall(w("p")):
                if p.find(w("r")) is not None:
                    continue
                ppr = _ppr(p)
                sp = ppr.find(w("spacing"))
                if sp is None:
                    sp = etree.Element(w("spacing"))
                    _insert(ppr, sp, _PPR_AFTER_TABS[_PPR_AFTER_TABS.index("w:spacing") + 1:])
                n = max(1, len(tc.findall(w("p"))))
                for k, v in (("before", "0"), ("after", "0"),
                             ("line", str(max(20, int((min_h - 0.5) * 20 / n)))), ("lineRule", "exact")):
                    sp.set(w(k), v)


def story_height(root, width: float | None = None) -> float:
    """Оценка высоты колонтитула, pt: «плавающие» таблицы и рисунки места в потоке не занимают."""
    total = 0.0
    for el in root:
        if el.tag == w("p"):
            total += paragraph_height(el, width)
        elif el.tag == w("tbl"):
            if el.find(f"{w('tblPr')}/{w('tblpPr')}") is not None:
                continue
            total += sum(max(*_row_height(tr)) for tr in el.findall(w("tr")))
    return total


def _section_groups(body):
    groups, cur = [], []
    for el in body:
        if el.tag == w("sectPr"):
            break
        cur.append(el)
        if el.tag == w("p") and el.find(f"{w('pPr')}/{w('sectPr')}") is not None:
            groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)
    return groups


def _spacer(p) -> bool:
    if p.tag != w("p") or p.find(w("r")) is not None or p.find(f"{w('pPr')}/{w('sectPr')}") is not None:
        return False
    sp = p.find(f"{w('pPr')}/{w('spacing')}")
    return sp is not None and sp.get(w("lineRule")) == "exact" and _twip_attr(sp, "line") >= 4.0


def _sect_layout(sect) -> tuple[str, int]:
    """(тип начала раздела, число колонок) по w:sectPr."""
    if sect is None:
        return "nextPage", 1
    t = sect.find(w("type"))
    cols = sect.find(w("cols"))
    try:
        num = int(cols.get(w("num"), "1")) if cols is not None else 1
    except ValueError:
        num = 1
    return (t.get(w("val"), "nextPage") if t is not None else "nextPage"), max(1, num)


def _column_chunks(elems: list) -> list[list]:
    """Элементы раздела по колонкам: абзац с разрывом колонки начинает следующую."""
    chunks: list[list] = [[]]
    for el in elems:
        if el.tag == w("p") and any(br.get(w("type")) == "column" for br in el.iter(w("br"))):
            chunks.append([])
        chunks[-1].append(el)
    return [c for c in chunks if c]


def _el_height(el, width: float | None) -> tuple[float, int]:
    """(высота, число строк таблицы) элемента тела."""
    if el.tag == w("p"):
        return paragraph_height(el, width), 0
    if el.tag == w("tbl"):
        if el.find(f"{w('tblPr')}/{w('tblpPr')}") is not None:
            return 0.0, 0      # плавающая таблица места в потоке не занимает
        rows = el.findall(w("tr"))
        return sum(max(*_row_height(tr)) for tr in rows), len(rows)
    return 0.0, 0


def _chunk_height(elems, width) -> tuple[float, int]:
    total, rows = 0.0, 0
    for el in elems:
        h, r = _el_height(el, width)
        total += h
        rows += r
    return total, rows


class _PagePart:
    """Раздел Word в пределах листа PDF: одна колонка или несколько (высота — по самой высокой)."""

    def __init__(self, elems: list, cols: int, width: float):
        self.elems, self.cols = elems, cols
        self.width = width / cols if cols > 1 else width
        self.chunks = _column_chunks(elems) if cols > 1 else [elems]

    def height(self) -> tuple[float, int]:
        sizes = [_chunk_height(c, self.width) for c in self.chunks]
        return max(sizes, key=lambda hr: hr[0]) if sizes else (0.0, 0)


def _page_groups(docx) -> list[tuple[object, list[_PagePart]]]:
    """Листы PDF: раздел «с новой страницы» и следующие за ним непрерывные разделы
    (pdf2docx делит лист на разделы по колонкам)."""
    body = docx.element.body
    groups = _section_groups(body)
    pages: list[tuple[object, list[_PagePart]]] = []
    for section, elems in zip(docx.sections, groups):
        kind, cols = _sect_layout(section._sectPr)
        try:
            width = (section.page_width - section.left_margin - section.right_margin) / 12700.0
        except Exception:
            width = None
        part = _PagePart(elems, cols, width or 0.0)
        if kind == "continuous" and pages:
            pages[-1][1].append(part)
        else:
            pages.append((section, [part]))
    return pages


def _reclaim(elems: list, amount: float) -> float:
    """Забирает до amount pt у отступов между блоками; возвращает, сколько забрано."""
    excess = amount
    # 1) пустые абзацы-отступы между блоками
    for p in sorted((el for el in elems if _spacer(el)), key=paragraph_height, reverse=True):
        if excess <= 0:
            break
        sp = p.find(f"{w('pPr')}/{w('spacing')}")
        line = _twip_attr(sp, "line")
        take = min(excess, line * 0.8)
        sp.set(w("line"), str(max(20, int(round((line - take) * 20)))))
        excess -= take
    # 2) большие интервалы «после» у абзацев: ими pdf2docx воспроизводит промежуток до следующего блока
    for p in sorted((el for el in elems if el.tag == w("p")),
                    key=lambda el: _twip_attr(el.find(f"{w('pPr')}/{w('spacing')}"), "after"), reverse=True):
        if excess <= 0:
            break
        sp = p.find(f"{w('pPr')}/{w('spacing')}")
        after = _twip_attr(sp, "after")
        if after < 6.0:
            break
        take = min(excess, after * 0.8)
        sp.set(w("after"), str(int(round((after - take) * 20))))
        excess -= take
    # 3) большие интервалы «перед» абзацами: титульные листы, разрядка блоков
    for p in sorted((el for el in elems if el.tag == w("p")),
                    key=lambda el: _twip_attr(el.find(f"{w('pPr')}/{w('spacing')}"), "before"), reverse=True):
        if excess <= 0:
            break
        sp = p.find(f"{w('pPr')}/{w('spacing')}")
        before = _twip_attr(sp, "before")
        if before < 12.0:
            break
        take = min(excess, before * 0.5)
        sp.set(w("before"), str(int(round((before - take) * 20))))
        excess -= take
    return amount - excess


def _shrink_last_rows(tbl, amount: float) -> float:
    """4) нижние строки таблицы (у строки остаётся её содержимое)."""
    excess = amount
    for tr in reversed(tbl.findall(w("tr"))):
        if excess <= 0:
            break
        h = tr.find(f"{w('trPr')}/{w('trHeight')}")
        if h is None or h.get(w("hRule")) == "exact":
            continue
        min_h, content = _row_height(tr)
        room = min_h - max(content, min_h * 0.6)
        if room <= 0.5:
            continue
        take = min(excess, room)
        h.set(w("val"), str(int(round((min_h - take) * 20))))
        excess -= take
    return amount - excess


def fit_pages(docx) -> int:
    """Лист PDF не должен растягиваться на две страницы Word.

    pdf2docx воспроизводит положение блоков точными отступами, поэтому лист, заполненный до
    штампа, оказывается в Word «впритык»: толщина линий таблиц и округления выталкивают хвост на
    новую страницу (пустую или с обрывком таблицы). Если оценка высоты листа не оставляет запаса,
    недостающие пункты забираются у пустых абзацев-отступов между блоками, затем у нижней строки
    последней таблицы. Лист — это раздел «с новой страницы» вместе с непрерывными разделами за ним;
    у раздела в несколько колонок высота — по самой высокой колонке.
    Возвращает число подогнанных листов.
    """
    body = docx.element.body
    for tbl in body.iter(w("tbl")):
        _thin_filler_rows(tbl)
    fitted = 0
    header_h = footer_h = 0.0
    for section, parts in _page_groups(docx):
        # колонтитулы выше полей раздвигают поля — Word считает от их реальной высоты
        try:
            width = (section.page_width - section.left_margin - section.right_margin) / 12700.0
            if not section.header.is_linked_to_previous:
                header_h = story_height(section.header._element, width)
            if not section.footer.is_linked_to_previous:
                footer_h = story_height(section.footer._element, width)
            top = max(section.top_margin / 12700.0, section.header_distance / 12700.0 + header_h)
            bottom = max(section.bottom_margin / 12700.0, section.footer_distance / 12700.0 + footer_h)
            avail = section.page_height / 12700.0 - top - bottom
        except Exception:
            continue
        sizes = [part.height() for part in parts]
        budget = sum(h for h, _ in sizes)
        rows = sum(r for _, r in sizes)
        excess = budget - (avail - FIT_SAFETY - FIT_SAFETY_PER_ROW * rows)
        if excess <= 0:
            continue
        fitted += 1
        # первый элемент листа не трогаем: он ставит таблицу к верху рамки
        first = parts[0].elems[0] if parts[0].elems else None
        single = [el for part in parts if part.cols == 1 for el in part.elems if el is not first]
        excess -= _reclaim(single, excess)
        for part in parts:
            if excess <= 0 or part.cols == 1:
                continue
            # колонки: ужимаем те, что выше целевой высоты раздела
            sec_h = part.height()[0]
            target = sec_h - excess
            for chunk in part.chunks:
                h = _chunk_height(chunk, part.width)[0]
                if h > target:
                    _reclaim([el for el in chunk if el is not first], h - target)
            excess -= sec_h - part.height()[0]
        tables = [el for part in parts if part.cols == 1 for el in part.elems if el.tag == w("tbl")]
        if excess > 0 and tables:
            excess -= _shrink_last_rows(tables[-1], excess)
    return fitted


def improve_structure(docx) -> dict[str, int]:
    body = docx.element.body
    fix_default_spacing(docx)
    shrink_section_breaks(body)
    res = {"grids": fix_table_grids(body), "heading_tables": heading_tables_to_paragraphs(body)}
    res["toc"] = fix_toc(docx)
    res["headings"] = mark_headings(docx)
    return res


# --------------------------------------------------------------------------- режим совместимости

WORD_URI = "http://schemas.microsoft.com/office/word"


def set_compatibility_mode(docx, mode: int) -> None:
    """Режим совместимости Word (14 — Word 2010, 15 — Word 2013 и новее).

    В режиме 15 Word (и LibreOffice с 24.2) сжимает пробелы в строках с выравниванием по ширине,
    как это делал Word, сохранивший исходный PDF, и ставит таблицы по границе, а не по тексту
    первой ячейки; документ к тому же не открывается «в режиме ограниченной функциональности».
    """
    settings = docx.settings.element
    compat = settings.find(w("compat"))
    if compat is None:
        compat = etree.SubElement(settings, w("compat"))
    for el in compat.findall(w("compatSetting")):
        if el.get(w("name")) == "compatibilityMode":
            el.set(w("val"), str(mode))
            return
    el = etree.SubElement(compat, w("compatSetting"))
    el.set(w("name"), "compatibilityMode")
    el.set(w("uri"), WORD_URI)
    el.set(w("val"), str(mode))
