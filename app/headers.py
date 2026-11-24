"""Колонтитулы Word из оформления листа (см. app/furniture.py).

* band_pass — разбор pdf2docx только областей оформления (фильтры app/regionfix.py): на каждую
  страницу — отдельный python-docx Document;
* prepare — содержимое такого документа для колонтитула: без служебных пустых абзацев и разрывов
  разделов, картинки — по отпечатку (привязываются к части колонтитула при вставке);
* floating_table — боковой штамп: «плавающая» таблица в точных координатах листа;
* frame_paragraph — рамка листа (векторный прямоугольник за текстом);
* page_fields — номер листа в штампе → поле PAGE (Word сам пронумерует листы);
* attach — вставка по разделам; одинаковые колонтитулы соседних листов связываются
  («Как в предыдущем разделе»), поэтому правка шапки или штампа меняет все листы сразу.
"""
from __future__ import annotations

import copy
import hashlib
import io
import logging
from dataclasses import dataclass, field

from lxml import etree

log = logging.getLogger(__name__)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W10 = "urn:schemas-microsoft-com:office:word"
IMG_KEY = "{urn:pdf2word}img"          # временный атрибут: отпечаток картинки до привязки к части


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def twips(pt: float) -> str:
    return str(int(round(pt * 20)))


@dataclass
class Band:
    elements: list                       # элементы w:p / w:tbl
    images: dict[str, bytes] = field(default_factory=dict)


# --------------------------------------------------------------------------- разбор областей

def band_pass(fitz_doc, cfg: dict, specs: dict) -> dict:
    """{номер страницы: (Document, [bbox таблиц верхнего уровня])} для областей из specs."""
    from docx import Document
    from pdf2docx.page.Page import Page
    from pdf2docx.page.Pages import Pages

    from .regionfix import regions
    from .textfix import parse_page_safely

    out: dict = {}
    if not specs:
        return out
    pages = Pages()
    pages.reset([Page(id=i, skip_parsing=i not in specs) for i in range(fitz_doc.page_count)])
    with regions(fitz_doc, specs):
        pages.parse(fitz_doc, **cfg)
        for page in pages:
            if page.skip_parsing:
                continue
            try:
                parse_page_safely(page, cfg)
                d = Document()
                page.make_docx(d)
                tables = [tuple(b.bbox) for section in page.sections for column in section
                          for b in column.blocks if b.is_table_block]
                out[page.id] = (d, tables)
            except Exception as exc:
                log.warning("header/footer pass failed on page %s: %s", page.id + 1, exc)
    return out


# --------------------------------------------------------------------------- подготовка содержимого

def _is_empty_paragraph(p) -> bool:
    if p.tag != w("p"):
        return False
    if any((t.text or "").strip() for t in p.iter(w("t"))):
        return False
    for tag in ("drawing", "pict", "object", "fldSimple", "fldChar"):
        if next(p.iter(w(tag)), None) is not None:
            return False
    return True


def _make_tiny(p) -> None:
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    sp = ppr.find(w("spacing"))
    if sp is None:
        sp = etree.Element(w("spacing"))
        ind = ppr.find(w("ind"))
        (ind.addprevious(sp) if ind is not None else ppr.append(sp))
    for k, v in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
        sp.set(w(k), v)
    for attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
        sp.attrib.pop(w(attr), None)


def _zero_space_before(p) -> None:
    ppr = p.find(w("pPr"))
    if ppr is None:
        return
    sp = ppr.find(w("spacing"))
    if sp is not None:
        sp.set(w("before"), "0")
        for attr in ("beforeLines", "beforeAutospacing"):
            sp.attrib.pop(w(attr), None)


def _section_layout(sp) -> tuple[str, list[int]]:
    """(тип начала, ширины колонок в twips) по w:sectPr раздела pdf2docx."""
    if sp is None:
        return "nextPage", []
    t = sp.find(w("type"))
    kind = t.get(w("val"), "nextPage") if t is not None else "nextPage"
    cols = sp.find(w("cols"))
    try:
        num = int(cols.get(w("num"), "1")) if cols is not None else 1
    except ValueError:
        num = 1
    if num < 2:
        return kind, []
    widths = [int(float(c.get(w("w"), "0"))) for c in cols.findall(w("col"))]
    if len(widths) != num or not all(widths):
        size, mar = sp.find(w("pgSz")), sp.find(w("pgMar"))
        try:
            text = int(size.get(w("w"))) - int(mar.get(w("left"))) - int(mar.get(w("right")))
        except (AttributeError, TypeError, ValueError):
            text = 9000
        widths = [text // num] * num
    return kind, widths


def _has_content(elems) -> bool:
    return any(el.tag == w("tbl") or any((t.text or "").strip() for t in el.iter(w("t")))
               or next(el.iter(f"{{{WP}}}inline"), None) is not None for el in elems)


def _layout_table(chunks: list[list], widths: list[int]):
    """Блоки рядом — таблица без границ в одну строку (в колонтитуле разделов с колонками нет)."""
    tbl = etree.SubElement(etree.Element(w("body")), w("tbl"))
    pr = etree.SubElement(tbl, w("tblPr"))
    etree.SubElement(pr, w("tblW"), {w("w"): str(sum(widths)), w("type"): "dxa"})
    etree.SubElement(pr, w("tblLayout"), {w("type"): "fixed"})
    borders = etree.SubElement(pr, w("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        etree.SubElement(borders, w(side), {w("val"): "nil"})
    mar = etree.SubElement(pr, w("tblCellMar"))
    for side in ("left", "right"):
        etree.SubElement(mar, w(side), {w("w"): "0", w("type"): "dxa"})
    grid = etree.SubElement(tbl, w("tblGrid"))
    for width in widths:
        etree.SubElement(grid, w("gridCol"), {w("w"): str(width)})
    tr = etree.SubElement(tbl, w("tr"))
    gap = 0
    for k, width in enumerate(widths):
        tc = etree.SubElement(tr, w("tc"))
        tcpr = etree.SubElement(tc, w("tcPr"))
        etree.SubElement(tcpr, w("tcW"), {w("w"): str(width), w("type"): "dxa"})
        # блоки по центру строки (название организации — по середине логотипа), а интервал под
        # последним абзацем — отступ до текста ниже — выносится за таблицу
        etree.SubElement(tcpr, w("vAlign"), {w("val"): "center"})
        for el in chunks[k] if k < len(chunks) else []:
            tc.append(el)
        while len(tc) > 1 and tc[-1].tag == w("p") and not _has_content([tc[-1]]):
            tc.remove(tc[-1])                # пустые абзацы-отступы в конце колонки
        last = tc[-1] if len(tc) > 1 else None
        if last is not None and last.tag == w("p"):
            sp = last.find(f"{w('pPr')}/{w('spacing')}")
            if sp is not None:
                try:
                    gap = max(gap, int(float(sp.get(w("after"), "0"))))
                except ValueError:
                    pass
                sp.set(w("after"), "0")
        if tc[-1].tag != w("p"):
            etree.SubElement(tc, w("p"))     # ячейка Word кончается абзацем
    if gap <= 20:
        return [tbl]
    spacer = etree.Element(w("p"))
    ppr = etree.SubElement(spacer, w("pPr"))
    etree.SubElement(ppr, w("spacing"), {w("before"): str(gap - 20), w("after"): "0", w("line"): "20",
                                         w("lineRule"): "exact"})
    return [tbl, spacer]


def _columns_as_tables(elements: list) -> list:
    """pdf2docx раскладывает блоки рядом (логотип и название организации) по колонкам — разделам
    «с новой колонки». В колонтитуле разделов нет: колонки шли бы друг под другом, поэтому они
    становятся таблицей без границ."""
    sections, current = [], []
    for el in elements:
        current.append(el)
        sp = el.find(f"{w('pPr')}/{w('sectPr')}") if el.tag == w("p") else None
        if sp is not None:
            sections.append((current, sp))
            current = []
    if current:
        sections.append((current, None))
    groups: list[tuple[list[int], list[list]]] = []
    for elems, sp in sections:
        kind, widths = _section_layout(sp)
        if widths and kind == "nextColumn" and groups and groups[-1][0]:
            groups[-1][1].append(elems)
        else:
            groups.append((widths, [elems]))
    out: list = []
    for widths, chunks in groups:
        if widths and len(chunks) == len(widths) and sum(1 for c in chunks if _has_content(c)) >= 2:
            out.extend(_layout_table(chunks, widths))
        else:
            for c in chunks:
                out.extend(c)
    return out


def prepare(src_doc, floating: list | None = None) -> Band:
    """Содержимое тела документа src_doc → элементы для колонтитула."""
    body = src_doc.element.body
    band = Band([])
    for el in body:
        if el.tag in (w("p"), w("tbl")):
            band.elements.append(copy.deepcopy(el))
    original = {id(el) for el in band.elements if el.tag == w("tbl")}
    try:
        band.elements = _columns_as_tables(band.elements)
    except Exception:
        log.debug("header columns kept as is", exc_info=True)
    for el in band.elements:
        for sp in list(el.iter(w("sectPr"))):
            sp.getparent().remove(sp)
        for br in list(el.iter(w("br"))):
            if br.get(w("type")) == "column":
                br.getparent().remove(br)
        for hl in list(el.iter(w("hyperlink"))):
            parent = hl.getparent()
            idx = parent.index(hl)
            for child in list(hl):
                parent.insert(idx, child)
                idx += 1
            parent.remove(hl)
        for node, attr in [(b, f"{{{R}}}embed") for b in el.iter(f"{{{A}}}blip")] + \
                          [(i, f"{{{R}}}id") for i in el.iter(f"{{{V}}}imagedata")]:
            rid = node.get(attr)
            try:
                blob = src_doc.part.related_parts[rid].blob
            except Exception:
                continue
            key = hashlib.sha1(blob).hexdigest()
            band.images[key] = blob
            node.attrib.pop(attr, None)
            node.set(IMG_KEY, f"{attr}|{key}")
    from .docxfix import fix_table_grids
    for el in band.elements:
        fix_table_grids(el)
    while band.elements and _is_empty_paragraph(band.elements[0]):
        band.elements.pop(0)
    while band.elements and _is_empty_paragraph(band.elements[-1]):
        band.elements.pop()
    # служебный абзац-отступ pdf2docx, к которому привязаны картинки (подписи в штампе): картинки
    # стоят по координатам листа, а отступ в сотни пунктов раздул бы колонтитул на всю страницу
    for el in band.elements:
        if el.tag != w("p") or any((t.text or "").strip() for t in el.iter(w("t"))):
            break
        if next(el.iter(f"{{{WP}}}inline"), None) is not None:
            break       # картинка в строке (логотип) занимает высоту абзаца — его не сжимать
        _make_tiny(el)
    if band.elements and band.elements[0].tag == w("p"):
        _zero_space_before(band.elements[0])
    if floating:
        # таблицы-раскладки колонок в pdf2docx не было — у них нет координат
        tables = [el for el in band.elements if el.tag == w("tbl") and id(el) in original]
        for tbl, bbox in zip(tables, floating):
            floating_table(tbl, bbox)
    return band


def floating_table(tbl, bbox) -> None:
    """Таблица в точных координатах листа (не влияет на поток текста колонтитула)."""
    tblpr = tbl.find(w("tblPr"))
    if tblpr is None:
        tblpr = etree.Element(w("tblPr"))
        tbl.insert(0, tblpr)
    for tag in ("tblpPr", "tblOverlap", "jc", "tblInd"):
        old = tblpr.find(w(tag))
        if old is not None:
            tblpr.remove(old)
    pos = etree.Element(w("tblpPr"))
    for k, v in (("leftFromText", "0"), ("rightFromText", "0"), ("topFromText", "0"), ("bottomFromText", "0"),
                 ("vertAnchor", "page"), ("horzAnchor", "page"),
                 ("tblpX", twips(bbox[0])), ("tblpY", twips(bbox[1]))):
        pos.set(w(k), v)
    overlap = etree.Element(w("tblOverlap"))
    overlap.set(w("val"), "overlap")
    style = tblpr.find(w("tblStyle"))
    at = 1 if style is not None else 0
    tblpr.insert(at, pos)
    tblpr.insert(at + 1, overlap)


def textbox_paragraph(elements, bbox):
    """Абзац-якорь с надписью VML в точных координатах листа и без обтекания.

    Плавающая таблица в колонтитуле — препятствие: Word и LibreOffice заставляют текст страницы её
    обтекать, и таблица спецификации, задевшая боковой штамп хоть на долю пункта, уезжала на
    следующую страницу. Надпись без обтекания ничему не мешает.
    """
    p = tiny_paragraph()
    x0, y0, x1, y1 = bbox
    r = etree.SubElement(p, w("r"))
    rpr = etree.SubElement(r, w("rPr"))
    etree.SubElement(rpr, w("sz")).set(w("val"), "2")
    pict = etree.SubElement(r, w("pict"))
    rect = etree.SubElement(pict, f"{{{V}}}rect", nsmap={"v": V, "w10": W10})
    rect.set("style", (
        f"position:absolute;margin-left:{x0:.2f}pt;margin-top:{y0:.2f}pt;"
        f"width:{x1 - x0 + 2:.2f}pt;height:{y1 - y0 + 2:.2f}pt;z-index:251659264;"
        "mso-position-horizontal:absolute;mso-position-horizontal-relative:page;"
        "mso-position-vertical:absolute;mso-position-vertical-relative:page"))
    rect.set("filled", "f")
    rect.set("stroked", "f")
    box = etree.SubElement(rect, f"{{{V}}}textbox")
    box.set("inset", "0,0,0,0")
    content = etree.SubElement(box, w("txbxContent"))
    for el in elements:
        if el.tag == w("tbl"):
            tblpr = el.find(w("tblPr"))
            if tblpr is not None:
                for tag in ("tblpPr", "tblOverlap", "jc", "tblInd"):
                    old = tblpr.find(w(tag))
                    if old is not None:
                        tblpr.remove(old)
        content.append(el)
    if not len(content) or content[-1].tag != w("p"):
        content.append(tiny_paragraph())
    wrap = etree.SubElement(rect, f"{{{W10}}}wrap")
    wrap.set("type", "none")
    wrap.set("anchorx", "page")
    wrap.set("anchory", "page")
    return p


def tiny_paragraph():
    p = etree.Element(w("p"))
    ppr = etree.SubElement(p, w("pPr"))
    sp = etree.SubElement(ppr, w("spacing"))
    for k, v in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
        sp.set(w(k), v)
    rpr = etree.SubElement(ppr, w("rPr"))
    etree.SubElement(rpr, w("sz")).set(w("val"), "2")
    return p


def frame_paragraph(frame, width: float = 1.0):
    """Абзац-якорь с рамкой листа: прямоугольник VML за текстом, координаты от края листа."""
    p = tiny_paragraph()
    x0, y0, x1, y1 = frame
    r = etree.SubElement(p, w("r"))
    rpr = etree.SubElement(r, w("rPr"))
    etree.SubElement(rpr, w("sz")).set(w("val"), "2")
    pict = etree.SubElement(r, w("pict"))
    rect = etree.SubElement(pict, f"{{{V}}}rect", nsmap={"v": V, "w10": W10})
    rect.set("style", (
        f"position:absolute;margin-left:{x0:.2f}pt;margin-top:{y0:.2f}pt;"
        f"width:{x1 - x0:.2f}pt;height:{y1 - y0:.2f}pt;z-index:-251657216;"
        "mso-position-horizontal:absolute;mso-position-horizontal-relative:page;"
        "mso-position-vertical:absolute;mso-position-vertical-relative:page"))
    rect.set("filled", "f")
    rect.set("strokecolor", "black")
    rect.set("strokeweight", f"{width:.2f}pt")
    wrap = etree.SubElement(rect, f"{{{W10}}}wrap")
    wrap.set("anchorx", "page")
    wrap.set("anchory", "page")
    return p


# --------------------------------------------------------------------------- номер листа

def _digit_runs(elements):
    for el in elements:
        for t in el.iter(w("t")):
            text = (t.text or "").strip()
            if text.isdigit() and len(text) <= 4 and t.getparent() is not None and t.getparent().tag == w("r"):
                yield int(text), t


def page_fields(contents: dict[int, dict[str, Band]]) -> int | None:
    """Номер листа, совпадающий с номером страницы (со сдвигом), → поле PAGE.
    contents: {номер страницы (0…): {'header'|'footer': Band}}. Возвращает номер первого листа."""
    pages = sorted(contents)
    if len(pages) < 2:
        return None
    votes: dict[int, int] = {}
    for pid in pages:
        seen = set()
        for kind in ("footer", "header"):
            band = contents[pid].get(kind)
            if band is None:
                continue
            for n, _ in _digit_runs(band.elements):
                seen.add(n - (pid + 1))
        for k in seen:
            votes[k] = votes.get(k, 0) + 1
    if not votes:
        return None
    k, support = max(votes.items(), key=lambda kv: (kv[1], -abs(kv[0])))
    if support < max(2, round(0.6 * len(pages))):
        return None
    for pid in pages:
        target = pid + 1 + k
        done = False
        for kind in ("footer", "header"):
            band = contents[pid].get(kind)
            if band is None or done:
                continue
            for n, t in _digit_runs(band.elements):
                if n == target:
                    run = t.getparent()
                    fld = etree.Element(w("fldSimple"))
                    fld.set(w("instr"), " PAGE ")
                    run.addprevious(fld)
                    fld.append(run)
                    done = True
                    break
    return pages[0] + 1 + k


# --------------------------------------------------------------------------- вставка в разделы

def _canonical(band: Band | None) -> bytes:
    if band is None or not band.elements:
        return b""
    parts = []
    for el in band.elements:
        el = copy.deepcopy(el)
        for fld in el.iter(w("fldSimple")):      # результат поля PAGE на каждом листе свой
            for t in fld.iter(w("t")):
                t.text = "#"
        parts.append(etree.tostring(el))
    return b"".join(parts)


def _bind_images(elements, images: dict[str, bytes], part) -> None:
    cache: dict[str, str] = {}
    for el in elements:
        for node in list(el.iter()):
            mark = node.attrib.pop(IMG_KEY, None) if isinstance(node.tag, str) else None
            if not mark:
                continue
            attr, key = mark.split("|", 1)
            if key not in cache:
                rid, _ = part.get_or_add_image(io.BytesIO(images[key]))
                cache[key] = rid
            node.set(attr, cache[key])
        etree.cleanup_namespaces(el)


def _fill(hf, band: Band | None) -> None:
    root = hf._element
    for child in list(root):
        root.remove(child)
    elements = [copy.deepcopy(el) for el in band.elements] if band else []
    if elements:
        _bind_images(elements, band.images, hf.part)
    if not elements or elements[-1].tag != w("p"):
        elements.append(tiny_paragraph())
    for el in elements:
        root.append(el)


def attach(page_sections: list, contents: dict[int, dict[str, Band]],
           distances: dict[int, dict[str, float]]) -> int:
    """page_sections: [(номер страницы, python-docx Section первого раздела страницы)].
    Возвращает число созданных колонтитулов."""
    from docx.shared import Pt

    made = 0
    prev = {"header": None, "footer": None}
    for pid, section in page_sections:
        bands = contents.get(pid, {})
        for kind in ("header", "footer"):
            band = bands.get(kind)
            canon = _canonical(band)
            hf = section.header if kind == "header" else section.footer
            if prev[kind] is None and not canon:
                continue                       # колонтитулов ещё не было — оставляем пустым
            if canon == prev[kind]:
                hf.is_linked_to_previous = True
                continue
            hf.is_linked_to_previous = False
            _fill(hf, band)
            prev[kind] = canon
            made += 1
        dist = distances.get(pid, {})
        if "header" in dist:
            section.header_distance = Pt(max(0.0, dist["header"]))
        if "footer" in dist:
            section.footer_distance = Pt(max(0.0, dist["footer"]))
    return made


def set_start_page(section, number: int) -> None:
    sect = section._sectPr
    el = sect.find(w("pgNumType"))
    if el is None:
        el = etree.Element(w("pgNumType"))
        sect.insert_element_before(el, "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                                   "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
                                   "w:printerSettings", "w:sectPrChange")
    el.set(w("start"), str(number))
