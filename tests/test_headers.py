"""Колонтитулы из оформления листа (app/headers.py) — только python-docx, без PyMuPDF."""
from __future__ import annotations

import io
import zipfile

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt

from app import headers

PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde"
       b"\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0\x00\x00\x03\x01\x01\x00\xc9\xfe\x92\xef\x00\x00\x00\x00IEND"
       b"\xaeB`\x82")


def _band_doc(sheet: int, with_logo: bool = True) -> Document:
    """Как pdf2docx: пустой абзац-отступ сверху, таблица штампа, маленький абзац после таблицы."""
    d = Document()
    spacer = d.add_paragraph()
    spacer.paragraph_format.space_after = Pt(700)
    if with_logo:
        d.add_paragraph().add_run().add_picture(io.BytesIO(PNG))
    t = d.add_table(rows=1, cols=3)
    t.cell(0, 0).text = "2026"
    t.cell(0, 1).text = "Трубопровод технического конденсата"
    t.cell(0, 2).text = str(sheet)
    d.add_paragraph()
    return d


def _main(n: int) -> Document:
    d = Document()
    for k in range(n):
        if k:
            d.add_section(WD_SECTION.NEW_PAGE)
        d.add_paragraph(f"Текст листа {k + 1}")
    return d


def _contents(pages, sheets):
    return {p: {"footer": headers.prepare(_band_doc(s))} for p, s in zip(pages, sheets)}


def test_prepare_strips_spacers_and_keeps_image():
    band = headers.prepare(_band_doc(3))
    tags = [el.tag.split("}")[1] for el in band.elements]
    assert tags == ["p", "tbl"]                 # отступ сверху и хвостовой абзац убраны
    assert len(band.images) == 1                # логотип по отпечатку, без чужих rId
    assert b"r:embed" not in b"".join(__import__("lxml.etree", fromlist=["x"]).tostring(e) for e in band.elements)


def test_page_fields_and_linking(tmp_path):
    d = _main(3)
    contents = _contents([0, 1, 2], [2, 3, 4])          # листы 2, 3, 4 на страницах 1–3
    first = headers.page_fields(contents)
    assert first == 2
    page_secs = [s for s in d.sections if s.start_type == WD_SECTION.NEW_PAGE]
    made = headers.attach(list(zip([0, 1, 2], page_secs)), contents, {0: {"footer": 20.0}})
    assert made == 1                                      # после замены номера на PAGE колонтитулы одинаковы
    headers.set_start_page(page_secs[0], first)
    out = tmp_path / "f.docx"
    d.save(out)
    z = zipfile.ZipFile(out)
    footers = [n for n in z.namelist() if n.startswith("word/footer")]
    assert len(footers) == 1
    xml = z.read(footers[0]).decode()
    assert 'w:instr=" PAGE "' in xml and "Трубопровод" in xml and "r:embed" in xml
    doc_xml = z.read("word/document.xml").decode()
    assert 'w:start="2"' in doc_xml
    d2 = Document(out)                                    # файл открывается заново
    assert d2.sections[1].footer.is_linked_to_previous
    assert abs(d2.sections[0].footer_distance.pt - 20.0) < 0.1


def test_different_footers_are_not_linked(tmp_path):
    d = _main(2)
    contents = {0: {"footer": headers.prepare(_band_doc(7))},
                1: {"footer": headers.prepare(_band_doc(7, with_logo=False))}}
    assert headers.page_fields(contents) is None          # номер не совпадает с номером страницы
    page_secs = list(d.sections)
    assert headers.attach(list(zip([0, 1], page_secs)), contents, {}) == 2
    out = tmp_path / "g.docx"
    d.save(out)
    assert not Document(out).sections[1].footer.is_linked_to_previous


def test_page_without_furniture_gets_empty_footer(tmp_path):
    d = _main(2)
    contents = {0: {"footer": headers.prepare(_band_doc(1))}}
    headers.attach(list(zip([0, 1], d.sections)), contents, {})
    out = tmp_path / "e.docx"
    d.save(out)
    d2 = Document(out)
    assert not d2.sections[1].footer.is_linked_to_previous
    assert "Трубопровод" not in "".join(p.text for p in d2.sections[1].footer.paragraphs)


def test_frame_and_floating_side_stamp(tmp_path):
    d = _main(1)
    side = Document()
    t = side.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "Инв. № подл."
    band = headers.prepare(side, floating=[(22.7, 587.0, 56.7, 828.0)])
    band.elements.append(headers.frame_paragraph((56.7, 14.2, 1176.4, 827.7)))
    headers.attach([(0, d.sections[0])], {0: {"header": band}}, {0: {"header": 0.0}})
    out = tmp_path / "h.docx"
    d.save(out)
    z = zipfile.ZipFile(out)
    xml = z.read([n for n in z.namelist() if n.startswith("word/header")][0]).decode()
    assert 'w:tblpX="454"' in xml and 'w:tblpY="11740"' in xml and 'w:vertAnchor="page"' in xml
    assert "v:rect" in xml and "margin-left:56.70pt" in xml
    assert xml.rstrip().endswith("</w:p></w:hdr>")       # колонтитул заканчивается абзацем


def test_side_stamp_in_textbox(tmp_path):
    d = _main(1)
    side = Document()
    side.add_table(rows=2, cols=2).cell(0, 0).text = "Инв. № подл."
    band = headers.prepare(side)
    tbl = band.elements[0]
    box = headers.Band([headers.textbox_paragraph([tbl], (22.7, 587.0, 56.7, 828.0))])
    headers.attach([(0, d.sections[0])], {0: {"header": box}}, {})
    out = tmp_path / "s.docx"
    d.save(out)
    xml = zipfile.ZipFile(out).read([n for n in zipfile.ZipFile(out).namelist() if n.startswith("word/header")][0]).decode()
    assert "w:txbxContent" in xml and "Инв. № подл." in xml and 'type="none"' in xml
    assert "margin-left:22.70pt" in xml and "margin-top:587.00pt" in xml


def test_spacer_with_anchored_image_is_flattened():
    from docx.oxml.ns import qn
    d = Document()
    spacer = d.add_paragraph()
    spacer.paragraph_format.space_after = Pt(568)
    spacer.add_run().add_picture(io.BytesIO(PNG))          # как подпись, привязанная к отступу
    inline = next(spacer._p.iter(qn("wp:inline")))
    inline.tag = qn("wp:anchor")                           # pdf2docx вставляет такие картинки плавающими
    d.add_table(rows=1, cols=2).cell(0, 0).text = "Штамп"
    band = headers.prepare(d)
    first = band.elements[0]
    sp = first.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    assert sp.get(qn("w:after")) == "0" and sp.get(qn("w:line")) == "20"
    assert len(band.images) == 1


def test_inline_logo_keeps_its_height():
    from docx.oxml.ns import qn
    d = Document()
    logo = d.add_paragraph()
    logo.add_run().add_picture(io.BytesIO(PNG), width=Pt(119))   # логотип в шапке — картинка в строке
    d.add_paragraph("Е110-0162-УКК.25.1087-РД-01-03.03.086-АТХ.ОЛ1")
    band = headers.prepare(d)
    sp = band.elements[0].find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    assert sp is None or sp.get(qn("w:lineRule")) != "exact"


def test_header_columns_become_borderless_table():
    """pdf2docx кладёт «логотип | название организации» в два раздела-колонки; в колонтитуле
    разделов нет — блоки стоят рядом в таблице без границ, а не друг под другом."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn

    d = Document()
    logo = d.add_paragraph()
    logo.add_run().add_picture(io.BytesIO(PNG))
    logo.paragraph_format.space_after = Pt(26)

    def end_section(kind: str) -> None:
        p = d.add_paragraph()
        p._p.get_or_add_pPr().append(parse_xml(
            f'<w:sectPr {nsdecls("w")}><w:type w:val="{kind}"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:cols w:num="2" w:equalWidth="0"><w:col w:w="3467" w:space="0"/>'
            '<w:col w:w="6398" w:space="0"/></w:cols></w:sectPr>'))

    end_section("continuous")
    d.add_paragraph("Общество с ограниченной ответственностью")
    end_section("nextColumn")
    d.add_paragraph("Е110-0162-АТХ.ОЛ")
    band = headers.prepare(d)
    tables = [el for el in band.elements if el.tag == qn("w:tbl")]
    assert len(tables) == 1
    cells = tables[0].findall(f"{qn('w:tr')}/{qn('w:tc')}")
    assert len(cells) == 2
    assert next(cells[0].iter(qn("w:drawing")), None) is not None
    assert "Общество" in "".join(t.text or "" for t in cells[1].iter(qn("w:t")))
    assert [g.get(qn("w:w")) for g in tables[0].iter(qn("w:gridCol"))] == ["3467", "6398"]
    assert not any(el.find(f".//{qn('w:sectPr')}") is not None for el in band.elements)
    # отступ под логотипом — после таблицы, чтобы название встало по середине логотипа
    after = band.elements[band.elements.index(tables[0]) + 1]
    assert after.find(f"{qn('w:pPr')}/{qn('w:spacing')}").get(qn("w:before")) == str(26 * 20 - 20)


def test_empty_header_after_real_one_stays_inside_margin(tmp_path):
    # лист А3 со штампом: поле сверху 15 pt, своей шапки нет; пустой колонтитул на 36 pt
    # сдвигал бы текст вниз и лист переезжал на следующую страницу
    d = _main(2)
    d.sections[1].top_margin = Pt(15)
    d.sections[1].bottom_margin = Pt(13)
    contents = {0: {"header": headers.prepare(_band_doc(1)), "footer": headers.prepare(_band_doc(1))}}
    headers.attach(list(zip([0, 1], d.sections)), contents, {})
    out = tmp_path / "m.docx"
    d.save(out)
    s0, s1 = Document(out).sections
    assert s1.header_distance.pt <= 13.0 and s1.footer_distance.pt <= 11.0
    assert s0.header_distance.pt > 30                     # у листа с шапкой расстояние прежнее
