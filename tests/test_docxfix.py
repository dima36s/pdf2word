"""Структура документа (app/docxfix.py): заголовки, таблицы-заголовки, содержание — только python-docx."""
from __future__ import annotations

import zipfile

import pytest
from docx import Document
from docx.shared import Pt

from app import docxfix


def _p(d, text, bold=False, size=11):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    return p


def _style(p) -> str | None:
    ppr = p._p.pPr
    return None if ppr is None or ppr.pStyle is None else ppr.pStyle.val


@pytest.mark.parametrize("text,bold,size,level", [
    ("2 Характеристика принятой технологической схемы", True, 12, 1),
    ("2.1 Общие характеристики", True, 11, 2),
    ("2.2.1 Конвейерно-транспортная галерея", True, 11, 3),
    ("3.1.2.4 Насосы", True, 11, 4),
    ("1. Анализ текущего процесса", False, 11, None),        # пункт списка, не выделен
    ("2026 г. Срок оплаты", True, 11, None),                  # год, а не номер раздела
    ("2.1 Перечень:", True, 11, None),
    ("Содержание", True, 12, 1),
    ("Приложение А", True, 12, 1),
    ("Введение", False, 16, 1),                               # выделен размером
    ("2.1 Общие характеристики\t5", True, 11, None),          # строка содержания
])
def test_heading_level(text, bold, size, level):
    assert docxfix.heading_level(text, 1.0 if bold else 0.0, size, 11) == level


def test_mark_headings_keeps_look(tmp_path):
    d = Document()
    h = _p(d, "2 Характеристика принятой технологической схемы", True, 12)
    sub = _p(d, "2.1 Общие характеристики", True, 11)
    body = [_p(d, "Технологическая схема предусматривает переработку руды. " * 3) for _ in range(4)]
    assert docxfix.mark_headings(d) == 2
    assert _style(h) == "Heading1" and _style(sub) == "Heading2" and all(_style(p) is None for p in body)
    # стиль заголовка не навязывает цвет, шрифт и интервалы — вид остаётся как в PDF
    st = d.styles["Heading 1"].element
    assert st.find(docxfix.w("rPr")) is None
    assert st.find(f"{docxfix.w('pPr')}/{docxfix.w('outlineLvl')}") is not None
    out = tmp_path / "h.docx"
    d.save(out)
    assert Document(out).paragraphs[0].style.name == "Heading 1"


def test_heading_table_becomes_paragraph():
    d = Document()
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).paragraphs[0].add_run("2").bold = True
    t.cell(0, 1).paragraphs[0].add_run("Характеристика принятой технологической схемы").bold = True
    t.cell(1, 0).paragraphs[0].add_run("2.1").bold = True
    t.cell(1, 1).paragraphs[0].add_run("Общие характеристики").bold = True
    framed = d.add_table(rows=1, cols=2)            # обычная таблица с данными не трогается
    framed.cell(0, 0).text = "1"
    framed.cell(0, 1).text = "Насос"
    framed.style = "Table Grid"
    for c in framed.rows[0].cells:
        docxfix_set_border(c)
    res = docxfix.improve_structure(d)
    assert res["heading_tables"] == 1 and res["headings"] == 2
    assert len(d.tables) == 1
    texts = [p.text for p in d.paragraphs if p.text.strip()]
    assert texts[:2] == ["2\tХарактеристика принятой технологической схемы", "2.1\tОбщие характеристики"]
    assert d.paragraphs[0].paragraph_format.first_line_indent < 0   # выступ: текст под текстом


def docxfix_set_border(cell):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcpr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    e = OxmlElement("w:top")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), "4")
    b.append(e)
    tcpr.append(b)


def test_toc_leaders(tmp_path):
    d = Document()
    entries = [_p(d, "2.1 Общие характеристики принятой схемы " + "." * 60 + " 5"),
               _p(d, "2.2 Характеристика параметров .........................  6"),
               _p(d, "Обозначения и сокращения" + "." * 30 + "3")]
    plain = _p(d, "Итого за 2026 год 12")
    assert docxfix.fix_toc(d) == 3
    assert entries[0].text == "2.1 Общие характеристики принятой схемы\t5"
    assert entries[1].text == "2.2 Характеристика параметров\t6"
    assert entries[2].text == "Обозначения и сокращения\t3"
    assert plain.text == "Итого за 2026 год 12"
    out = tmp_path / "t.docx"
    d.save(out)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert xml.count('w:leader="dot"') == 3 and 'w:val="right"' in xml


def test_table_grid_from_cell_widths():
    from docx.oxml.ns import qn
    d = Document()
    t = d.add_table(rows=2, cols=3)
    widths = [(400, 1600, 900), None]
    for c, wd in zip(t.rows[0].cells, widths[0]):
        c.width = wd * 635                       # python-docx: EMU; 1 твип = 635 EMU
    merged = t.cell(1, 0).merge(t.cell(1, 1))    # объединённая ячейка не мешает
    merged.width = 2000 * 635
    t.cell(1, 2).width = 900 * 635
    t._tbl.tblPr.append(d.element.makeelement(qn("w:tblInd"), {qn("w:w"): "11906.0", qn("w:type"): "dxa"}))
    assert docxfix.fix_table_grids(d.element.body) == 1
    grid = [int(g.get(qn("w:w"))) for g in t._tbl.iter(qn("w:gridCol"))]
    assert grid == [400, 1600, 900]
    # продолжение вертикального объединения с шириной-заглушкой не портит сетку
    t2 = d.add_table(rows=2, cols=2)
    t2.cell(0, 0).width, t2.cell(0, 1).width = 300 * 635, 700 * 635
    t2.cell(0, 1).merge(t2.cell(1, 1))
    t2.cell(1, 0).width = 300 * 635
    t2._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))[1].tcPr.find(qn("w:tcW")).set(qn("w:w"), "2491")
    docxfix.fix_table_grid(t2._tbl)
    assert [int(g.get(qn("w:w"))) for g in t2._tbl.iter(qn("w:gridCol"))] == [300, 700]
    assert t._tbl.tblPr.find(qn("w:tblInd")).get(qn("w:w")) == "11906"


def test_empty_cells_and_default_spacing(tmp_path):
    from docx.oxml.ns import qn
    d = Document()
    t = d.add_table(rows=2, cols=1)
    r = t.cell(0, 0).paragraphs[0].add_run("Датчик давления")
    r.font.size = Pt(8)
    docxfix.improve_structure(d)
    empty = t.cell(1, 0).paragraphs[0]._p
    assert empty.pPr.find(qn("w:rPr")).find(qn("w:sz")).get(qn("w:val")) == "16"
    out = tmp_path / "s.docx"
    d.save(out)
    styles = zipfile.ZipFile(out).read("word/styles.xml").decode()
    assert 'w:after="0" w:line="240"' in styles


def test_sliver_rows_dropped():
    from docx.shared import Pt as P
    d = Document()
    t = d.add_table(rows=3, cols=2)
    t.cell(0, 0).text = "Поз."
    t.rows[1].height = P(22.7)
    t.rows[2].height = P(2)            # щель до верхней линии штампа
    docxfix.fix_table_grids(d.element.body)
    assert len(t.rows) == 2


def test_table_grid_ignores_placeholder_spans():
    from docx.oxml.ns import qn
    d = Document()
    t = d.add_table(rows=4, cols=3)
    t.cell(0, 0).merge(t.cell(0, 1)).width = 300 * 635       # столбцы 0–1: 300
    t.cell(1, 0).merge(t.cell(1, 1)).width = 300 * 635
    t.cell(2, 0).merge(t.cell(2, 1)).width = 8692 * 635      # заглушка python-docx
    t.cell(3, 0).width = 160 * 635          # (3, 1) без ширины — у неё заглушка python-docx
    for r in range(4):
        t.cell(r, 2).width = 270 * 635
    docxfix.fix_table_grid(t._tbl)
    assert [int(g.get(qn("w:w"))) for g in t._tbl.iter(qn("w:gridCol"))] == [160, 140, 270]


def test_page_end_section_break_is_shrunk():
    from docx.enum.section import WD_SECTION
    d = Document()
    d.add_paragraph("лист 1, имя файла внизу")
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("лист 2")
    d.add_section(WD_SECTION.CONTINUOUS)
    d.add_paragraph("колонки")
    body = d.element.body
    breaks = [p for p in body if p.tag.endswith("}p") and p.find(f"{docxfix.w('pPr')}/{docxfix.w('sectPr')}") is not None]
    assert len(breaks) == 2
    docxfix.shrink_section_breaks(body)
    def line(p):
        sp = p.find(f"{docxfix.w('pPr')}/{docxfix.w('spacing')}")
        return None if sp is None else sp.get(docxfix.w("line"))
    assert line(breaks[0]) == "20"          # дальше новый лист — абзац-разрыв в 1 pt
    assert line(breaks[1]) is None          # дальше раздел на том же листе — не трогаем


def test_toc_common_right_edge_and_label_tabs():
    d = Document()
    for text, left, hang, right in (("2.1 Общая характеристика схемы ........ 14", 1148, 566, 1584),
                                    ("2.3 Отделение удаления отходов ........ 39", 582, 0, 0),
                                    ("3.1 Потребность в сырье ........ 43", 582, 0, 0)):
        p = d.add_paragraph(text)
        p.paragraph_format.left_indent = Pt(left / 20)
        p.paragraph_format.right_indent = Pt(right / 20)
        if hang:
            p.paragraph_format.first_line_indent = Pt(-hang / 20)
    assert docxfix.fix_toc(d) == 3
    for p in d.paragraphs:
        ind = p._p.find(f"{docxfix.w('pPr')}/{docxfix.w('ind')}")
        tabs = [(t.get(docxfix.w("val")), t.get(docxfix.w("pos"))) for t in p._p.iter(docxfix.w("tab"))
                if t.get(docxfix.w("pos"))]
        assert ind.get(docxfix.w("right")) == "1584"                 # у всех один правый край
        assert ind.get(docxfix.w("left")) == "1148" and ind.get(docxfix.w("hanging")) == "566"
        assert ("left", "1148") in tabs and tabs[-1][0] == "right"
        assert "\t" in docxfix._ptext(p._p).split(" ", 1)[0]          # «2.3<TAB>Отделение»


def test_compatibility_mode():
    d = Document()
    docxfix.set_compatibility_mode(d, 15)
    vals = [el.get(docxfix.w("val")) for el in d.settings.element.iter(docxfix.w("compatSetting"))
            if el.get(docxfix.w("name")) == "compatibilityMode"]
    assert vals == ["15"]
