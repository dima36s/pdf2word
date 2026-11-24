"""Выгрузка в Excel (app/xlsxexport.py): структура таблиц, числа, табуляция, листы по страницам."""
from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from app import xlsxexport as xe  # noqa: E402
from app.converter import ConvertOptions, convert_pdf  # noqa: E402

SAMPLES = Path(__file__).parent / "samples"


def test_number_detection():
    assert xe.to_value("50") == (50, "0")
    assert xe.to_value("3,0") == (3.0, "0.0")
    assert xe.to_value("105,25") == (105.25, "0.00")
    # коды, диапазоны и размеры остаются текстом
    for text in ("0162", "400-8000", "1100х600", "М20х140", "3,2-10,0", "10.2.1"):
        assert xe.to_value(text) == (text, None)


def _catalog_docx(path: Path) -> None:
    """Таблица как у pdf2docx: объединение по вертикали и горизонтали, заливка, границы «4.0»."""
    import docx
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    d = docx.Document()
    d.add_paragraph("Минимальная толщина стенки\tТаблица 1")
    t = d.add_table(rows=3, cols=3)
    t.cell(0, 0).merge(t.cell(1, 0)).text = "Диаметр, мм"
    t.cell(0, 1).merge(t.cell(0, 2)).text = "Толщина при Р, МПа"
    t.cell(1, 1).text, t.cell(1, 2).text = "1,0", "1,6"
    t.cell(2, 0).text, t.cell(2, 1).text, t.cell(2, 2).text = "50", "3,0", "400-8000"
    for row in t.rows:
        for cell in row.cells:
            tcpr = cell._tc.get_or_add_tcPr()
            tcpr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4.0" w:color="#000000"/>'
                                  '<w:bottom w:val="single" w:sz="4.0" w:color="#000000"/></w:tcBorders>'))
    t.cell(1, 1)._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="dfdf35"/>'))
    d.save(path)


def test_table_structure(tmp_path):
    src, out = tmp_path / "t.docx", tmp_path / "t.xlsx"
    _catalog_docx(src)
    assert xe.docx_to_xlsx(src, out, [7]) == 1
    ws = openpyxl.load_workbook(out)["Стр. 7"]
    assert ws["A1"].value == "Минимальная толщина стенки"
    assert any(ws.cell(row=1, column=c).value == "Таблица 1" for c in range(2, 12))   # по табуляции
    merged = {str(r) for r in ws.merged_cells.ranges}
    assert {"A2:A3", "B2:C2"} <= merged
    assert ws["A2"].value == "Диаметр, мм" and ws["B3"].value == 1.0 and ws["B3"].number_format == "0.0"
    assert ws["A4"].value == 50 and ws["C4"].value == "400-8000"
    assert ws["B3"].fill.start_color.rgb.endswith("DFDF35")
    assert ws["A4"].border.top.style == "thin"


def test_report_to_excel(tmp_path):
    out, xlsx = tmp_path / "r.docx", tmp_path / "r.xlsx"
    st = convert_pdf(SAMPLES / "report.pdf", out, ConvertOptions(), xlsx_dst=xlsx)
    assert st.xlsx_sheets == 3 and xlsx.is_file()
    wb = openpyxl.load_workbook(xlsx)
    assert wb.sheetnames == ["Стр. 1", "Стр. 2", "Стр. 3"]
    values = [c.value for ws in wb for row in ws.iter_rows() for c in row if c.value is not None]
    assert any(isinstance(v, str) and "Отчёт о пилотном проекте" in v for v in values)
    assert any(isinstance(v, (int, float)) for v in values)          # числа из таблиц — числами


def test_export_failure_keeps_docx(tmp_path, monkeypatch):
    """Сбой выгрузки в Excel не роняет конвертацию: Word готов, Excel просто нет."""
    from app import converter

    def boom(*_a, **_k):
        raise RuntimeError("xlsx broke")

    monkeypatch.setattr(xe, "docx_to_xlsx", boom)
    out, xlsx = tmp_path / "r.docx", tmp_path / "r.xlsx"
    st = converter.convert_pdf(SAMPLES / "report.pdf", out, ConvertOptions(), xlsx_dst=xlsx)
    assert out.is_file() and not xlsx.exists() and st.xlsx_sheets == 0
