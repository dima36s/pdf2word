"""Тесты ядра конвертации. Запуск: pytest -q tests/test_core.py"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from docx import Document

from app.converter import ConvertOptions, convert_pdf
from app.ocr import tesseract_available
from app.pdftools import ConversionError, classify_page, open_pdf, parse_page_range
from app.textfix import normalize_text_blocks

SAMPLES = Path(__file__).parent / "samples"


def _need(name: str) -> Path:
    path = SAMPLES / name
    if not path.exists():
        pytest.skip(f"нет {name}: сначала запустите python tests/make_samples.py")
    return path


def _text(docx: Path) -> str:
    d = Document(docx)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            parts.extend(c.text for c in row.cells)
    return "\n".join(parts)


# --------------------------------------------------------------------------- диапазоны

@pytest.mark.parametrize("spec,total,expected", [
    ("", 5, None),
    ("1-3", 5, [0, 1, 2]),
    ("1, 3,5", 5, [0, 2, 4]),
    ("4-", 6, [3, 4, 5]),
    ("-2", 6, [0, 1]),
    ("2 – 3", 5, [1, 2]),
    ("1-100", 3, [0, 1, 2]),
    ("3,1,3", 5, [0, 2]),
])
def test_page_range(spec, total, expected):
    assert parse_page_range(spec, total) == expected


@pytest.mark.parametrize("spec", ["abc", "0", "3-1", "9", "1;DROP"])
def test_page_range_errors(spec):
    with pytest.raises(ConversionError):
        parse_page_range(spec, 5)


# --------------------------------------------------------------------------- нормализация

def _char(c, x0, x1, y=100.0):
    return {"c": c, "origin": (x0, y), "bbox": (x0, y - 9, x1, y + 2)}


def _line(text, x0, y=100.0, w=6.0):
    chars = [_char(c, x0 + i * w, x0 + (i + 1) * w, y) for i, c in enumerate(text)]
    return {"dir": (1, 0), "wmode": 0, "bbox": (x0, y - 9, x0 + len(text) * w, y + 2),
            "spans": [{"font": "Arial", "size": 10, "chars": chars, "bbox": (x0, y - 9, x0 + len(text) * w, y + 2)}]}


def _text_of(line):
    return "".join(c["c"] for s in line["spans"] for c in s["chars"])


def test_justified_fragments_are_merged():
    block = {"type": 0, "lines": [_line("Цифровая ", 0), _line("трансформация", 120)]}
    out = normalize_text_blocks([block])
    assert len(out[0]["lines"]) == 1
    assert _text_of(out[0]["lines"][0]) == "Цифровая трансформация"


def test_table_cells_are_not_merged():
    # без пробела на стыке — это разные ячейки, а не одна строка текста
    block = {"type": 0, "lines": [_line("Итого", 0), _line("2 036 000", 200)]}
    assert len(normalize_text_blocks([block])[0]["lines"]) == 2


def test_cells_split_by_table_line_are_not_merged():
    # Word кончает текст ячейки пробелом («315 », «920х580 »): без учёта линий сетки
    # строки соседних ячеек (и соседних таблиц) склеивались и выпадали из таблицы
    from app.textfix import Barriers

    block = {"type": 0, "lines": [_line("315 ", 0), _line("920х580 ", 60), _line("110 ", 200)]}
    grid = Barriers([(40.0, 80, 120), (150.0, 80, 120)])
    assert len(normalize_text_blocks([block], grid)[0]["lines"]) == 3
    # линия выше строки (другая строка таблицы) — не мешает склейке
    block = {"type": 0, "lines": [_line("Цифровая ", 0), _line("трансформация", 120)]}
    assert len(normalize_text_blocks([block], Barriers([(80.0, 20, 60)]))[0]["lines"]) == 1


def test_dash_inside_sentence_is_not_list_item():
    from app.textfix import list_item_follows

    # перенос внутри фразы: строка выше дошла до правого края и не кончилась знаком
    assert not list_item_follows("включает в себя четыре нитки «ленточный сгуститель", 485.0, 486.0)
    # настоящие пункты: перед ними двоеточие, точка с запятой или короткая строка
    assert list_item_follows("состоит из следующих узлов:", 485.0, 486.0)
    assert list_item_follows("– насос подачи;", 120.0, 486.0)
    assert list_item_follows("– насос подачи", 120.0, 486.0)


def test_list_marker_joined_with_space():
    block = {"type": 0, "lines": [_line("1.", 0), _line("Анализ", 18)]}
    out = normalize_text_blocks([block])
    assert _text_of(out[0]["lines"][0]) == "1. Анализ"


def test_symbol_bullet_mapped():
    line = _line("", 0)
    line["spans"][0]["font"] = "OpenSymbol"
    out = normalize_text_blocks([{"type": 0, "lines": [line]}])
    assert _text_of(out[0]["lines"][0]) == "•"


def test_cell_text_with_nested_table():
    # вложенная таблица отдаёт text списком — проверка ложных таблиц не должна падать
    from types import SimpleNamespace
    from app.textfix import _cell_text
    cell = SimpleNamespace(blocks=[SimpleNamespace(text="Поз."), SimpleNamespace(text=[["1", "Насос"], ["2", None]])])
    assert _cell_text(cell) == "Поз. 1 Насос 2"


def test_line_end_space_between_wrapped_lines():
    block = {"type": 0, "lines": [_line("трансформация", 0, y=100), _line("производственных", 0, y=114)]}
    out = normalize_text_blocks([block])
    assert _text_of(out[0]["lines"][0]).endswith(" ")


# --------------------------------------------------------------------------- классификация

def test_classify_pages():
    assert classify_page(open_pdf(_need("report.pdf"))[0]) == "text"
    assert classify_page(open_pdf(_need("scan.pdf"))[0]) == "scan"


def test_password():
    path = _need("protected.pdf")
    with pytest.raises(ConversionError, match="паролем"):
        open_pdf(path)
    with pytest.raises(ConversionError, match="Неверный"):
        open_pdf(path, "wrong")
    doc = open_pdf(path, "1234")
    assert "пилотном" in doc[0].get_text()


# --------------------------------------------------------------------------- конвертация

def test_report(tmp_path):
    out = tmp_path / "r.docx"
    st = convert_pdf(_need("report.pdf"), out, ConvertOptions(title="Отчёт"))
    assert st.tables == 2 and st.images == 1 and not st.warnings
    d = Document(out)
    assert d.core_properties.title == "Отчёт"
    text = _text(out)
    for needle in ("Отчёт о пилотном проекте", "•\tСократить время", "1. Анализ текущего процесса",
                   "Удовлетворённость пользователей", "Жирный фрагмент"):
        assert needle in text           # после маркера списка — табуляция, как в Word
    doc_xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert re.search(r'<w:tabs><w:tab w:pos="\d+" w:val="left"/></w:tabs>', doc_xml)   # позиция табуляции
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert "Times New Roman" in xml and "Liberation Serif" not in xml
    assert 'w:hRule="exact"' not in xml.split("<w:trHeight")[1] if "<w:trHeight" in xml else True


def test_two_columns_no_false_tables(tmp_path):
    out = tmp_path / "c.docx"
    st = convert_pdf(_need("two_columns.pdf"), out, ConvertOptions())
    assert st.tables == 0
    text = _text(out)
    assert "трансформацияпроизводственных" not in text
    assert "Цифровая трансформация производственных процессов" in text
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert 'w:num="2"' in xml  # двухколоночная секция
    # колонки через разрыв колонки, а не через разделы nextColumn (их не понимают LibreOffice/R7)
    assert "nextColumn" not in xml and 'w:type="column"' in xml


def test_borderless_table(tmp_path):
    out = tmp_path / "b.docx"
    st = convert_pdf(_need("borderless_table.pdf"), out, ConvertOptions())
    assert st.tables == 1
    rows = Document(out).tables[0].rows
    assert len(rows) == 6


def test_page_subset(tmp_path):
    out = tmp_path / "p.docx"
    st = convert_pdf(_need("report.pdf"), out, ConvertOptions(pages="3"))
    assert st.pages_converted == 1
    text = _text(out)
    assert "План работ" in text and "Введение" not in text


def test_protected(tmp_path):
    out = tmp_path / "x.docx"
    st = convert_pdf(_need("protected.pdf"), out, ConvertOptions(password="1234"))
    assert st.tables == 2


def test_browser_panels(tmp_path):
    out = tmp_path / "i.docx"
    st = convert_pdf(_need("invoice_chrome.pdf"), out, ConvertOptions())
    text = _text(out)
    assert "Ноутбук" in text and "736 000,00" in text and "Покупатель" in text
    assert st.tables <= 3  # плашки не должны распадаться на десяток таблиц 1×1
    assert "2026 г.Срок" not in text  # пробел между строками разных блоков (кириллица)


def _part(docx: Path, prefix: str) -> str:
    z = zipfile.ZipFile(docx)
    return "\n".join(z.read(n).decode() for n in z.namelist() if n.startswith(prefix))


def test_report_headers_footers(tmp_path):
    # повторяющиеся шапка и низ страницы → настоящие колонтитулы Word
    out = tmp_path / "rh.docx"
    st = convert_pdf(_need("report.pdf"), out, ConvertOptions())
    assert st.headers >= 2
    assert "Внутренний отчёт" in _part(out, "word/header") and "Конфиденциально" in _part(out, "word/footer")
    assert "Внутренний отчёт" not in _text(out)


def test_gost_spec_sheet(tmp_path):
    """Спецификация с рамкой из САПР: таблица — таблицей, а не картинкой; штампы — в колонтитулах."""
    out = tmp_path / "spec.docx"
    st = convert_pdf(_need("gost_spec.pdf"), out, ConvertOptions())
    assert st.images <= 1                          # раньше весь лист становился одной картинкой
    assert st.tables >= 1
    body = _text(out)
    assert "Метран-150" in body and "Расходомер электромагнитный DN50" in body
    assert "Разраб." not in body and "Инв. № подл." not in body
    footer, header = _part(out, "word/footer"), _part(out, "word/header")
    assert "Спецификация оборудования" in footer and "Листов" in footer
    assert "Инв. № подл." in header and "w:txbxContent" in header and "v:rect" in header
    assert 'w:textDirection w:val="btLr"' in header or 'w:textDirection w:val="tbRl"' in header
    d = Document(out)
    assert len(d.sections) >= 2 and all(s.page_width > s.page_height for s in d.sections)


def test_gost_text_sheets(tmp_path):
    """Текстовая часть: шапка с логотипом и линией, штамп с номером листа, содержание, заголовки."""
    out = tmp_path / "text.docx"
    st = convert_pdf(_need("gost_text.pdf"), out, ConvertOptions())
    header, footer = _part(out, "word/header"), _part(out, "word/footer")
    assert "Е110-0162-УКК.25.1087-РД-01-03.03.086-АТХ.ОЛ" in header and "w:pBdr" in header  # линия под шапкой
    assert "Трубопровод технического конденсата" in footer and 'w:instr=" PAGE "' in footer
    body = _text(out)
    assert "ИНЖИНИРИНГ" not in body and "Трубопровод технического" not in body
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert xml.count('w:leader="dot"') >= 3         # строки содержания — табуляция с отточием
    assert 'w:start="2"' in xml                    # нумерация листов начинается со 2
    d = Document(out)
    styles = {p.style.name for p in d.paragraphs}
    assert "Heading 1" in styles
    assert d.sections[-1].page_width > d.sections[-1].page_height   # лист A3 альбомный
    assert st.headers >= 2


@pytest.mark.skipif(not tesseract_available(), reason="tesseract не установлен")
def test_scan_ocr(tmp_path):
    out = tmp_path / "s.docx"
    st = convert_pdf(_need("scan_color.pdf"), out, ConvertOptions(ocr_lang="rus+eng"))
    assert st.ocr_pages == 1 and st.tables == 1
    text = _text(out)
    for needle in ("Отчёт о пилотном проекте", "Цифровая трансформация", "Показатель", "Снизить число ошибок"):
        assert needle in text
    d = Document(out)
    bold = [r.text for p in d.paragraphs for r in p.runs if r.bold]
    assert any("Введение" in t for t in bold)


def _ocr_result(words, photos=()):
    from app.ocr import Line, OcrResult, Paragraph, Word

    ws = [Word(t, 10 * k, 10, 10 * k + 8, 20, conf) for k, (t, conf) in enumerate(words)]
    return OcrResult([Paragraph([Line((0, 10, 800, 20), 20, 8, ws)])] if ws else [], list(photos))


def test_picture_page_is_not_replaced_by_ocr_text():
    """Обложка-фотоколлаж: с картинки читаются 5–7 слов — страница остаётся изображением."""
    from app.ocr import looks_like_picture

    cover = _ocr_result([("пецСтрои", 80), ("каталог", 95), ("Дегазационные", 70), ("керамика", 75),
                         ("Абразивостойкая", 72), ("р", 30)], photos=[(0, 0, 800, 760)])
    assert looks_like_picture(cover, 800, 1130)
    assert looks_like_picture(_ocr_result([]), 800, 1130)
    page = [(f"слово{k}", 92) for k in range(110)]
    assert not looks_like_picture(_ocr_result(page), 800, 1130)
    # страница текста с большой иллюстрацией — всё равно текст
    assert not looks_like_picture(_ocr_result(page, photos=[(0, 0, 800, 700)]), 800, 1130)


def test_scan_without_ocr_keeps_image(tmp_path):
    out = tmp_path / "n.docx"
    st = convert_pdf(_need("scan.pdf"), out, ConvertOptions(ocr="off"))
    assert st.ocr_pages == 0 and st.images >= 2 and st.warnings


def test_toc_entry_ends_paragraph_and_size_jump():
    from app.textfix import size_jump, toc_entry_ends
    assert toc_entry_ends("обогащение в главный корпус второй очереди ................ 16 ")
    assert toc_entry_ends("3.7 Сжатый воздух ....................................... 53")
    assert not toc_entry_ends("Характеристика принятой технологической схемы производства в целом ")
    assert not toc_entry_ends("в 2024 г. было принято 16")
    assert size_jump(15.0, 12.0)          # «Содержание» 15 pt над пунктами 12 pt
    assert not size_jump(12.0, 11.6)      # «°С» в строке — тот же абзац
    assert not size_jump(0.0, 12.0)       # строка из картинки


def test_list_marker_gap_becomes_tab():
    from app.textfix import _try_merge, marker_gap_is_tab, marker_tab_stops

    def ch(c, x0, x1, y=100.0):
        return {"c": c, "origin": (x0, y), "bbox": (x0, y - 10, x1, y + 2)}

    def line(chars, size=12.0):
        return {"dir": (1, 0), "wmode": 0, "bbox": (chars[0]["bbox"][0], 90, chars[-1]["bbox"][2], 102),
                "spans": [{"size": size, "chars": chars, "bbox": (chars[0]["bbox"][0], 90, chars[-1]["bbox"][2], 102)}]}

    # «−» и пробел (Symbol), текст пункта — отдельная строка PDF на 18 pt правее: табуляция Word
    a = line([ch("−", 106.3, 112.9), ch(" ", 112.9, 116.3)])
    b = line([ch("к", 134.8, 141.0), ch("о", 141.0, 147.7)])
    assert _try_merge(a, b)
    text = "".join(c["c"] for s in a["spans"] for c in s["chars"])
    assert text == "−\tко"
    tab = next(c for s in a["spans"] for c in s["chars"] if c["c"] == "\t")
    assert tab["bbox"][2] == 134.8
    # обычный пробел после тире — не табуляция
    assert not marker_gap_is_tab(3.4, 12.0) and marker_gap_is_tab(21.9, 12.0)
    assert marker_tab_stops(70.9, 0.7, [134.8, 134.9]) == [64.6]


def test_squeezed_spaces_detected():
    from app.pdftools import space_ratios

    def ch(c, x0, w):
        return {"c": c, "origin": (x0, 0), "bbox": (x0, 0, x0 + w, 10)}

    # пробел 3.34 pt, а следующая буква через 2.7 pt — Word сжал пробел
    chars = []
    x = 0.0
    for _ in range(6):
        chars += [ch("а", x, 6.7), ch(" ", x + 6.7, 3.34)]
        x += 6.7 + 2.7
    chars.append(ch("б", x, 6.9))
    ratios = space_ratios(chars)
    assert len(ratios) == 6 and all(0.79 < r < 0.82 for r in ratios)


def test_merged_cell_vertical_align():
    from app.textfix import cell_vertical_align
    # «Лаборатория в пристройке АБК 8.1» посередине трёх строк таблицы
    assert cell_vertical_align((730, 239, 790, 507), (732, 353, 788, 392)) == "center"
    assert cell_vertical_align((730, 239, 790, 507), (732, 460, 788, 500)) == "bottom"
    assert cell_vertical_align((730, 239, 790, 507), (732, 242, 788, 280)) is None     # сверху — как есть
