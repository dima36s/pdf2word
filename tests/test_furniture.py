"""Поиск оформления листа и фильтры областей — чистая логика, без PyMuPDF."""
from __future__ import annotations

from types import SimpleNamespace

from app import furniture as fu
from app import regionfix as rf

MM = fu.MM


def _line(x0, y0, x1, y1):
    return fu.Elem((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)), "path",
                   segs=[(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))])


def _box(x0, y0, x1, y1):
    segs = [(x0, y0, x1, y0), (x1, y0, x1, y1), (x0, y1, x1, y1), (x0, y0, x0, y1)]
    return fu.Elem((x0, y0, x1, y1), "path", segs=segs)


def _text(x0, y0, x1, y1, text, vertical=False):
    return fu.Elem((x0, y0, x1, y1), "text", text, vertical)


def _spec_page(sheet: int, big: bool) -> fu.PageElems:
    """Лист A3 по ГОСТ: рамка, таблица спецификации до штампа, штамп 185 мм, боковой штамп."""
    W, H = 420 * MM, 297 * MM
    fx0, fy0, fx1, fy1 = 20 * MM, 5 * MM, W - 5 * MM, H - 5 * MM
    stamp_top = fy1 - (40 if big else 15) * MM
    el = [_box(fx0, fy0, fx1, fy1)]
    xs = [fx0 + k * 60 * MM for k in range(1, 7)]
    el += [_line(x, fy0, x, stamp_top) for x in xs]
    el += [_line(fx0, y, fx1, y) for y in (fy0 + 15 * MM, fy0 + 23 * MM, fy0 + 31 * MM)]
    el += [_text(fx0 + 5, fy0 + 17 * MM, fx0 + 200, fy0 + 20 * MM, "Датчик давления")]
    el += [_line(fx1 - 185 * MM, stamp_top, fx1, stamp_top), _line(fx1 - 185 * MM, stamp_top, fx1 - 185 * MM, fy1)]
    el += [_text(fx1 - 150 * MM, stamp_top + 5, fx1 - 60 * MM, stamp_top + 15, "Е110-0162-АТХ.СО"),
           _text(fx1 - 8 * MM, fy1 - 10, fx1 - 4 * MM, fy1 - 2, str(sheet))]
    for y0, y1, lab in ((H - 30 * MM, H - 5 * MM, "Инв. № подл."), (H - 65 * MM, H - 30 * MM, "Подп. и дата")):
        el += [_box(8 * MM, y0, 20 * MM, y1), _text(9 * MM, y0 + 5, 12 * MM, y1 - 5, lab, vertical=True)]
    return fu.PageElems(W, H, el)


def _text_page(sheet: int) -> fu.PageElems:
    """Текстовый лист: логотип и шифр над линией, нижний штамп-рамка, имя файла под ним."""
    W, H = 595.3, 841.9
    el = [fu.Elem((57, 22, 80, 42), "graphic"), _text(85, 26, 140, 38, "ПроТех"),
          _text(400, 26, 566, 34, "Е110-0162-АТХ.ОЛ"), _line(57, 47, 566, 47)]
    el += [_text(71, 90 + 16 * k, 540, 102 + 16 * k, f"Строка основного текста {k}") for k in range(20)]
    el += [_box(57, 773, 566, 801), _line(102, 773, 102, 801), _line(532, 773, 532, 801),
           _text(65, 782, 90, 792, "2026"), _text(110, 778, 500, 796, "Трубопровод технического конденсата"),
           _text(545, 782, 552, 792, str(sheet)), _text(400, 805, 566, 812, "Е110-0162-АТХ.ОЛ_0_0_RU_IFC.docx")]
    return fu.PageElems(W, H, el)


def test_gost_spec_sheet():
    first, second = fu.detect([_spec_page(1, True), _spec_page(2, False)])
    for f, height in ((first, 40), (second, 15)):
        assert f.frame is not None and f.header is None
        # штамп с подписями под рамкой — до низа листа
        assert f.footer is not None and abs(f.footer[1] - (297 - 5 - height) * MM) < 2 and f.footer[3] == 297 * MM
        assert len(f.sides) == 1 and f.sides[0][2] <= 20 * MM + 2
        assert f.drop      # пустая сетка левее штампа убирается из текста страницы


def test_single_spec_sheet():
    (f,) = fu.detect([_spec_page(1, True)])
    assert f.footer is not None and f.sides


def test_text_sheets_header_and_footer():
    pages = [_text_page(k) for k in (2, 3, 4)]
    for f in fu.detect(pages):
        assert f.header is not None and 47 <= f.header[3] <= 49
        assert f.footer is not None and 771 <= f.footer[1] <= 773 and f.footer[3] >= 812
    (single,) = fu.detect([_text_page(2)])
    assert single.header is not None and single.footer is not None


def test_title_with_rule_is_not_header():
    # заголовок документа с чертой под ним и без штампа — не шапка
    W, H = 595.3, 841.9
    el = [_text(150, 60, 450, 80, "Отчёт о пилотном проекте"), _line(57, 90, 540, 90)]
    el += [_text(71, 110 + 16 * k, 540, 122 + 16 * k, f"Абзац {k}") for k in range(30)]
    (f,) = fu.detect([fu.PageElems(W, H, el)])
    assert f.header is None and f.footer is None and f.empty


def test_repeated_header_and_page_numbers():
    W, H = 612, 792
    pages = []
    for k in range(3):
        el = [_text(72, 38, 300, 50, "ООО «Пример» · Внутренний отчёт"),
              _text(290, 744, 320, 756, f"Стр. {k + 1}")]
        el += [_text(72, 100 + 20 * j, 540, 112 + 20 * j, f"Текст {k} {j}") for j in range(10 + k)]
        pages.append(fu.PageElems(W, H, el))
    for f in fu.detect(pages):
        assert f.header == (0.0, 0.0, W, 51.0)
        assert f.footer is not None and abs(f.footer[1] - 743) < 0.01


# --------------------------------------------------------------------------- фильтры областей

def test_clip_stroke():
    body = rf.Spec(exclude=[(0, 700, 600, 842)])
    assert rf.clip_stroke(body, (100, 100), (100, 800)) == [((100, 100), (100, 700))]
    assert rf.clip_stroke(body, (50, 750), (500, 750)) == []
    band = rf.Spec(keep=[(300, 700, 600, 842)])
    assert rf.clip_stroke(band, (50, 750), (500, 750)) == [((300, 750), (500, 750))]
    # вертикаль таблицы над штампом: от неё в области штампа остался бы «огрызок» в 0.75 pt
    assert rf.clip_stroke(rf.Spec(keep=[(300, 699.25, 600, 842)]), (350, 100), (350, 700)) == []
    assert rf.clip_stroke(rf.Spec(keep=[(0, 0, 600, 842)]), (350, 100), (350, 101)) == [((350, 100), (350, 101))]


def test_frame_side_kept_only_when_table_uses_it():
    frame = (57, 14, 1176, 828)
    page = SimpleNamespace(parent=SimpleNamespace(), number=0)
    top = {"start": (57, 14), "end": (1176, 14), "width": 1, "color": 0}
    left = {"start": (57, 14), "end": (57, 828), "width": 1, "color": 0}
    column = {"start": (300, 14), "end": (300, 600), "width": 1, "color": 0}   # вертикаль таблицы упирается в верх рамки
    with rf.regions(page.parent, {0: rf.Spec(frame=frame)}):
        shapes, _ = rf.filter_shapes(page, [top, left, column], [])
    starts = sorted(tuple(s["start"]) + tuple(s["end"]) for s in shapes)
    assert (57, 14, 1176, 14) in starts and (300, 14, 300, 600) in starts
    assert (57, 14, 57, 828) not in starts       # свободная сторона рамки не превращается в таблицу


def test_filter_text_by_region():
    page = SimpleNamespace(parent=SimpleNamespace(), number=3)

    def ch(c, x):
        return {"c": c, "bbox": (x, 790, x + 5, 800), "origin": (x, 798)}
    block = {"type": 0, "bbox": (10, 100, 300, 800), "lines": [
        {"bbox": (10, 100, 60, 110), "spans": [{"bbox": (10, 100, 60, 110),
                                                 "chars": [{"c": "Т", "bbox": (10, 100, 15, 110), "origin": (10, 108)}]}]},
        {"bbox": (10, 790, 300, 800), "spans": [{"bbox": (10, 790, 300, 800), "chars": [ch("1", 10), ch("2", 290)]}]}]}
    with rf.regions(page.parent, {3: rf.Spec(exclude=[(0, 780, 200, 842)])}):
        out = rf.filter_text(page, [block])
    lines = out[0]["lines"]
    assert [c["c"] for ln in lines for s in ln["spans"] for c in s["chars"]] == ["Т", "2"]
    assert lines[1]["bbox"][0] == 290
    assert rf.filter_text(page, [block]) == [block]          # вне контекста фильтр не действует


def test_word_footer_strip_without_sides():
    # штамп из Word: верх и низ таблички и внутренние перегородки, боковых сторон нет; линии кусками
    W, H = 595.3, 841.9
    el = [fu.Elem((57, 22, 80, 42), "graphic"), _text(400, 26, 566, 34, "Е110-0162-АТХ.ОЛ"),
          _line(57, 47, 300, 47), _line(301, 47, 566, 47)]
    el += [_text(71, 90 + 16 * k, 540, 102 + 16 * k, f"Строка {k}") for k in range(20)]
    el += [_line(71, 753.5, 113, 753.5), _line(114, 753.5, 530, 753.5), _line(530, 753.5, 558.6, 753.5),
           _line(71, 780, 113, 780), _line(114, 780, 530, 780), _line(530, 780, 558.6, 780),
           _line(113.4, 753.5, 113.4, 780), _line(529.9, 753.5, 529.9, 780),
           _text(81, 761, 106, 772, "2026"), _text(116, 755, 500, 767, "Трубопровод"), _text(542, 761, 550, 772, "1")]
    (f,) = fu.detect([fu.PageElems(W, H, el)])
    assert f.footer is not None and abs(f.footer[1] - 752.75) < 0.1
    assert f.header is not None and abs(f.header[3] - 47.75) < 0.1


def test_last_table_row_is_not_footer():
    W, H = 595.3, 841.9
    el = [_text(71, 90 + 16 * k, 540, 102 + 16 * k, f"Строка {k}") for k in range(20)]
    ys = [700, 726.5, 753.5, 780]
    for a, b in zip(ys, ys[1:]):          # таблица: каждая строка рисуется своими кусками
        el += [_line(71, a, 558, a), _line(71, b, 558, b), _line(113, a, 113, b), _line(530, a, 530, b)]
    (f,) = fu.detect([fu.PageElems(W, H, el)])
    assert f.footer is None

def test_word_fill_lines_are_clipped_and_frame_pieces_dropped():
    frame = (77, 18, 1167, 808)
    page = SimpleNamespace(parent=SimpleNamespace(), number=1)
    # Word: линии — тонкие залитые прямоугольники, рамка — кусками
    left_piece = {"bbox": [76.3, 100, 77.7, 300], "color": 0}           # кусок левой стороны рамки
    row = {"bbox": [77, 399.8, 700, 400.3], "color": 0}                  # строка таблицы, упирается в рамку
    column = {"bbox": [299.3, 18, 300.7, 790], "color": 0}               # вертикаль, уходит в штамп
    shading = {"bbox": [100, 100, 200, 150], "color": 0xEEEEEE}          # заливка ячейки — не линия
    spec = rf.Spec(exclude=[(643, 763, 1191, 842), (76, 765, 644, 809)], frame=frame)
    with rf.regions(page.parent, {1: spec}):
        shapes, _ = rf.filter_shapes(page, [left_piece, row, column, shading], [])
    boxes = [tuple(round(v, 1) for v in sh["bbox"]) for sh in shapes]
    assert (299.3, 18, 300.7, 765) in boxes                                 # обрезана по полосе штампа
    assert any(b[1] == 399.8 for b in boxes) and (100, 100, 200, 150) in boxes
    assert not any(b[0] == 76.3 for b in boxes)    # кусок рамки к таблице не примыкает — убран

def test_caption_below_frame_goes_to_footer():
    # титульный лист: рамка без штампа, под рамкой — имя файла
    W, H = 595.3, 841.9
    el = [_box(71, 23, 561, 798), _text(150, 300, 450, 320, "РАБОЧАЯ ДОКУМЕНТАЦИЯ"),
          _text(150, 760, 300, 775, "Санкт-Петербург"), _text(250, 805, 560, 815, "Е110-0162-АТХ.ОЛ_0_0_RU_IFC.docm")]
    (f,) = fu.detect([fu.PageElems(W, H, el)])
    assert f.footer is not None and f.footer[1] > 798 and f.footer[3] == H


def _catalog_page(sheet: int, header_row: bool) -> fu.PageElems:
    """Каталог из Word: картинка-плашка с названием вверху, полоса «стр. N» внизу, без рамки.
    На листе с header_row таблица начинается строкой заголовков между двумя линиями."""
    W, H = 595.3, 841.9
    el = [fu.Elem((14, 12, 567, 48), "graphic"), _text(330, 16, 560, 26, "Труба футерованная"),
          _text(300, 30, 560, 40, "ТУ 2296-002-26612968-2000"),
          fu.Elem((1, 795, 594, 835), "graphic"), _text(510, 808, 545, 818, f"стр. {sheet}")]
    if header_row:
        el += [_line(40, 60, 555, 60), _line(40, 76, 555, 76), _line(40, 60, 40, 76), _line(330, 60, 330, 76),
               _line(555, 60, 555, 76), _text(45, 63, 200, 73, "Крепеж для Ду 1,0МПа"),
               _text(340, 63, 400, 73, "Ширина бурта, мм")]
    for k in range(18 - sheet):
        y = 90 + 7 * sheet + 26 * k
        el += [_line(40, y, 555, y), _text(45, y + 4, 200, y + 14, f"Болт М{20 + k + 30 * sheet}")]
    return fu.PageElems(W, H, el)


def test_table_header_row_is_not_page_header():
    """Повторяющийся низ листа (не штамп) не делает лист «бланком»: строка заголовков таблицы
    под плашкой — часть таблицы, а не шапка, даже если над ней логотип и линия."""
    pages = [_catalog_page(n, n == 3) for n in range(1, 6)]
    result = fu.detect(pages)
    f = result[2]
    assert f.footer is not None and f.footer[1] > 780          # «стр. N» — в нижнем колонтитуле
    assert f.header is not None and f.header[3] < 58           # шапка — только плашка
