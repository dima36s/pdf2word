"""Геометрия векторной графики (app/geometry.py) — без PyMuPDF и pdf2docx."""
from __future__ import annotations

import pytest

from app import geometry as geo


@pytest.mark.parametrize("p0,p1,axis", [
    ((0, 100), (300, 100), "h"),
    ((0, 100), (300, 100.04), "h"),        # микронаклон из САПР
    ((10, 0), (10.3, 500), "v"),
    ((0, 0), (0.1, 0.1), "p"),
    ((0, 0), (100, 30), None),             # настоящий наклон
    ((0, 0), (20, 1.0), None),             # короткая линия с заметным наклоном
])
def test_segment_axis(p0, p1, axis):
    assert geo.segment_axis(p0, p1) == axis


def test_snap_segment():
    (a, b) = geo.snap_segment((0, 100), (300, 100.04), "h")
    assert a[1] == b[1] == pytest.approx(100.02) and (a[0], b[0]) == (0, 300)
    (a, b) = geo.snap_segment((10, 0), (10.2, 50), "v")
    assert a[0] == b[0] == pytest.approx(10.1) and (a[1], b[1]) == (0, 50)


def test_polyline_is_axis_aligned():
    # рамка листа одной ломаной: pdf2docx считал её «рисунком» (площадь 0 при ненулевом габарите)
    frame = [("l", (57, 14), (581, 14)), ("l", (581, 14), (581, 828.02)),
             ("l", (581, 828.02), (57, 828)), ("l", (57, 828), (57, 14))]
    assert geo.items_axis_aligned(frame)


def test_curves_are_not_axis_aligned():
    logo = [("c", (0, 0), (5, 10), (15, 10), (20, 0))]
    assert not geo.items_axis_aligned(logo)
    straight_curve = [("c", (0, 5), (7, 5), (14, 5.01), (20, 5))]   # «кривая» на одной горизонтали
    assert geo.items_axis_aligned(straight_curve)
    assert not geo.items_axis_aligned([("l", (0, 0), (50, 50))])


def test_quads():
    assert geo.items_axis_aligned([("qu", ((0, 0), (10, 0), (0, 5), (10, 5)))])
    assert not geo.items_axis_aligned([("qu", ((0, 0), (10, 3), (-2, 5), (8, 8)))])


def test_cluster_boxes():
    boxes = [(0, 0, 10, 10), (12, 0, 20, 10), (100, 100, 110, 110)]
    clusters = sorted(geo.cluster_boxes(boxes, 4, 4))
    assert clusters == [(0, 0, 20, 10), (100, 100, 110, 110)]


def _grid(x0, y0, x1, y1, cols, rows):
    """Пути таблицы: горизонтали и вертикали."""
    paths = []
    for k in range(rows + 1):
        y = y0 + (y1 - y0) * k / rows
        paths.append(geo.PathGeom((x0, y - 0.5, x1, y + 0.5), True))
    for k in range(cols + 1):
        x = x0 + (x1 - x0) * k / cols
        paths.append(geo.PathGeom((x - 0.5, y0, x + 0.5, y1), True))
    return paths


def test_plan_all_straight_is_table():
    paths = _grid(50, 50, 550, 300, 5, 8)
    plan = geo.plan_region((50, 50, 550, 300), paths)
    assert plan.images == [] and plan.shapes == list(range(len(paths)))


def test_plan_logo_in_frame_keeps_table():
    # рамка листа со штампом и логотипом: картинкой должен стать только логотип
    paths = _grid(57, 14, 581, 828, 6, 20)
    paths += [geo.PathGeom((70, 20, 110, 40), False), geo.PathGeom((112, 22, 150, 38), False)]
    plan = geo.plan_region((57, 14, 581, 828), paths)
    assert len(plan.images) == 1
    x0, y0, x1, y1 = plan.images[0]
    assert (x0, y0, x1, y1) == (70, 20, 150, 40)
    assert len(plan.shapes) == len(paths) - 2


def test_plan_chart_stays_image():
    # график: кривая на всю область и пара осей — растрируется целиком, как раньше
    region = (100, 100, 400, 300)
    paths = [geo.PathGeom((100, 299, 400, 300), True), geo.PathGeom((100, 100, 101, 300), True),
             geo.PathGeom((110, 120, 390, 290), False)]
    plan = geo.plan_region(region, paths)
    assert plan.images == [region] and plan.shapes == []


def test_shading_rows_get_hidden_borders():
    # «полосатая» таблица из Word: вертикали на всю высоту, горизонтали только вокруг групп строк
    v = [(100, 50, 300), (200, 50, 300), (300, 50, 300)]
    h = [(50, 100, 300), (300, 100, 300), (150, 200, 300)]
    fills = [(200, 70, 300, 90), (200, 110, 300, 130), (200, 150.2, 300, 170)]
    rows = geo.shading_row_borders(fills, h, v)
    assert [round(y) for y, _, _ in rows] == [70, 90, 110, 130, 170]      # 150 — уже линия
    assert all((x0, x1) == (200, 300) for _, x0, x1 in rows)


def test_shading_without_grid_is_not_a_row():
    # заливка-плашка вне таблицы и заливка без правой вертикали — не трогаем
    v = [(100, 50, 300)]
    assert geo.shading_row_borders([(100, 70, 300, 90)], [], v) == []
    assert geo.shading_row_borders([(400, 70, 500, 90)], [], v) == []
    # тонкая заливка — это линия
    assert geo.shading_row_borders([(100, 70, 300, 71)], [], [(100, 0, 400), (300, 0, 400)]) == []


def test_white_margins_trimmed_from_picture():
    import io

    from PIL import Image, ImageDraw

    from app.imagetrim import trim_white_margins

    img = Image.new("RGB", (200, 100), "white")
    ImageDraw.Draw(img).rectangle((20, 10, 139, 89), fill=(30, 30, 30))   # чертёж, справа белое поле
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    raw = {"type": 1, "bbox": (100.0, 200.0, 300.0, 300.0), "width": 200, "height": 100, "image": buf.getvalue()}
    out = trim_white_margins(raw)
    x0, y0, x1, y1 = out["bbox"]
    assert x0 == pytest.approx(119, abs=1.5) and x1 == pytest.approx(240, abs=1.5)
    assert y0 == pytest.approx(209, abs=1.5) and y1 == pytest.approx(290, abs=1.5)
    assert Image.open(io.BytesIO(out["image"])).size == (out["width"], out["height"])
    # без заметных полей — без изменений
    full = Image.new("RGB", (100, 100), (200, 0, 0))
    buf = io.BytesIO()
    full.save(buf, format="PNG")
    raw = {"bbox": (0.0, 0.0, 50.0, 50.0), "width": 100, "height": 100, "image": buf.getvalue()}
    assert trim_white_margins(raw) is raw


def test_group_row_without_verticals_is_bridged():
    # таблица «Крепеж»: части сетки, между ними строка «Ду215  40  16» без вертикалей
    xs = [58, 146, 319, 383, 468, 546]
    parts = [(60, 110), (126, 176)]          # просвет 110…126 — строка подзаголовка
    v = [(x, a, b) for a, b in parts for x in xs]
    h = [(y, 58, 546) for a, b in parts for y in (a, b)] + [(85, 58, 546), (151, 58, 546)]
    bridges = geo.band_bridges(h, v)
    assert sorted(round(x) for x, _, _ in bridges) == xs
    assert all((round(y0), round(y1)) == (110, 126) for _, y0, y1 in bridges)


def test_separate_tables_are_not_bridged():
    xs = [58, 146, 319]
    # просвет больше строки (подпись «Таблица 2» с отступами) — две разные таблицы
    v = [(x, 60, 110) for x in xs] + [(x, 160, 200) for x in xs]
    h = [(y, 58, 319) for y in (60, 110, 160, 200)]
    assert geo.band_bridges(h, v) == []
    # две вертикали — мало для строки таблицы; нет горизонтали снизу — не закрытая строка
    v = [(x, 60, 110) for x in xs[:2]] + [(x, 126, 176) for x in xs[:2]]
    assert geo.band_bridges([(y, 58, 146) for y in (60, 110, 126, 176)], v) == []
    v = [(x, 60, 110) for x in xs] + [(x, 126, 176) for x in xs]
    assert geo.band_bridges([(y, 58, 319) for y in (60, 110, 176)], v) == []


def test_full_width_group_title_is_not_bridged():
    # «8.1.4. Перегрузочный узел …» по центру через все графы — объединённая ячейка, не делим
    xs = [58, 146, 319, 383, 468, 546]
    parts = [(60, 110), (126, 176)]
    v = [(x, a, b) for a, b in parts for x in xs]
    h = [(y, 58, 546) for a, b in parts for y in (a, b)]
    title = [(150.0, 113.0, 420.0, 123.0)]
    assert geo.band_bridges(h, v, texts=title) == []
    # подписи групп в своих графах («Ду215», «40», «16») — делим
    cells = [(180.0, 113.0, 210.0, 123.0), (400.0, 113.0, 412.0, 123.0), (500.0, 113.0, 510.0, 123.0)]
    assert len(geo.band_bridges(h, v, texts=cells)) == len(xs)


def test_group_title_with_values_keeps_value_columns():
    xs = [58, 146, 319, 383, 468, 546]
    parts = [(60, 110), (126, 176)]
    v = [(x, a, b) for a, b in parts for x in xs]
    h = [(y, 58, 546) for a, b in parts for y in (a, b)]
    # каталог: «Ду215» по центру своей графы, «40» и «16» — в своих: вся сетка
    texts = [(216.8, 113.0, 248.8, 123.0), (419.2, 113.0, 431.3, 123.0), (500.8, 113.0, 512.9, 123.0)]
    assert sorted({round(b[0]) for b in geo.band_bridges(h, v, texts=texts)}) == xs
    # заголовок через первые графы и значения в своих: первая граница не проводится
    texts = [(100.0, 113.0, 300.0, 123.0), (420.0, 113.0, 430.0, 123.0), (500.0, 113.0, 510.0, 123.0)]
    assert sorted({round(b[0]) for b in geo.band_bridges(h, v, texts=texts)}) == [58, 319, 383, 468, 546]


def test_text_chunks_split_on_column_gaps():
    spans = [((216.8, 104, 248.8, 116), 12.0), ((419.2, 104, 431.3, 116), 12.0), ((500.8, 104, 512.9, 116), 12.0)]
    assert len(geo.text_chunks(spans)) == 3
    words = [((250, 104, 300, 116), 11.0), ((303, 104, 400, 116), 11.0)]       # «8.1.4.» + « Перегрузочный…»
    assert geo.text_chunks(words) == [(250, 104, 400, 116)]
