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
