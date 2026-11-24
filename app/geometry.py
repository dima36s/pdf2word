"""Геометрия векторной графики PDF без зависимостей от PyMuPDF и pdf2docx (тестируется отдельно).

Зачем: pdf2docx решает, что область страницы — «векторный рисунок», и вставляет её в Word картинкой
(а текст кладёт поверх), если в ней есть хоть один «непрямой» путь. Прямым он считает только
идеально горизонтальную/вертикальную линию или замкнутый контур, заполняющий свой прямоугольник.
На чертёжной документации из САПР это ломается:

* сетка таблиц и рамка рисуются ломаными (одна команда пути — несколько отрезков подряд), у ломаной
  площадь 0 при ненулевом габарите → «непрямая»;
* координаты отрезков чуть «плывут» (100.0 → 100.02) — горизонталь с микронаклоном тоже «непрямая»;
* таблица внутри рамки листа оказывается на втором уровне вложенности контуров, а такие области
  pdf2docx всегда считает «рисунком в ячейке» и растрирует целиком;
* логотип (кривые Безье) в шапке листа превращает в картинку всю рамку вместе со штампом.

Здесь — проверка «прямоты» по отрезкам и решение, какие пути области отдать в таблицы, а какие
вырезать картинкой (только компактные группы кривых).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Sequence

Box = tuple[float, float, float, float]
Pt = tuple[float, float]

AXIS_TOL_ABS = 0.75   # pt: допустимое отклонение конца отрезка от горизонтали/вертикали
AXIS_TOL_REL = 0.015  # доля длины (≈0.9°) — для длинных линий
POINT_LEN = 0.25      # отрезок короче — точка, на «прямоту» не влияет


# --------------------------------------------------------------------------- отрезки

def segment_axis(p0: Pt, p1: Pt) -> str | None:
    """'h' — горизонталь, 'v' — вертикаль, 'p' — вырожденный отрезок, None — наклонный."""
    dx, dy = abs(p1[0] - p0[0]), abs(p1[1] - p0[1])
    length = max(dx, dy)
    if length < POINT_LEN:
        return "p"
    tol = max(AXIS_TOL_ABS * 0.5, min(AXIS_TOL_ABS, AXIS_TOL_REL * length))
    if dy <= tol and dx > dy:
        return "h"
    if dx <= tol and dy > dx:
        return "v"
    return None


def snap_segment(p0: Pt, p1: Pt, axis: str) -> tuple[Pt, Pt]:
    """Выпрямить почти горизонтальный/вертикальный отрезок (pdf2docx требует точного совпадения)."""
    if axis == "h":
        y = (p0[1] + p1[1]) / 2.0
        return (p0[0], y), (p1[0], y)
    if axis == "v":
        x = (p0[0] + p1[0]) / 2.0
        return (x, p0[1]), (x, p1[1])
    return p0, p1


def curve_axis(points: Sequence[Pt]) -> str | None:
    """Кривая Безье, у которой все опорные точки лежат на одной горизонтали/вертикали, — это отрезок
    (так некоторые САПР пишут прямые). Возвращает ось или None для настоящей кривой."""
    p0, p3 = points[0], points[-1]
    axis = segment_axis(p0, p3)
    if axis is None:
        return None
    if axis == "p":
        # замкнутая петля из одной кривой (окружность-точка) — не отрезок
        span = max(max(abs(p[0] - p0[0]), abs(p[1] - p0[1])) for p in points)
        return "p" if span < POINT_LEN else None
    for p in points[1:-1]:
        if segment_axis(p0, p) not in (axis, "p") or segment_axis(p, p3) not in (axis, "p"):
            return None
    return axis


def _quad_axis_aligned(quad: Sequence[Pt]) -> bool:
    pts = list(quad)
    for a, b in zip(pts, pts[1:] + pts[:1]):
        if segment_axis(a, b) is None:
            return False
    return True


def _as_points(item) -> list[Pt]:
    return [(float(p[0]), float(p[1])) for p in item[1:]]


def items_axis_aligned(items: Iterable) -> bool:
    """Все команды пути (формат page.get_drawings: 'l', 'c', 're', 'qu') — горизонтали/вертикали."""
    for item in items:
        kind = item[0]
        if kind == "l":
            p0, p1 = _as_points(item)[:2]
            if segment_axis(p0, p1) is None:
                return False
        elif kind == "c":
            if curve_axis(_as_points(item)) is None:
                return False
        elif kind == "re":
            continue  # прямоугольник всегда ортогонален (повёрнутые MuPDF отдаёт как 'qu')
        elif kind == "qu":
            q = item[1]
            pts = [(q[0][0], q[0][1]), (q[1][0], q[1][1]), (q[3][0], q[3][1]), (q[2][0], q[2][1])] \
                if not hasattr(q, "ul") else [tuple(q.ul), tuple(q.ur), tuple(q.lr), tuple(q.ll)]
            if not _quad_axis_aligned(pts):
                return False
        else:
            return False
    return True


# --------------------------------------------------------------------------- прямоугольники

def area(b: Box) -> float:
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def union(a: Box, b: Box) -> Box:
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def contains(outer: Box, inner: Box, tol: float = 0.0) -> bool:
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def near(a: Box, b: Box, gap_x: float, gap_y: float) -> bool:
    return (a[0] - gap_x <= b[2] and b[0] - gap_x <= a[2]
            and a[1] - gap_y <= b[3] and b[1] - gap_y <= a[3])


def cluster_boxes(boxes: Iterable[Box], gap_x: float = 4.0, gap_y: float = 4.0) -> list[Box]:
    """Объединить близкие прямоугольники в группы (габарит каждой группы)."""
    clusters: list[Box] = []
    for b in boxes:
        cur = tuple(map(float, b))
        merged = True
        while merged:
            merged = False
            rest = []
            for c in clusters:
                if near(cur, c, gap_x, gap_y):
                    cur = union(cur, c)
                    merged = True
                else:
                    rest.append(c)
            clusters = rest
        clusters.append(cur)
    return clusters


# --------------------------------------------------------------------------- решение по области

@dataclass
class PathGeom:
    """Минимум сведений о пути для решения «таблица или рисунок»."""
    bbox: Box
    iso: bool          # все отрезки горизонтальные/вертикальные (или заливка-прямоугольник)


@dataclass
class RegionPlan:
    shapes: list[int]      # индексы путей → линии/заливки (таблицы, подчёркивания)
    images: list[Box]      # области → картинки


GRAPHIC_SHARE = 0.35       # кривые занимают больше этой доли области — это рисунок целиком
MIN_STRUCTURE_LINES = 4    # столько длинных прямых линий — признак таблицы/рамки
LONG_LINE_SHARE = 0.2      # «длинная» линия — не короче этой доли меньшей стороны области
SVG_GAP = 6.0              # pt: кривые ближе — одна картинка


def plan_region(region: Box, paths: Sequence[PathGeom]) -> RegionPlan:
    """Что сделать с путями области, найденной pdf2docx по контурам (внешняя область или область
    внутри ячейки — в оригинале pdf2docx вторые растрируются всегда, даже если там обычная таблица).

    * все пути прямые → линии (pdf2docx соберёт таблицу), картинок нет;
    * есть кривые, но это компактные вкрапления (логотип, значок) в структуре из прямых линий →
      прямые линии отдаются в таблицы, картинкой вырезаются только группы кривых;
    * кривые занимают заметную часть области или прямых линий мало (график, схема) →
      вся область — картинка, как в оригинале.
    """
    idx = list(range(len(paths)))
    curved = [i for i in idx if not paths[i].iso]
    straight = [i for i in idx if paths[i].iso]
    if not curved:
        return RegionPlan(shapes=straight, images=[])

    rw, rh = region[2] - region[0], region[3] - region[1]
    ref = max(min(rw, rh), 1.0)
    long_lines = sum(1 for i in straight
                     if max(paths[i].bbox[2] - paths[i].bbox[0], paths[i].bbox[3] - paths[i].bbox[1])
                     >= LONG_LINE_SHARE * ref)
    clusters = cluster_boxes((paths[i].bbox for i in curved), SVG_GAP, SVG_GAP)
    graphic_area = sum(area(c) for c in clusters)
    if long_lines < MIN_STRUCTURE_LINES or graphic_area >= GRAPHIC_SHARE * max(area(region), 1.0):
        return RegionPlan(shapes=[], images=[region])
    # прямые пути, целиком лежащие внутри картинки, не дублируем линиями
    keep = [i for i in straight if not any(contains(c, paths[i].bbox, 0.5) for c in clusters)]
    return RegionPlan(shapes=keep, images=clusters)
