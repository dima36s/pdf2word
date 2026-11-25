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


# --------------------------------------------------------------------------- строки по заливке

SHADING_EDGE_TOL = 1.5     # pt: край заливки и линия сетки совпадают с этой точностью
SHADING_COVER = 0.75       # такая доля стороны заливки должна быть покрыта линией
SHADING_MIN = 3.0          # pt: заливка тоньше — это линия, а не ячейка

HLine = tuple[float, float, float]   # (y, x0, x1)
VLine = tuple[float, float, float]   # (x, y0, y1)


def _covered(pos: float, lo: float, hi: float, lines: Sequence[tuple[float, float, float]],
             tol: float) -> float:
    """Доля отрезка [lo, hi] на уровне pos, покрытая линиями (pos, a, b)."""
    spans = sorted((max(a, lo), min(b, hi)) for p, a, b in lines if abs(p - pos) <= tol and b > lo and a < hi)
    total, reach = 0.0, lo
    for a, b in spans:
        if b <= reach:
            continue
        total += b - max(a, reach)
        reach = b
    return total / max(hi - lo, 1e-6)


def shading_row_borders(fills: Sequence[Box], h_lines: Sequence[HLine], v_lines: Sequence[VLine],
                        tol: float = SHADING_EDGE_TOL) -> list[HLine]:
    """Невидимые границы строк таблицы по краям заливки ячеек.

    Word рисует «полосатые» таблицы без линий между строками: строки различаются только серой
    заливкой через одну. pdf2docx строит сетку таблицы только по линиям, поэтому такие строки
    сливались в одну огромную ячейку, а текст из неё рассыпался абзацами под таблицей.

    Заливка считается ячейкой, если слева и справа от неё идут вертикальные линии сетки; верх и низ
    такой заливки без горизонтальной линии становятся границами строк (без отрисовки).
    """
    out: list[HLine] = []
    seen: set[tuple[float, float, float]] = set()
    for x0, y0, x1, y1 in fills:
        if x1 - x0 < SHADING_MIN or y1 - y0 < SHADING_MIN:
            continue
        if (_covered(x0, y0, y1, v_lines, tol) < SHADING_COVER
                or _covered(x1, y0, y1, v_lines, tol) < SHADING_COVER):
            continue
        for y in (y0, y1):
            if _covered(y, x0, x1, h_lines, tol) >= SHADING_COVER:
                continue
            key = (round(y, 1), round(x0, 1), round(x1, 1))
            if key not in seen:
                seen.add(key)
                out.append((y, x0, x1))
    return out


# --------------------------------------------------------------------------- строка без вертикалей

BAND_MAX_GAP = 20.0        # pt: строка-подзаголовок группы («Ду215  40  16») — одна строка текста
BAND_MIN_COLUMNS = 3       # столько вертикалей должно прерваться на одной и той же строке


def band_bridges(h_lines: Sequence[HLine], v_lines: Sequence[VLine],
                 max_gap: float = BAND_MAX_GAP, tol: float = 1.0,
                 texts: Sequence[Box] = ()) -> list[VLine]:
    """Невидимые вертикали через строку таблицы, нарисованную без вертикальных линий.

    В каталогах Word строка-подзаголовок группы («Ду215», «Ду265»…) идёт на всю ширину без
    перегородок. pdf2docx видел вместо одной таблицы десяток: части сетки между подзаголовками,
    а подзаголовки — абзацами между ними. Если на одной высоте прерываются сразу несколько
    вертикалей, а сверху и снизу просвет закрыт горизонталями на всю ширину, — это строка той же
    таблицы: вертикали продолжаются через неё (невидимыми).

    Если текст строки пересекает место такой вертикали («8.1.4. Перегрузочный узел тракта подачи
    руды» по центру через все графы), это объединённая ячейка-заголовок — её не делим.
    """
    by_band: dict[tuple[int, int], list[VLine]] = {}
    by_x: dict[int, list[VLine]] = {}
    for ln in v_lines:
        by_x.setdefault(round(ln[0]), []).append(ln)
    for x, y0, y1 in v_lines:
        near = [b for k in (round(x) - 1, round(x), round(x) + 1) for b in by_x.get(k, ()) if abs(b[0] - x) <= tol]
        nxt = [b for b in near if y1 + 2 < b[1] <= y1 + max_gap]
        if not nxt:
            continue
        start = min(b[1] for b in nxt)
        if any(b[1] < start - 0.5 and b[2] > y1 + 0.5 for b in near):
            continue          # просвет уже перекрыт другой вертикалью
        by_band.setdefault((round(y1), round(start)), []).append((x, y1, start))
    out: list[VLine] = []
    for (_, _), bridges in by_band.items():
        xs = sorted({round(x, 1) for x, _, _ in bridges})
        if len(xs) < BAND_MIN_COLUMNS:
            continue
        top = min(b[1] for b in bridges)
        bottom = max(b[2] for b in bridges)
        lo, hi = xs[0], xs[-1]
        if (_covered(top, lo, hi, h_lines, 1.5) < 0.9 or _covered(bottom, lo, hi, h_lines, 1.5) < 0.9):
            continue
        inner = [x for x in xs if lo + tol < x < hi - tol]
        band_texts = [t for t in texts if t[1] < bottom and t[3] > top and t[0] < hi and t[2] > lo]
        keep = set(xs) - set(inner) | set(split_points(inner, band_texts, tol))
        if inner and not keep & set(inner):
            continue          # текст через границы граф — строка-заголовок на всю ширину
        seen = set()
        for x, y0, y1 in sorted(bridges):
            if round(x, 1) in keep and round(x, 1) not in seen:
                seen.add(round(x, 1))
                out.append((x, y0, y1))
    return out


def text_chunks(spans: Sequence[tuple[Box, float]], gap_em: float = 0.6) -> list[Box]:
    """Склеить соседние фрагменты строки (рамка, кегль) в куски: промежуток меньше 0,6 кегля — тот же
    кусок текста (пробел между словами), больше — разные графы."""
    out: list[list[float]] = []
    for box, size in sorted(spans, key=lambda s: s[0][0]):
        if out and box[0] - out[-1][2] <= gap_em * max(size, 1.0):
            cur = out[-1]
            cur[1], cur[2], cur[3] = min(cur[1], box[1]), max(cur[2], box[2]), max(cur[3], box[3])
        else:
            out.append(list(box))
    return [tuple(b) for b in out]


def split_points(inner: Sequence[float], texts: Sequence[Box], tol: float = 1.0) -> list[float]:
    """Какие границы граф провести через строку без вертикалей.

    Если текст не пересекает ни одной границы — все (обычная строка сетки: «Ду215», «40», «16» каталога
    стоят каждый в своей графе). Если заголовок идёт через границы граф, проводятся только
    непересечённые границы, по обе стороны от которых есть текст: «8.1.4. Перегрузочный узел…» по
    центру — одна ячейка на всю ширину; «Итого по разделу | 40 | 16» — заголовок и значения."""
    crossed = {x for x in inner for t in texts if t[0] < x - tol and t[2] > x + tol}
    if not crossed:
        return list(inner)
    out = []
    for x in inner:
        if any(t[0] < x - tol and t[2] > x + tol for t in texts):
            continue
        if any(t[2] <= x + tol for t in texts) and any(t[0] >= x - tol for t in texts):
            out.append(x)
    return out


# --------------------------------------------------------------------------- стыки линий САПР

JOINT_MIN = 2.0      # pt: pdf2docx отбрасывает фигуры, у которых обе стороны меньше этого
JOINT_TOL = 0.3      # pt


def bridge_joints(boxes: Sequence[Box], colors: Sequence, min_dim: float = JOINT_MIN,
                  tol: float = JOINT_TOL) -> list[Box]:
    """Квадратики на стыках линий (САПР рисует линию сетки кусками от стыка до стыка, а стык —
    отдельным квадратиком 1,5×1,5 pt) растягиваются вдоль линии до min_dim.

    pdf2docx выбрасывает фигуры меньше 2 pt, и линия рвётся на каждом стыке: таблица делится на
    куски, а текст ячейки, что пересекает разрыв («Подпись и дата» бокового штампа рядом с основным
    штампом), выпадает из таблицы горизонтальной строкой."""
    out = [tuple(b) for b in boxes]
    vert: dict[int, list[int]] = {}
    horz: dict[int, list[int]] = {}
    for i, b in enumerate(out):
        w, h = b[2] - b[0], b[3] - b[1]
        if h >= min_dim and w < min_dim:
            vert.setdefault(round(b[0] * 2), []).append(i)
        elif w >= min_dim and h < min_dim:
            horz.setdefault(round(b[1] * 2), []).append(i)

    def neighbours(index, key):
        for k in (key - 1, key, key + 1):
            yield from index.get(k, ())

    for i, b in enumerate(out):
        w, h = b[2] - b[0], b[3] - b[1]
        if max(w, h) >= min_dim:
            continue
        grow = None
        for j in neighbours(vert, round(b[0] * 2)):
            o = out[j]
            if colors[j] == colors[i] and abs(o[0] - b[0]) <= tol and abs(o[2] - b[2]) <= tol \
                    and (abs(o[3] - b[1]) <= tol or abs(o[1] - b[3]) <= tol):
                grow = "v"
                break
        if grow is None:
            for j in neighbours(horz, round(b[1] * 2)):
                o = out[j]
                if colors[j] == colors[i] and abs(o[1] - b[1]) <= tol and abs(o[3] - b[3]) <= tol \
                        and (abs(o[2] - b[0]) <= tol or abs(o[0] - b[2]) <= tol):
                    grow = "h"
                    break
        half = min_dim / 2 + 0.05
        if grow == "v":
            c = (b[1] + b[3]) / 2
            out[i] = (b[0], c - half, b[2], c + half)
        elif grow == "h":
            c = (b[0] + b[2]) / 2
            out[i] = (c - half, b[1], c + half, b[3])
    return out


def merge_collinear(boxes: Sequence[Box], colors: Sequence, thin: float = 3.0,
                    tol: float = JOINT_TOL) -> list[list[int]]:
    """Группы индексов тонких заливок одного цвета, что лежат на одной прямой встык (кусок —
    стык — кусок). Каждую группу можно заменить одной линией.

    Линии сетки из САПР нарисованы кусками от стыка до стыка. pdf2docx склеивает куски, только
    если вся связная сетка — одна сплошная заливка, иначе оставляет их порознь, и таблица рвётся
    на стыках: боковой штамп рядом с основным делился на обрывки, «Подпись и дата» выпадала из
    клетки горизонтальной строкой."""
    groups: list[list[int]] = []
    for axis in ("v", "h"):
        lo, hi, a0, a1 = (1, 3, 0, 2) if axis == "v" else (0, 2, 1, 3)
        items = []
        for i, b in enumerate(boxes):
            w, h = b[2] - b[0], b[3] - b[1]
            if (axis == "v" and w <= thin and h >= w) or (axis == "h" and h <= thin and w > h):
                items.append(i)
        items.sort(key=lambda i: (round(boxes[i][a0], 1), boxes[i][lo]))
        cur: list[int] = []
        for i in items:
            b = boxes[i]
            if cur:
                last = boxes[cur[-1]]
                end = max(boxes[k][hi] for k in cur)
                if colors[i] == colors[cur[-1]] and abs(b[a0] - last[a0]) <= tol \
                        and abs(b[a1] - last[a1]) <= tol and b[lo] <= end + tol:
                    cur.append(i)
                    continue
                if len(cur) > 1:
                    groups.append(cur)
            cur = [i]
        if len(cur) > 1:
            groups.append(cur)
    return groups

