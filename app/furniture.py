"""Поиск «оформления» листа: шапки, нижнего штампа, бокового штампа и рамки.

Эти элементы повторяются на каждом листе проектной документации. В Word их место — колонтитулы:
тогда штамп не «уезжает» на следующую страницу вслед за переполненным текстом, а правится в одном
месте для всех листов. Модуль чистый (без PyMuPDF), кроме адаптера elements_from_fitz.

Правила (координаты PDF, y вниз):

* рамка — прямоугольник из линий почти во всю страницу (ГОСТ: 20 мм слева, 5 мм с остальных сторон);
* боковой штамп — всё, что левее рамки; без рамки — вертикальный текст у левого края с его клетками;
* основная надпись по ГОСТ — горизонталь шириной 185 мм у правого края рамки над её низом;
* нижний штамп-рамка без ГОСТ-рамки — прямоугольник из линий у низа листа шириной от 40 % листа
  (плюс подписи под ним, например имя файла);
* шапка — длинная горизонтальная линия в верхней части листа, над которой немного элементов
  (логотип, шифр);
* если листов несколько — повторяющиеся на одних и тех же местах элементы у верха и низа листа
  (номера страниц совпадают с точностью до цифр).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import geometry as geo

Box = tuple[float, float, float, float]
MM = 72 / 25.4

STAMP_WIDTH = 185 * MM
MAX_BAND_SHARE = 0.36        # полоса выше этой доли листа — не оформление
LINE_THICKNESS = 2.5         # pt: путь не толще — линия
STAMP_WORDS = re.compile(r"(Изм\.|Кол\.\s?уч|Лист(?:ов)?\b|№\s?док|Подп(?:\.|ись)|Дата\b|Разраб|Пров(?:\.|ерил)"
                         r"|Н\.\s?контр|Утв(?:\.|ердил)|ГИП\b|Стадия|Нач\.\s?отд)", re.I)
MIN_BAND_GAP = 6.0           # pt: колонтитул отделён от текста хотя бы таким промежутком


@dataclass
class Elem:
    bbox: Box
    kind: str                       # text | path | graphic
    text: str = ""
    vertical: bool = False          # повёрнутый текст
    segs: list[Box] = field(default_factory=list)   # прямые отрезки пути (горизонтали/вертикали)


@dataclass
class PageElems:
    width: float
    height: float
    elems: list[Elem]


@dataclass
class Furniture:
    header: Box | None = None       # область шапки (верх листа)
    footer: Box | None = None       # область нижнего штампа
    sides: list[Box] = field(default_factory=list)   # боковые штампы
    frame: Box | None = None        # рамка листа
    drop: list[Box] = field(default_factory=list)    # пустые части полосы штампа (сетка без текста)
    body_footer: Box | None = None  # что убрать из тела вместо footer (верхняя линия штампа —
    #                                 одновременно нижняя граница таблицы спецификации)

    @property
    def empty(self) -> bool:
        return not (self.header or self.footer or self.sides)

    def exclusions(self) -> list[Box]:
        """Области, которые убираются из основного текста страницы."""
        footer = self.body_footer or self.footer
        return [b for b in (self.header, footer) if b] + list(self.sides) + list(self.drop)


# --------------------------------------------------------------------------- вспомогательное

def merge_segments(segs, horizontal: bool, gap: float = 3.0, tol: float = 1.0) -> list[Box]:
    """Склеить куски одной линии (Word рисует линию таблицы отдельно для каждой ячейки)."""
    items = []
    for s in segs:
        if horizontal:
            items.append(((s[1] + s[3]) / 2, s[0], s[2]))
        else:
            items.append(((s[0] + s[2]) / 2, s[1], s[3]))
    items.sort()
    rows: list[list] = []
    for pos, a, b in items:
        if rows and abs(rows[-1][0] - pos) <= tol:
            rows[-1][1].append((a, b))
        else:
            rows.append([pos, [(a, b)]])
    out = []
    for pos, spans in rows:
        spans.sort()
        cur_a, cur_b = spans[0]
        for a, b in spans[1:]:
            if a <= cur_b + gap:
                cur_b = max(cur_b, b)
            else:
                out.append((cur_a, pos, cur_b, pos) if horizontal else (pos, cur_a, pos, cur_b))
                cur_a, cur_b = a, b
        out.append((cur_a, pos, cur_b, pos) if horizontal else (pos, cur_a, pos, cur_b))
    return out


def lines(page: PageElems, horizontal: bool, skip=None) -> list[Box]:
    """Склеенные горизонтали или вертикали страницы (кроме элементов, для которых skip(e))."""
    segs = []
    for e in page.elems:
        if skip is not None and skip(e):
            continue
        for s in e.segs:
            # толстые линии (ГОСТ: 0,5–1,4 мм) Word и САПР рисуют тонкими заливками
            if horizontal and s[3] - s[1] <= LINE_THICKNESS and s[2] - s[0] > s[3] - s[1]:
                segs.append(s)
            elif not horizontal and s[2] - s[0] <= LINE_THICKNESS and s[3] - s[1] > s[2] - s[0]:
                segs.append(s)
    return merge_segments(segs, horizontal)


def _inside_any(b: Box, boxes, tol: float = 1.0) -> bool:
    return any(geo.contains(o, b, tol) for o in boxes)


def _on_frame_edge(s: Box, frame: Box, tol: float = 2.0) -> bool:
    x0, y0, x1, y1 = s
    if y1 - y0 <= 1.0:      # горизонталь на верхней или нижней стороне
        return (abs(y0 - frame[1]) < tol or abs(y0 - frame[3]) < tol) and x0 >= frame[0] - tol and x1 <= frame[2] + tol
    if x1 - x0 <= 1.0:      # вертикаль на левой или правой стороне
        return (abs(x0 - frame[0]) < tol or abs(x0 - frame[2]) < tol) and y0 >= frame[1] - tol and y1 <= frame[3] + tol
    return False


def _is_frame_elem(e: Elem, frame: Box | None) -> bool:
    """Рамка целиком или её кусок (Word рисует стороны рамки по кускам)."""
    if frame is None or e.kind != "path" or not e.segs:
        return False
    x0, y0, x1, y1 = e.bbox
    if geo.contains((frame[0] - 2, frame[1] - 2, frame[2] + 2, frame[3] + 2), e.bbox) and \
            (x1 - x0 >= 0.9 * (frame[2] - frame[0]) and y1 - y0 >= 0.9 * (frame[3] - frame[1])):
        return True
    return all(_on_frame_edge(s, frame) for s in e.segs)


# --------------------------------------------------------------------------- рамка

FRAME_EDGE = 0.1     # рамка ГОСТ идёт по полям листа: не дальше 10 % высоты от верха и низа


def find_frame(page: PageElems) -> Box | None:
    W, H = page.width, page.height
    best = None
    for e in page.elems:
        if e.kind != "path":
            continue
        x0, y0, x1, y1 = e.bbox
        if x1 - x0 >= 0.7 * W and y1 - y0 >= 0.7 * H and len(e.segs) >= 4:
            if best is None or geo.area(e.bbox) > geo.area(best):
                best = e.bbox
    if best is None:
        hs = [s for s in lines(page, True) if s[2] - s[0] >= 0.7 * W]
        vs = [s for s in lines(page, False) if s[3] - s[1] >= 0.7 * H]
        if len(hs) >= 2 and len(vs) >= 2:
            top, bottom = min(hs, key=lambda s: s[1]), max(hs, key=lambda s: s[1])
            left, right = min(vs, key=lambda s: s[0]), max(vs, key=lambda s: s[0])
            box = (left[0], top[1], right[2], bottom[3])
            ok = all(abs(a - b) <= 3 for a, b in ((top[0], left[0]), (top[2], right[2]),
                                                    (left[1], top[1]), (left[3], bottom[3])))
            if ok and box[2] - box[0] >= 0.7 * W and box[3] - box[1] >= 0.7 * H:
                best = box
    if best is None:
        return None
    x0, y0, x1, y1 = best
    if x0 < 2 or y0 < 2 or W - x1 < 2 or H - y1 < 2:
        return None   # рамка по краю листа — это не ГОСТ-рамка, а обводка страницы
    if y0 > FRAME_EDGE * H or H - y1 > FRAME_EDGE * H:
        return None   # начинается под заголовком — это контур большой таблицы, а не рамка листа
    return tuple(round(v, 2) for v in best)


# --------------------------------------------------------------------------- боковой штамп

def find_sides(page: PageElems, frame: Box | None) -> list[Box]:
    W = page.width
    if frame is not None:
        left = [e for e in page.elems if e.bbox[2] <= frame[0] + 1.5 and not _is_frame_elem(e, frame)]
        right = [e for e in page.elems if e.bbox[0] >= frame[2] - 1.5 and not _is_frame_elem(e, frame)]
        boxes = []
        for group in (left, right):
            if any(e.kind == "text" for e in group):
                boxes += geo.cluster_boxes((e.bbox for e in group), 3, 25)
        return [b for b in boxes if any(e.kind == "text" and geo.contains(b, e.bbox, 1) for e in page.elems)]
    # без рамки боковой штамп стоит в поле подшивки (до ~20 мм от края листа)
    strip = min(22 * MM, 0.14 * W)
    vtexts = [e for e in page.elems if e.kind == "text" and e.vertical and e.bbox[2] <= strip]
    if not vtexts:
        return []
    core = vtexts[0].bbox
    for e in vtexts[1:]:
        core = geo.union(core, e.bbox)
    # клетки штампа: пути в той же полосе, пересекающиеся с текстом по вертикали
    limit = core[2] + 3 * MM
    members = [e.bbox for e in page.elems
               if e.bbox[2] <= limit and e.bbox[1] < core[3] + 40 and e.bbox[3] > core[1] - 40]
    boxes = geo.cluster_boxes(members, 3, 25)
    # линии, уходящие из «штампа» вправо, — это первый столбец таблицы (например, «№ п/п»)
    hs = lines(page, True)
    return [b for b in boxes
            if not any(b[1] - 1 <= h[1] <= b[3] + 1 and h[0] <= b[2] + 1 and h[2] > b[2] + 10 * MM for h in hs)]


# --------------------------------------------------------------------------- шапка

def find_header_by_rule(page: PageElems, frame: Box | None, sides: list[Box]) -> Box | None:
    W, H = page.width, page.height
    rx0, rx1 = (frame[0], frame[2]) if frame else (0.0, W)
    top_ref = frame[1] if frame else 0.0
    candidates = []
    skip = lambda e: _is_frame_elem(e, frame) or _inside_any(e.bbox, sides)   # noqa: E731
    for s in lines(page, True, skip):
        y = (s[1] + s[3]) / 2
        if s[2] - s[0] < 0.5 * (rx1 - rx0) or not (top_ref + 4 < y <= top_ref + 0.15 * H):
            continue
        candidates.append((y, s))
    for y, s in sorted(candidates, reverse=True):     # нижняя подходящая линия — граница шапки
        above, straddle = [], False
        for e in page.elems:
            if _is_frame_elem(e, frame) or _inside_any(e.bbox, sides):
                continue
            verticals = [v for v in e.segs if v[2] - v[0] <= 1 and s[0] - 1 <= v[0] <= s[2] + 1]
            # вертикаль, начинающаяся на линии и уходящая вниз, — линия оказалась верхом таблицы
            if any(abs(v[1] - y) <= 2 and v[3] > y + 5 for v in verticals):
                straddle = True
                break
            if e.bbox[3] <= y + 1.0:
                above.append(e)
            elif e.bbox[1] < y - 1.0:
                # элемент пересекает линию: текст или вертикаль таблицы — это не шапка
                if e.kind == "text" or any(v[1] < y - 2 and v[3] > y + 2 for v in verticals):
                    straddle = True
                    break
        texts = [e for e in above if e.kind == "text"]
        if straddle or not above or len(texts) > 8:
            continue
        # шапка листа — логотип или «слева название/справа шифр», а не заголовок с чертой под ним
        has_logo = any(e.kind == "graphic" for e in above)
        left_right = any(a.bbox[2] < (rx0 + rx1) / 2 < b.bbox[0] and abs(a.bbox[1] - b.bbox[1]) < 15
                         for a in texts for b in texts)
        if not (has_logo or left_right):
            continue
        return (rx0, 0.0 if not frame else top_ref - 1, rx1, y + 0.75)
    return None


# --------------------------------------------------------------------------- нижний штамп

def find_stamp(page: PageElems, frame: Box) -> Box | None:
    """Основная надпись по строению: вертикали штампа доходят до низа рамки и длиной 8–70 мм
    (столбцы таблицы листа кончаются на верхней линии штампа или идут от верха рамки), а внутри
    есть хотя бы два слова граф штампа («Изм.», «Лист», «Подп.», «Дата», «Разраб.»…).
    Ширина не важна: на A4 штамп бывает во всю рамку, у Word-шаблонов — не ровно 185 мм."""
    fx0, fy0, fx1, fy1 = frame
    skip = lambda e: _is_frame_elem(e, frame)   # noqa: E731
    vs = [v for v in lines(page, False, skip)
          if abs(v[3] - fy1) <= 3 and fx0 + 3 < v[0] < fx1 - 3 and 8 * MM <= v[3] - v[1] <= 70 * MM]
    if not vs:
        return None
    top = min(v[1] for v in vs)
    hs = [h for h in lines(page, True, skip) if abs(h[1] - top) <= 2 and h[2] >= fx1 - 3]
    x0 = min((h[0] for h in hs), default=min(v[0] for v in vs))
    box = (max(x0, fx0), top, fx1, fy1)
    words = set()
    for e in page.elems:
        if e.kind == "text" and geo.contains((box[0] - 1, box[1] - 1, box[2] + 1, box[3] + 1), e.bbox, 1):
            words.update(m.group(1).lower() for m in STAMP_WORDS.finditer(e.text))
    return box if len(words) >= 2 else None


def find_gost_stamp(page: PageElems, frame: Box) -> Box | None:
    best = None
    for s in lines(page, True):
        y = (s[1] + s[3]) / 2
        if abs(s[2] - frame[2]) > 3 or abs((s[2] - s[0]) - STAMP_WIDTH) > 6 * MM:
            continue
        if not (frame[3] - 65 * MM <= y <= frame[3] - 8 * MM):
            continue
        if best is None or y < best[1]:
            best = (s[0], y, frame[2], frame[3])
    return best


STAMP_MAX_TEXTS = 6      # надписей в нижней рамке-колонтитуле «2026 | название | 3»; у таблицы больше


def _stamp_like(page: PageElems, box: Box) -> bool:
    """Рамка у низа листа — штамп, а не таблица в конце страницы: графы штампа («Изм.», «Лист»,
    «№ докум.», «Подп.»), или пусто, или несколько надписей с номером листа у края рамки."""
    texts = [e for e in page.elems if e.kind == "text" and geo.contains(box, e.bbox, 1.5)]
    if not texts or any(STAMP_WORDS.search(e.text) for e in texts):
        return True
    if len(texts) > STAMP_MAX_TEXTS:
        return False
    edge = 0.12 * (box[2] - box[0])
    return any(e.text.strip().isdigit() and len(e.text.strip()) <= 3
               and (e.bbox[0] - box[0] < edge or box[2] - e.bbox[2] < edge) for e in texts)


def find_footer_box(page: PageElems, frame: Box | None, sides: list[Box]) -> Box | None:
    W, H = page.width, page.height
    rx0, rx1 = (frame[0], frame[2]) if frame else (0.0, W)
    bottom_ref = frame[3] if frame else H
    skip = (lambda e: _is_frame_elem(e, frame)) if frame else None
    hs = [s for s in lines(page, True, skip) if s[2] - s[0] >= 0.4 * (rx1 - rx0) and s[1] >= 0.72 * H]
    vs = lines(page, False, skip)
    boxes = []
    for t in hs:
        for b in hs:
            dy = b[1] - t[1]
            if not (6 <= dy <= 0.15 * H) or abs(b[0] - t[0]) > 3 or abs(b[2] - t[2]) > 3:
                continue
            spanning = [v for v in vs if t[0] - 3 <= v[0] <= t[2] + 3 and v[1] <= t[1] + 3 and v[3] >= b[1] - 3]
            has_l = any(abs(v[0] - t[0]) <= 3 for v in spanning)
            has_r = any(abs(v[0] - t[2]) <= 3 for v in spanning)
            inner = any(t[1] + 2 < h[1] < b[1] - 2 and h[2] - h[0] >= 0.5 * (t[2] - t[0]) for h in hs)
            # рамка целиком или «табличка» в одну строку с перегородками (штамп Word без боковых сторон)
            if (has_l and has_r) or (spanning and not inner):
                boxes.append((t[0], t[1], t[2], b[3]))
    if not boxes:
        return None
    box = max(boxes, key=lambda bx: (bx[3], -bx[1]))   # самый нижний, из вложенных — самый высокий
    y = box[1]
    # вертикаль начинается выше верхней линии — таблица продолжается вверх, это не штамп
    if any(box[0] - 1 <= v[0] <= box[2] + 1 and v[1] < y - 5 and v[3] > y - 2 for v in vs):
        return None
    # те же колонки, что у таблицы выше (между ними строка без линий, «Ду830»), — тоже таблица
    own = {round(v[0], 1) for v in vs if box[0] - 1 <= v[0] <= box[2] + 1 and v[1] >= y - 3 and v[3] <= box[3] + 3}
    above = [v[0] for v in vs if y - 0.25 * H <= v[3] <= y + 2 and v[1] < y - 5]
    same = sum(1 for x in own if any(abs(x - a) <= 1.0 for a in above))
    if len(own) >= 3 and same >= max(3, 0.6 * len(own)):
        return None
    for e in page.elems:
        if _is_frame_elem(e, frame) or _inside_any(e.bbox, sides):
            continue
        if e.bbox[1] < y - 1.0 < e.bbox[3] - 1.0:
            if e.kind == "text" or any(v[1] < y - 2 and v[3] > y + 2 for v in e.segs if v[2] - v[0] <= 1):
                return None     # что-то пересекает верх рамки — это таблица текста, не штамп
    below = [e for e in page.elems if e.kind == "text" and e.bbox[1] >= box[3] - 1]
    if len(below) > 3 or box[3] < bottom_ref - 0.2 * H:
        return None
    return (rx0, y - 0.75, rx1, H if not frame else frame[3] + 1)


# --------------------------------------------------------------------------- повторы

_DIGITS = re.compile(r"\d+")


def _norm(text: str) -> str:
    return _DIGITS.sub("#", re.sub(r"\s+", " ", text.strip()))


def _same_place(a: Elem, b: Elem, pa: PageElems, pb: PageElems, top: bool, tol: float = 3.0) -> bool:
    if a.kind != b.kind or (a.kind == "text" and _norm(a.text) != _norm(b.text)):
        return False
    if top:
        dy = abs(a.bbox[1] - b.bbox[1])
    else:
        dy = abs((pa.height - a.bbox[3]) - (pb.height - b.bbox[3]))
    if dy > tol:
        return False
    if a.kind == "text" and _norm(a.text):
        # та же надпись на той же высоте у края листа — колонтитул, где бы по горизонтали он ни стоял
        # (чётные страницы — слева, нечётные — справа; поля при этом бывают разные)
        return True
    return (abs(a.bbox[0] - b.bbox[0]) <= tol or abs((pa.width - a.bbox[2]) - (pb.width - b.bbox[2])) <= tol
            or abs((a.bbox[0] + a.bbox[2]) / pa.width - (b.bbox[0] + b.bbox[2]) / pb.width) * pa.width <= 2 * tol
            # зеркальные колонтитулы чётных и нечётных страниц: слева на одной — справа на другой
            or abs(a.bbox[0] - (pb.width - b.bbox[2])) <= tol and abs((pa.width - a.bbox[2]) - b.bbox[0]) <= max(
                tol, 0.5 * (a.bbox[2] - a.bbox[0])))


def _elem_key(e: Elem) -> tuple:
    return tuple(e.bbox), e.kind, e.text, e.vertical, tuple(tuple(s) for s in e.segs)


def repeated_bands(pages: list[PageElems], zone: float = 0.13) -> list[tuple[float | None, float | None]]:
    """Для каждого листа: (низ повторяющейся шапки, верх повторяющегося низа) или None."""
    n = len(pages)
    out: list[tuple[float | None, float | None]] = [(None, None)] * n
    if n < 2:
        return out

    def in_zone(e: Elem, pg: PageElems, top: bool) -> bool:
        return e.bbox[3] <= zone * pg.height if top else e.bbox[1] >= (1 - zone) * pg.height

    # элементы у верхнего и нижнего края каждого листа — один раз на документ
    zones = {top: [[e for e in pg.elems if in_zone(e, pg, top)] for pg in pages] for top in (True, False)}
    for i, p in enumerate(pages):
        # сравниваем с соседними листами (до 6 с каждой стороны): быстро и на длинных документах
        window = [j for j in range(max(0, i - 6), min(n, i + 7)) if j != i]
        need = max(2, (len(window) + 2) // 2)
        res = []
        for top in (True, False):
            reps = []
            zone_elems = zones[top][i]
            if len(zone_elems) > 400:      # чертёж у края листа — повторы не ищем
                res.append(None)
                continue
            for e in zone_elems:
                hits = 1 + sum(1 for j in window if any(
                    _same_place(e, f, p, pages[j], top) for f in zones[top][j]))
                if hits >= need:
                    reps.append(e)
            # повтор, вплотную примыкающий к следующему тексту, — это начало тела (например,
            # одинаковая строка заголовков таблицы на каждом листе), а не шапка
            while reps:
                if top:
                    last = max(reps, key=lambda e: e.bbox[3])
                    nxt = [e.bbox[1] - last.bbox[3] for e in p.elems if e is not last and e.bbox[1] >= last.bbox[3] - 1]
                else:
                    last = min(reps, key=lambda e: e.bbox[1])
                    nxt = [last.bbox[1] - e.bbox[3] for e in p.elems if e is not last and e.bbox[3] <= last.bbox[1] + 1]
                gap = min(nxt) if nxt else 1e9
                if gap >= MIN_BAND_GAP or last.kind == "path":
                    break
                reps.remove(last)
            if not any(e.kind == "text" for e in reps):
                res.append(None)
                continue
            rep_keys = {_elem_key(e) for e in reps}    # как «e in reps»: равенство по полям
            if top:
                edge = max(e.bbox[3] for e in reps) + 1.0
                bad = any(e.bbox[1] < edge - 1.0 and e.bbox[3] > 0 and _elem_key(e) not in rep_keys
                          for e in p.elems if e.bbox[3] > edge + 0.5)
            else:
                edge = min(e.bbox[1] for e in reps) - 1.0
                bad = any(e.bbox[3] > edge + 1.0 and _elem_key(e) not in rep_keys
                          for e in p.elems if e.bbox[1] < edge - 0.5)
            res.append(None if bad else edge)
        out[i] = (res[0], res[1])
    return out


# --------------------------------------------------------------------------- всё вместе

def detect(pages: list[PageElems]) -> list[Furniture]:
    reps = repeated_bands(pages)
    result = []
    for p, (rep_top, rep_bottom) in zip(pages, reps):
        W, H = p.width, p.height
        f = Furniture()
        f.frame = find_frame(p)
        f.sides = find_sides(p, f.frame)
        stamp = (find_stamp(p, f.frame) or find_gost_stamp(p, f.frame)) if f.frame is not None else None
        if stamp is not None:
            # вместе с подписями под рамкой (имя файла, «Формат А3»)
            f.footer = (stamp[0] - 0.75, stamp[1] - 0.75, W, H)
            f.body_footer = (f.footer[0], stamp[1] + 0.75, f.footer[2], f.footer[3])
            # полоса левее штампа: пустая сетка — убрать, с текстом — тоже в нижний колонтитул
            # с левой стороной рамки: иначе она тянется в тексте страницы до низа листа и нижнее
            # поле получается узким
            strip = (f.frame[0] - 1.0, stamp[1] + 0.75, stamp[0] - 0.5, f.frame[3] + 1.0)
            if strip[2] > strip[0] + 5:
                if any(e.kind == "text" and geo.contains(strip, e.bbox, 1) for e in p.elems):
                    f.footer = (f.frame[0] - 1.0, f.footer[1], f.footer[2], f.footer[3])
                    f.body_footer = (f.footer[0], stamp[1] + 0.75, f.footer[2], f.footer[3])
                else:
                    f.drop.append(strip)
        else:
            f.footer = find_footer_box(p, f.frame, f.sides)
            if f.footer is not None and rep_bottom is not None and f.footer[1] < rep_bottom - 20:
                f.footer = None     # ниже есть повторяющийся низ листа — «рамка» была строками таблицы
            if f.footer is not None and rep_bottom is None and not _stamp_like(p, f.footer):
                f.footer = None     # таблица в конце страницы, а не штамп: не повторяется и без граф штампа
        # лист-«бланк»: рамка, штамп, боковой штамп или рамочный нижний блок
        blank = f.frame is not None or bool(f.sides) or f.footer is not None
        if f.footer is None and rep_bottom is not None:
            f.footer = (0.0, rep_bottom, W, H)
        if f.footer is None and f.frame is not None:
            # под рамкой без штампа (титульный лист): имя файла, «Формат А4» — тоже оформление листа
            below = [e for e in p.elems if e.kind == "text" and e.bbox[1] >= f.frame[3] - 0.5]
            if below:
                f.footer = (0.0, f.frame[3] + 0.5, W, H)
        # шапка: повтор на соседних листах надёжнее; по линии — только на листе-«бланке»
        # (с рамкой, штампом или боковым штампом), иначе это может быть заголовок документа
        # или строка заголовков таблицы (каталог: повторяющийся низ листа — ещё не бланк)
        if blank:
            f.header = find_header_by_rule(p, f.frame, f.sides)
        if f.header is None and rep_top is not None:
            f.header = (0.0, 0.0, W, rep_top)
        # санитарные проверки: полосы не больше трети листа, тело страницы не пустое
        if f.header and f.header[3] - f.header[1] > MAX_BAND_SHARE * H:
            f.header = None
        if f.footer and f.footer[3] - f.footer[1] > MAX_BAND_SHARE * H:
            f.footer, f.drop, f.body_footer = None, [], None
        excl = f.exclusions()
        body = [e for e in p.elems if e.kind == "text" and not _inside_any(e.bbox, excl, 1.5)]
        if not body:
            f = Furniture(frame=f.frame)
        result.append(f)
    return result


# --------------------------------------------------------------------------- PyMuPDF → элементы

MAX_DRAWINGS = 20000   # чертёж с огромным числом путей: оформление не ищем (долго и ненадёжно)


def elements_from_fitz(page) -> PageElems:
    """Текстовые строки, пути и картинки страницы (страница без поворота)."""
    W, H = page.rect.width, page.rect.height
    elems: list[Elem] = []
    drawings = page.get_cdrawings()
    if len(drawings) > MAX_DRAWINGS:
        return PageElems(W, H, [])
    for b in page.get_text("dict").get("blocks", []):
        if b.get("type") == 0:
            for ln in b.get("lines", []):
                text = "".join(s.get("text", "") for s in ln.get("spans", [])).strip()
                if text:
                    d = ln.get("dir", (1, 0))
                    elems.append(Elem(tuple(ln["bbox"]), "text", text, vertical=abs(d[0]) < 0.5))
        else:
            elems.append(Elem(tuple(b["bbox"]), "graphic"))
    for dr in drawings:
        items = dr.get("items", [])
        bbox = tuple(dr.get("rect", (0, 0, 0, 0)))
        if not items:
            continue
        if dr.get("fill") is not None and dr.get("color") is None:
            # заливка: тонкая — линия (так рисует линии Word: залитый прямоугольник 0,5–1,5 pt),
            # остальное — графика (фон ячейки и т.п.)
            x0, y0, x1, y1 = bbox
            if min(x1 - x0, y1 - y0) <= 3.0:
                if x1 - x0 >= y1 - y0:
                    yc = (y0 + y1) / 2
                    seg = (x0, yc, x1, yc)
                else:
                    xc = (x0 + x1) / 2
                    seg = (xc, y0, xc, y1)
                elems.append(Elem(bbox, "path", segs=[seg]))
            elif not all(it[0] == "re" for it in items):
                elems.append(Elem(bbox, "graphic"))
            continue
        if not geo.items_axis_aligned(items):
            elems.append(Elem(bbox, "graphic"))
            continue
        elems.append(Elem(bbox, "path", segs=list(path_segments(items))))
    return PageElems(W, H, elems)


def path_segments(items):
    """Горизонтальные/вертикальные отрезки пути как тонкие прямоугольники (x0, y0, x1, y1)."""
    def seg(p0, p1):
        axis = geo.segment_axis(p0, p1)
        if axis in ("h", "v"):
            (a, b) = geo.snap_segment(p0, p1, axis)
            return (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
        return None
    for it in items:
        k = it[0]
        if k == "l" or k == "c":
            s = seg(tuple(it[1]), tuple(it[-1]))
            if s:
                yield s
        elif k == "re":
            x0, y0, x1, y1 = it[1]
            for s in ((x0, y0, x1, y0), (x1, y0, x1, y1), (x0, y1, x1, y1), (x0, y0, x0, y1)):
                yield (min(s[0], s[2]), min(s[1], s[3]), max(s[0], s[2]), max(s[1], s[3]))
        elif k == "qu":
            q = [tuple(p) for p in it[1]]
            pts = [q[0], q[1], q[3], q[2]]
            for a, b in zip(pts, pts[1:] + pts[:1]):
                s = seg(a, b)
                if s:
                    yield s
