"""Разбор только части страницы в pdf2docx: фильтры текста, картинок и линий по областям.

Оформление листа (шапка, штамп, боковой штамп) переносится в колонтитулы так: страница
разбирается несколько раз — тело без областей оформления, затем каждая область отдельно. PDF при
этом не меняется: фильтры стоят на входе pdf2docx (сырой текст, картинки, линии) и включаются
для конкретной страницы на время прохода (контекст regions()).
"""
from __future__ import annotations

import logging
from contextlib import contextmanager
from dataclasses import dataclass, field

log = logging.getLogger(__name__)

Box = tuple[float, float, float, float]


@dataclass
class Spec:
    keep: list[Box] | None = None                 # оставить только эти области (проход колонтитула)
    exclude: list[Box] = field(default_factory=list)   # убрать эти области (тело страницы)
    frame: Box | None = None                      # рамка листа: её линии не должны стать таблицей
    frame_mode: str = "touched"                   # touched — сторона рамки остаётся, если к ней примыкает
    #                                               таблица (тело, штамп); drop — убрать (шапка, боковой штамп)
    top_limit: float | None = None                # тело: верхнее поле не выше низа шапки
    bottom_limit: float | None = None             # тело: нижнее поле не ниже верха штампа
    lr_margin: tuple[float, float] | None = None  # одинаковые левое/правое поля для тела и колонтитулов


MIN_CLIPPED = 3.0   # pt

_REG: dict[tuple[object, int], Spec] = {}
_TOKEN = "_p2w_regions_token"


def _doc_token(doc):
    """Метка документа: page.parent в PyMuPDF может быть weakref-прокси, поэтому не id()."""
    token = getattr(doc, _TOKEN, None)
    if token is None:
        token = f"doc{id(doc)}"
        try:
            setattr(doc, _TOKEN, token)
        except Exception:
            pass
    return token


@contextmanager
def regions(doc, specs: dict[int, Spec]):
    keys = []
    if specs:
        token = _doc_token(doc)
        for pno, spec in specs.items():
            key = (token, int(pno))
            _REG[key] = spec
            keys.append(key)
    try:
        yield
    finally:
        for key in keys:
            _REG.pop(key, None)


def spec_for(page_engine) -> Spec | None:
    if not _REG or page_engine is None:
        return None
    try:
        return _REG.get((_doc_token(page_engine.parent), page_engine.number))
    except Exception:
        return None


def allowed(spec: Spec, x: float, y: float) -> bool:
    if spec.keep is not None and not any(b[0] <= x <= b[2] and b[1] <= y <= b[3] for b in spec.keep):
        return False
    return not any(b[0] <= x <= b[2] and b[1] <= y <= b[3] for b in spec.exclude)


def _center(b) -> tuple[float, float]:
    return (b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0


# --------------------------------------------------------------------------- текст

def filter_text(page_engine, blocks: list) -> list:
    """Сырые блоки текста (rawdict pdf2docx): убрать символы вне разрешённых областей."""
    spec = spec_for(page_engine)
    if spec is None:
        return blocks
    out = []
    for block in blocks:
        if "lines" not in block:
            if allowed(spec, *_center(block.get("bbox", (0, 0, 0, 0)))):
                out.append(block)
            continue
        lines = []
        for line in block["lines"]:
            spans = []
            for span in line.get("spans", []):
                chars = [c for c in span.get("chars", ()) if allowed(spec, *_center(c["bbox"]))]
                if not chars:
                    continue
                if len(chars) != len(span.get("chars", ())):
                    span = dict(span)
                    span["chars"] = chars
                    span["bbox"] = _bbox_of(c["bbox"] for c in chars)
                spans.append(span)
            if not spans:
                continue
            line = dict(line)
            line["spans"] = spans
            line["bbox"] = _bbox_of(s["bbox"] for s in spans)
            lines.append(line)
        if lines:
            block = dict(block)
            block["lines"] = lines
            block["bbox"] = _bbox_of(ln["bbox"] for ln in lines)
            out.append(block)
    return out


def _bbox_of(boxes) -> tuple:
    boxes = list(boxes)
    return (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))


# --------------------------------------------------------------------------- картинки

def filter_images(page_engine, images: list) -> list:
    spec = spec_for(page_engine)
    if spec is None:
        return images
    return [im for im in images if allowed(spec, *_center(im.get("bbox", (0, 0, 0, 0))))]


# --------------------------------------------------------------------------- линии и заливки

def _intervals_keep(a: float, b: float, cuts: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """Части отрезка [a, b], попадающие в объединение интервалов cuts."""
    res = []
    for c0, c1 in sorted(cuts):
        lo, hi = max(a, c0), min(b, c1)
        if hi > lo:
            if res and lo <= res[-1][1]:
                res[-1] = (res[-1][0], max(res[-1][1], hi))
            else:
                res.append((lo, hi))
    return res


def _intervals_drop(a: float, b: float, cuts: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """Части отрезка [a, b] вне объединения интервалов cuts."""
    res, cur = [], a
    for c0, c1 in sorted(cuts):
        if c1 <= cur or c0 >= b:
            continue
        if c0 > cur:
            res.append((cur, c0))
        cur = max(cur, c1)
    if cur < b:
        res.append((cur, b))
    return res


def clip_stroke(spec: Spec, start, end) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """Обрезать горизонтальный/вертикальный отрезок по областям спецификации."""
    (x0, y0), (x1, y1) = start, end
    horizontal = abs(y1 - y0) <= abs(x1 - x0)
    if horizontal:
        y, a, b = (y0 + y1) / 2, min(x0, x1), max(x0, x1)
        rel = lambda bx: bx[1] <= y <= bx[3]               # noqa: E731
        rng = lambda bx: (bx[0], bx[2])                    # noqa: E731
    else:
        x, a, b = (x0 + x1) / 2, min(y0, y1), max(y0, y1)
        rel = lambda bx: bx[0] <= x <= bx[2]               # noqa: E731
        rng = lambda bx: (bx[1], bx[3])                    # noqa: E731
    parts = [(a, b)]
    if spec.keep is not None:
        cuts = [rng(bx) for bx in spec.keep if rel(bx)]
        parts = [p for lo, hi in parts for p in _intervals_keep(lo, hi, cuts)]
    cuts = [rng(bx) for bx in spec.exclude if rel(bx)]
    if cuts:
        parts = [p for lo, hi in parts for p in _intervals_drop(lo, hi, cuts)]
    out = []
    for lo, hi in parts:
        # короткий обрезок длинной линии (вертикаль таблицы, упёршаяся в край области) — не граница:
        # иначе pdf2docx построит по нему лишние столбцы
        if hi - lo < 0.5 or (hi - lo < MIN_CLIPPED and hi - lo < (b - a) - 0.5):
            continue
        out.append(((lo, y), (hi, y)) if horizontal else ((x, lo), (x, hi)))
    return out


def _frame_side(frame: Box, s, e) -> bool:
    """Отрезок лежит на стороне рамки (в том числе кусок стороны — Word рисует рамку кусками)."""
    (x0, y0), (x1, y1) = s, e
    if abs(y0 - y1) <= 0.5:
        return (abs(y0 - frame[1]) <= 1.5 or abs(y0 - frame[3]) <= 1.5) and \
            min(x0, x1) >= frame[0] - 2 and max(x0, x1) <= frame[2] + 2
    if abs(x0 - x1) <= 0.5:
        return (abs(x0 - frame[0]) <= 1.5 or abs(x0 - frame[2]) <= 1.5) and \
            min(y0, y1) >= frame[1] - 2 and max(y0, y1) <= frame[3] + 2
    return False


def _touched(s, e, others) -> bool:
    """К линии рамки примыкает конец другой линии (рамка служит границей таблицы)."""
    (x0, y0), (x1, y1) = s, e
    horizontal = abs(y0 - y1) <= 0.5
    for (a, b) in others:
        for p in (a, b):
            if horizontal and abs(p[1] - y0) <= 2.0 and min(x0, x1) - 1 <= p[0] <= max(x0, x1) + 1 \
                    and abs(a[0] - b[0]) <= 0.5:
                return True
            if not horizontal and abs(p[0] - x0) <= 2.0 and min(y0, y1) - 1 <= p[1] <= max(y0, y1) + 1 \
                    and abs(a[1] - b[1]) <= 0.5:
                return True
    return False


THIN = 3.0   # pt: заливка тоньше — это линия (так Word рисует границы таблиц и рамку)


def _as_line(sh):
    """Штрих или тонкая заливка → (начало, конец, толщина) по осевой линии; иначе None."""
    if "start" in sh and "end" in sh:
        return tuple(sh["start"]), tuple(sh["end"]), None
    if "bbox" in sh:
        x0, y0, x1, y1 = sh["bbox"]
        if min(x1 - x0, y1 - y0) <= THIN:
            if x1 - x0 >= y1 - y0:
                yc = (y0 + y1) / 2
                return (x0, yc), (x1, yc), y1 - y0
            xc = (x0 + x1) / 2
            return (xc, y0), (xc, y1), x1 - x0
    return None


def _rebuild(sh, s, e, thick):
    if thick is None:
        return {**sh, "start": s, "end": e}
    (x0, y0), (x1, y1) = s, e
    h = thick / 2
    if abs(y0 - y1) <= 0.5:
        return {**sh, "bbox": [min(x0, x1), y0 - h, max(x0, x1), y0 + h]}
    return {**sh, "bbox": [x0 - h, min(y0, y1), x0 + h, max(y0, y1)]}


def filter_shapes(page_engine, shapes: list, images: list) -> tuple[list, list]:
    spec = spec_for(page_engine)
    if spec is None:
        return shapes, images
    lines, frame_lines, out = [], [], []
    for sh in shapes:
        geom = _as_line(sh)
        if geom is not None:
            start, end, thick = geom
            for s, e in clip_stroke(spec, start, end):
                item = (s, e, _rebuild(sh, s, e, thick))
                if spec.frame is not None and _frame_side(spec.frame, s, e):
                    frame_lines.append(item)
                else:
                    lines.append(item)
        elif "bbox" in sh:
            if allowed(spec, *_center(sh["bbox"])):
                out.append(sh)
        else:
            out.append(sh)
    if spec.frame_mode == "touched":
        # сторона рамки остаётся, только если служит границей таблицы (спецификация, штамп)
        ends = [(s, e) for s, e, _ in lines]
        lines += [item for item in frame_lines if _touched(item[0], item[1], ends)]
    out = [sh for _, _, sh in lines] + out
    images = [im for im in images if allowed(spec, *_center(im.get("bbox", (0, 0, 0, 0))))]
    return out, images


# --------------------------------------------------------------------------- поля страницы

def adjust_margin(raw_page, margin: tuple) -> tuple:
    spec = spec_for(getattr(raw_page, "page_engine", None))
    if spec is None:
        return margin
    left, right, top, bottom = margin
    if spec.lr_margin is not None:
        left, right = spec.lr_margin
    if spec.top_limit is None and spec.bottom_limit is None:
        return left, right, top, bottom
    try:
        u0, v0, u1, v1 = raw_page.blocks.bbox | raw_page.shapes.bbox
        height = raw_page.height
    except Exception:
        return left, right, top, bottom
    if spec.top_limit is not None and v1 > v0:
        top = max(0.0, min(max(top, spec.top_limit), v0 - 0.5))
    if spec.bottom_limit is not None and v1 > v0:
        bottom = max(0.0, min(max(bottom, height - spec.bottom_limit), height - v1 - 0.5))
    return round(left, 1), round(right, 1), round(top, 1), round(bottom, 1)


_installed = False


def install() -> None:
    global _installed
    if _installed:
        return
    from pdf2docx.page.RawPage import RawPage
    from pdf2docx.page.RawPageFitz import RawPageFitz

    original_images = RawPageFitz._preprocess_images
    original_margin = RawPage.calculate_margin

    def _preprocess_images(self, **settings):
        images = original_images(self, **settings)
        try:
            images = filter_images(self.page_engine, images)
        except Exception:
            log.debug("image filter failed", exc_info=True)
        from .imagetrim import trim_images
        return trim_images(images)       # белые поля картинок (app/imagetrim.py)

    def calculate_margin(self, **settings):
        margin = original_margin(self, **settings)
        try:
            return adjust_margin(self, margin)
        except Exception:
            log.debug("margin adjust failed", exc_info=True)
            return margin

    RawPageFitz._preprocess_images = _preprocess_images
    RawPage.calculate_margin = calculate_margin
    _installed = True
