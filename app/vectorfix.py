"""Патчи pdf2docx для векторной графики: таблицы и рамки листов не превращаются в картинку.

Логика решений — в app/geometry.py (без PyMuPDF, покрыта тестами). Здесь — подмена трёх мест
pdf2docx 0.5.13:

* Path.is_iso_oriented — «прямой» путь проверяется по отрезкам: ломаные и линии с микронаклоном
  из САПР считаются прямыми;
* Path._to_strokes — такие отрезки выпрямляются (pdf2docx требует точной горизонтали/вертикали);
* Paths.to_shapes_and_images — в областях с кривыми картинкой вырезаются только сами кривые
  (логотип, значок), а таблицы внутри рамки листа остаются таблицами.

Любая ошибка в патче возвращает исходное поведение pdf2docx для этой страницы.
"""
from __future__ import annotations

import logging

from . import geometry as geo

log = logging.getLogger(__name__)

_installed = False


def install() -> None:
    global _installed
    if _installed:
        return
    from pdf2docx.shape.Path import C, L, Path
    from pdf2docx.shape.Paths import Paths

    original_iso = Path.is_iso_oriented
    original_to_strokes = Path._to_strokes
    original_to_shapes_and_images = Paths.to_shapes_and_images

    # ------------------------------------------------------------------ прямота пути
    def is_iso_oriented(self) -> bool:
        cached = self.__dict__.get("_p2w_iso")
        if cached is not None:
            return cached
        try:
            ok = True
            if self.is_fill:
                # заливка должна быть прямоугольником (иначе её габарит исказит форму)
                ok = all(s.is_iso_oriented for s in self.items)
            if ok and self.is_stroke:
                ok = geo.items_axis_aligned(self.raw.get("items", ()))
            if not (self.is_fill or self.is_stroke):
                ok = original_iso.fget(self)
        except Exception:
            ok = original_iso.fget(self)
        self.__dict__["_p2w_iso"] = ok
        return ok

    def to_strokes(self, width: float, color):
        try:
            strokes = []
            for segments in self.items:
                for seg in segments:
                    if isinstance(seg, (L, C)):
                        p0, p1 = tuple(seg.points[0]), tuple(seg.points[-1])
                        axis = geo.segment_axis(p0, p1) if isinstance(seg, L) else geo.curve_axis(
                            [tuple(p) for p in seg.points])
                        if axis == "p":
                            continue
                        if axis is None:   # сюда попадают только прямые пути, но на всякий случай
                            strokes.extend(seg.to_strokes(width, color))
                            continue
                        p0, p1 = geo.snap_segment(p0, p1, axis)
                        strokes.extend(L(("l", p0, p1)).to_strokes(width, color))
                    else:
                        strokes.extend(seg.to_strokes(width, color))
            return strokes
        except Exception:
            log.debug("stroke conversion fallback", exc_info=True)
            return original_to_strokes(self, width, color)

    # ------------------------------------------------------------------ таблицы против картинок
    def to_shapes_and_images(self, min_svg_gap_dx: float = 15, min_svg_gap_dy: float = 15,
                             min_w: float = 2, min_h: float = 2, clip_image_res_ratio: float = 3.0):
        try:
            shapes, images = _shapes_and_images(self, min_svg_gap_dx, min_svg_gap_dy, min_w, min_h,
                                                clip_image_res_ratio)
        except Exception:
            log.warning("vector graphics analysis failed, using pdf2docx defaults", exc_info=True)
            shapes, images = original_to_shapes_and_images(self, min_svg_gap_dx, min_svg_gap_dy, min_w,
                                                           min_h, clip_image_res_ratio)
        try:
            from .regionfix import filter_shapes
            return filter_shapes(self.parent.page_engine, shapes, images)
        except Exception:
            log.debug("region shape filter failed", exc_info=True)
            return shapes, images

    Path.is_iso_oriented = property(is_iso_oriented)
    Path._to_strokes = to_strokes
    Paths.to_shapes_and_images = to_shapes_and_images
    _install_rules()
    _installed = True


# --------------------------------------------------------------------------- одиночные линии

MIN_RULE_LEN = 40.0       # pt: короче — не разделительная линия
MAX_RULE_GAP = 30.0       # pt: дальше от текста линию к абзацу не привязываем (предел Word — 31 pt)


def _install_rules() -> None:
    """Одиночные горизонтальные линии вне таблиц (линия под шапкой листа, линия в штампе).

    pdf2docx выбрасывает линию, если она не стала границей таблицы или подчёркиванием. Здесь такая
    линия превращается в нижнюю границу абзаца над ней (или верхнюю — абзаца под ней).
    """
    from pdf2docx.layout.Layout import Layout
    from pdf2docx.text.TextBlock import TextBlock

    original_parse = Layout.parse
    original_make_docx = TextBlock.make_docx

    def parse(self, **settings):
        original_parse(self, **settings)
        try:
            _attach_rules(self)
        except Exception:
            log.debug("standalone rules skipped", exc_info=True)

    def make_docx(self, p):
        res = original_make_docx(self, p)
        for side, width, color, space in getattr(self, "_p2w_rules", ()):
            try:
                _add_paragraph_border(p, side, width, color, space)
            except Exception:
                log.debug("paragraph border failed", exc_info=True)
        return res

    Layout.parse = parse
    TextBlock.make_docx = make_docx


def _rule_geometry(shape) -> tuple[float, float, float, float, int] | None:
    """(x0, x1, y, толщина, цвет) для горизонтальной линии или тонкой заливки-линии."""
    from pdf2docx.shape.Shape import Fill, Stroke

    if isinstance(shape, Stroke):
        if not shape.horizontal:
            return None
        return shape.x0, shape.x1, shape.y0, max(shape.width, 0.25), shape.color
    if isinstance(shape, Fill):
        x0, y0, x1, y1 = shape.bbox
        if y1 - y0 <= 2.0 and x1 - x0 > 10 * (y1 - y0):
            return x0, x1, (y0 + y1) / 2, max(y1 - y0, 0.25), shape.color
    return None


def _attach_rules(layout) -> None:
    blocks = list(layout.blocks)
    texts = [b for b in blocks if b.is_text_block]
    if not texts:
        return
    tables = [tuple(b.bbox) for b in blocks if b.is_table_block]
    x0w, _, x1w, _ = layout.working_bbox
    min_len = max(MIN_RULE_LEN, 0.25 * (x1w - x0w))
    for shape in list(layout.shapes):
        geom = _rule_geometry(shape)
        if geom is None:
            continue
        x0, x1, y, width, color = geom
        if x1 - x0 < min_len:
            continue
        if any(t[0] - 1 <= x0 and x1 <= t[2] + 1 and t[1] - 1 <= y <= t[3] + 1 for t in tables):
            continue   # граница таблицы
        if _is_underline(texts, x0, x1, y):
            continue
        overlap = [b for b in texts if min(x1, b.bbox[2]) - max(x0, b.bbox[0]) > 0.2 * min(x1 - x0, b.bbox[2] - b.bbox[0])]
        above = [b for b in overlap if b.bbox[3] <= y + 1]
        below = [b for b in overlap if b.bbox[1] >= y - 1]
        target = side = None
        if above:
            b = max(above, key=lambda b: b.bbox[3])
            if y - b.bbox[3] <= MAX_RULE_GAP:
                target, side, gap = b, "bottom", y - b.bbox[3]
        if target is None and below:
            b = min(below, key=lambda b: b.bbox[1])
            if b.bbox[1] - y <= MAX_RULE_GAP:
                target, side, gap = b, "top", b.bbox[1] - y
        if target is None:
            continue
        rules = getattr(target, "_p2w_rules", [])
        if not any(r[0] == side for r in rules):
            rules.append((side, width, color, max(0.0, gap - width)))
            target._p2w_rules = rules


def _is_underline(texts, x0: float, x1: float, y: float) -> bool:
    for block in texts:
        for line in block.lines:
            lx0, ly0, lx1, ly1 = line.bbox
            if ly0 <= y <= ly1 + 2.5 and min(x1, lx1) - max(x0, lx0) > 0.5 * (x1 - x0):
                return True
    return False


def _add_paragraph_border(p, side: str, width: float, color: int, space: float) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    ppr = p._p.get_or_add_pPr()
    bdr = ppr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        ppr.insert_element_before(
            bdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
            "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
            "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
            "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(max(2, min(96, round(width * 8)))))
    el.set(qn("w:space"), str(max(0, min(31, round(space)))))
    el.set(qn("w:color"), f"{int(color) & 0xFFFFFF:06X}")
    # порядок внутри pBdr: top, left, bottom, right, between, bar
    if side == "top" and len(bdr):
        bdr.insert(0, el)
    else:
        bdr.append(el)


def _shapes_and_images(paths, min_svg_gap_dx, min_svg_gap_dy, min_w, min_h, clip_image_res_ratio):
    import pymupdf
    from pdf2docx.image.ImagesExtractor import ImagesExtractor

    instances = list(paths)
    if all(p.is_iso_oriented for p in instances):
        return [s for p in instances for s in p.to_shapes()], []

    ie = ImagesExtractor(paths.parent.page_engine)
    groups = ie.detect_svg_contours(min_svg_gap_dx, min_svg_gap_dy, min_w, min_h)

    # путь → (область, вложенная область или None) — как в pdf2docx: первая пересекающаяся область
    buckets: dict[tuple[int, int | None], list] = {}
    for path in instances:
        pb = tuple(path.bbox)
        for gi, (bbox, inner) in enumerate(groups):
            if not pymupdf.Rect(pb).intersects(bbox):
                continue
            key = (gi, None)
            for ii, ib in enumerate(inner):
                if pymupdf.Rect(ib).contains(path.bbox):
                    key = (gi, ii)
                    break
            buckets.setdefault(key, []).append(path)
            break

    shapes, images = [], []
    for (gi, ii), members in buckets.items():
        bbox, inner = groups[gi]
        region = tuple(map(float, bbox if ii is None else inner[ii]))
        plan = geo.plan_region(region, [geo.PathGeom(tuple(p.bbox), p.is_iso_oriented) for p in members])
        for k in plan.shapes:
            shapes.extend(members[k].to_shapes())
        for box in plan.images:
            if box[2] - box[0] < min_w and box[3] - box[1] < min_h:
                continue
            images.append(ie.clip_page_to_dict(bbox=pymupdf.Rect(box), rm_image=True,
                                               clip_image_res_ratio=clip_image_res_ratio))
    return shapes, images
