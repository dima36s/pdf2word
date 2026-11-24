"""Белые поля картинок: рамка картинки — по видимому содержимому.

Word кладёт чертёж с белым фоном так, что его белый край заходит под соседнюю таблицу. Для
pdf2docx картинка и таблица тогда «перекрываются» и не могут стоять рядом в двух колонках: таблица
уезжает под чертёж, лист не помещается на страницу. Белые поля на белом листе не видны, поэтому
их можно отрезать — и картинка, и её рамка становятся такими, какими их видит читатель.
"""
from __future__ import annotations

import io
import logging

log = logging.getLogger(__name__)

WHITE = 250                 # светлее по всем каналам — фон
MIN_TRIM = 0.03             # отрезаем, только если поле не меньше этой доли стороны
MAX_PIXELS = 25_000_000     # огромные сканы не трогаем (и не тратим на них время)


def content_box(img) -> tuple[int, int, int, int] | None:
    """Рамка небелых пикселей (left, top, right, bottom) или None — картинка целиком белая."""
    from PIL import Image, ImageChops

    if img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        img = Image.alpha_composite(bg, rgba)
    rgb = img.convert("RGB")
    r, g, b = ImageChops.invert(rgb).split()
    ink = ImageChops.lighter(ImageChops.lighter(r, g), b).point(lambda v: 255 if v > 255 - WHITE else 0)
    return ink.getbbox()


def trim_white_margins(raw: dict) -> dict:
    """Обрезает белые поля картинки из raw dict pdf2docx и сдвигает её рамку (bbox)."""
    data, bbox = raw.get("image"), raw.get("bbox")
    if not data or not bbox:
        return raw
    from PIL import Image

    img = Image.open(io.BytesIO(data))
    W, H = img.size
    if W * H > MAX_PIXELS or W < 8 or H < 8:
        return raw
    img.load()
    box = content_box(img)
    if box is None:
        return raw
    left, top, right, bottom = max(0, box[0] - 1), max(0, box[1] - 1), min(W, box[2] + 1), min(H, box[3] + 1)
    if (left < MIN_TRIM * W and W - right < MIN_TRIM * W
            and top < MIN_TRIM * H and H - bottom < MIN_TRIM * H):
        return raw
    x0, y0, x1, y1 = bbox
    sx, sy = (x1 - x0) / W, (y1 - y0) / H
    out = io.BytesIO()
    img.crop((left, top, right, bottom)).save(out, format="PNG")
    cw, ch = right - left, bottom - top
    # pdf2docx хранит размер исходной картинки (у повёрнутой — до поворота)
    same = (raw.get("width"), raw.get("height")) == (W, H)
    return {**raw, "image": out.getvalue(),
            "bbox": (x0 + left * sx, y0 + top * sy, x0 + right * sx, y0 + bottom * sy),
            "width": cw if same else ch, "height": ch if same else cw}


PATTERN_MIN = 5.0          # pt: меньше — не плашка, а точка
PATTERN_MAX_SHARE = 0.6    # заливка на весь лист — фон, её не растрируем (текст и линии уже есть)


def _image_info(page) -> list:
    """page.get_image_info без xref-картинок: страница разбирается несколькими проходами (тело,
    шапка, низ, боковой штамп), а отрисовка списка картинок недешёвая — кэш на документе."""
    doc = page.parent
    try:
        cache = doc.__dict__.setdefault("_p2w_pattern_info", {})
    except (AttributeError, TypeError):
        cache = {}
    if page.number not in cache:
        try:
            cache[page.number] = [i for i in page.get_image_info(xrefs=True) if not i.get("xref")]
        except Exception:
            cache[page.number] = []
    return cache[page.number]


def pattern_images(page, images: list, ratio: float) -> list:
    """Градиенты, заливки узором и встроенные (inline) картинки как картинки pdf2docx.

    pdf2docx берёт только картинки-объекты (page.get_images): градиентная плашка под шапкой
    каталога (заливка узором /Pattern) и встроенные картинки пропадали. MuPDF при отрисовке
    сообщает о них как о картинках без xref — эти места рендерятся без текста, как это делает
    сам pdf2docx для векторных рисунков.
    """
    import pymupdf
    from pdf2docx.image.ImagesExtractor import ImagesExtractor

    infos = _image_info(page)
    if not infos:
        return []
    have = [pymupdf.Rect(im["bbox"]) for im in images if im.get("bbox")]
    limit = PATTERN_MAX_SHARE * abs(page.rect)
    out = []
    for info in infos:
        if info.get("xref"):
            continue          # обычная картинка — её уже взял pdf2docx
        r = pymupdf.Rect(info["bbox"]) & page.rect
        if r.is_empty or r.width < PATTERN_MIN or r.height < PATTERN_MIN or abs(r) > limit:
            continue
        if any(abs(r & h) > 0.5 * abs(r) for h in have):
            continue
        try:
            out.append(ImagesExtractor(page).clip_page_to_dict(r, False, ratio))
            have.append(r)
        except Exception:
            log.debug("pattern image skipped", exc_info=True)
    return out


def trim_images(images: list) -> list:
    out = []
    for raw in images:
        try:
            out.append(trim_white_margins(raw))
        except Exception:
            log.debug("image trim failed", exc_info=True)
            out.append(raw)
    return out


# --------------------------------------------------------------------------- поворот без рамки

_installed = False


def install() -> None:
    """Повёрнутые на 90°/180°/270° картинки pdf2docx вращает через cv.warpAffine: интерполяция
    оставляет по краю чёрную линию в пиксель (видна над чертежом в Word). Кратные 90° повороты —
    точной перестановкой пикселей."""
    global _installed
    if _installed:
        return
    from pdf2docx.image.ImagesExtractor import ImagesExtractor

    original = ImagesExtractor._rotate_image

    def rotate_image(pixmap, rotation: int):
        if rotation % 90:
            return original(pixmap, rotation)
        try:
            import cv2 as cv
            import numpy as np

            img = ImagesExtractor._pixmap_to_cv_image(pixmap)
            # положительный угол cv2 — против часовой стрелки, как и np.rot90
            ok, png = cv.imencode(".png", np.ascontiguousarray(np.rot90(img, (rotation // 90) % 4)))
            if ok:
                return png.tobytes()
        except Exception:
            log.debug("exact rotation failed", exc_info=True)
        return original(pixmap, rotation)

    ImagesExtractor._rotate_image = staticmethod(rotate_image)
    _installed = True
