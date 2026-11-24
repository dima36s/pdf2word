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
