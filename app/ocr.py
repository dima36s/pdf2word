"""OCR сканированных страниц.

Страница-скан рендерится в картинку, Tesseract возвращает hOCR (строки, слова, координаты,
оценку кегля, области-иллюстрации). Из этого собирается новая PDF-страница того же размера
с настоящим (видимым) текстом в тех же позициях и с вырезанными иллюстрациями — дальше её
обрабатывает pdf2docx как обычную текстовую страницу: абзацы, отступы, колонки, картинки.
"""
from __future__ import annotations

import io
import math
import os
import re
import shutil
import statistics
import subprocess
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import pymupdf
from PIL import Image

from .config import settings

X_SIZE_TO_PT = 0.91          # x_size из hOCR ≈ 0.91 кегля (откалибровано на Liberation Serif/Sans)
MAX_RENDER_PIXELS = 40_000_000
TESSERACT_TIMEOUT = 180

LANG_NAMES = {
    "rus": "Русский", "eng": "English", "ukr": "Українська", "bel": "Беларуская", "kaz": "Қазақша",
    "deu": "Deutsch", "fra": "Français", "spa": "Español", "ita": "Italiano", "por": "Português",
    "pol": "Polski", "nld": "Nederlands", "tur": "Türkçe", "chi_sim": "中文 (简体)", "chi_tra": "中文 (繁體)",
    "jpn": "日本語", "kor": "한국어", "ara": "العربية", "heb": "עברית", "hin": "हिन्दी", "tha": "ไทย",
    "vie": "Tiếng Việt", "ind": "Bahasa Indonesia",
}

_SERIF_FONTS = [
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/liberation-serif/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/liberation/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    "C:/Windows/Fonts/times.ttf",
    "/Library/Fonts/Times New Roman.ttf",
]


class OcrError(Exception):
    pass


# --------------------------------------------------------------------------- окружение

def tesseract_available() -> bool:
    return settings.ocr_enabled and shutil.which(settings.tesseract_cmd) is not None


@lru_cache(maxsize=1)
def available_languages() -> list[str]:
    if not tesseract_available():
        return []
    try:
        out = subprocess.run([settings.tesseract_cmd, "--list-langs"], capture_output=True,
                             text=True, timeout=20).stdout
    except Exception:
        return []
    langs = [ln.strip() for ln in out.splitlines()[1:] if ln.strip()]
    return [lang for lang in langs if lang not in ("osd", "equ")]


def resolve_lang(lang: str | None) -> str:
    """Оставляет только установленные языки; 'rus+eng' по умолчанию."""
    have = set(available_languages())
    wanted = [p for p in re.split(r"[+,\s]+", lang or settings.ocr_default_lang) if p]
    ok = [p for p in wanted if p in have]
    if not ok:
        ok = [p for p in ("rus", "eng") if p in have] or sorted(have)[:1]
    if not ok:
        raise OcrError("Для Tesseract не установлено ни одного языка.")
    return "+".join(ok)


@lru_cache(maxsize=1)
def ocr_font_file() -> str:
    env = os.environ.get("OCR_FONT")
    if env and Path(env).is_file():
        return env
    for path in _SERIF_FONTS:
        if Path(path).is_file():
            return path
    if shutil.which("fc-match"):
        try:
            out = subprocess.run(["fc-match", "-f", "%{file}", "serif:lang=ru"], capture_output=True,
                                 text=True, timeout=10).stdout.strip()
            if out and Path(out).is_file():
                return out
        except Exception:
            pass
    raise OcrError("Не найден шрифт с кириллицей для распознанного текста (установите fonts-liberation).")


@lru_cache(maxsize=1)
def ocr_bold_font_file() -> str | None:
    env = os.environ.get("OCR_FONT_BOLD")
    if env and Path(env).is_file():
        return env
    reg = Path(ocr_font_file())
    for cand in (reg.name.replace("Regular", "Bold"), reg.stem + "-Bold" + reg.suffix,
                 reg.name.replace(".ttf", "bd.ttf")):
        path = reg.with_name(cand)
        if path != reg and path.is_file():
            return str(path)
    return None


# --------------------------------------------------------------------------- tesseract

def _run_tesseract(png: bytes, args: list[str]) -> bytes:
    env = {**os.environ, "OMP_THREAD_LIMIT": "1"}
    try:
        proc = subprocess.run([settings.tesseract_cmd, "stdin", "stdout", *args], input=png,
                              capture_output=True, timeout=TESSERACT_TIMEOUT, env=env)
    except subprocess.TimeoutExpired as exc:
        raise OcrError("Распознавание страницы заняло слишком много времени.") from exc
    if proc.returncode != 0:
        raise OcrError(proc.stderr.decode(errors="ignore")[-300:] or "tesseract failed")
    return proc.stdout


def _png(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG", compress_level=1)
    return buf.getvalue()


def detect_rotation(gray: Image.Image) -> int:
    """Угол (0/90/180/270), на который надо повернуть страницу по часовой стрелке."""
    small = gray.reduce(2) if min(gray.size) > 1600 else gray
    try:
        out = _run_tesseract(_png(small), ["--psm", "0", "-l", "osd", "--dpi", "150"]).decode(errors="ignore")
    except OcrError:
        return 0
    rot = re.search(r"Rotate:\s*(\d+)", out)
    conf = re.search(r"Orientation confidence:\s*([\d.]+)", out)
    if not rot or not conf or float(conf.group(1)) < 2.0:
        return 0
    return int(rot.group(1)) % 360


# --------------------------------------------------------------------------- линии таблиц

Rect = tuple[float, float, float, float]


def _dilate(mask, k: int, axis: int):
    """Морфологическое расширение булевой маски окном k вдоль оси (через кумулятивные суммы)."""
    import numpy as np
    if k <= 1:
        return mask
    r = k // 2
    pad = [(0, 0), (0, 0)]
    pad[axis] = (r + 1, r)
    cs = np.cumsum(np.pad(mask, pad).astype(np.int32), axis=axis)
    n = mask.shape[axis]
    hi = np.take(cs, range(k, k + n), axis=axis)
    lo = np.take(cs, range(0, n), axis=axis)
    return (hi - lo) > 0


def _erode(mask, k: int, axis: int):
    return ~_dilate(~mask, k, axis)


def _segments(mask, min_len: int, max_thick: int, horizontal: bool) -> list[Rect]:
    """Длинные тонкие отрезки (линии таблиц) в бинарной маске."""
    import numpy as np
    m = mask if horizontal else mask.T
    d = np.diff(np.pad(m, ((0, 0), (1, 1))).astype(np.int8), axis=1)
    rs, cs = np.nonzero(d == 1)
    _, ce = np.nonzero(d == -1)
    keep = (ce - cs) >= min_len
    rs, cs, ce = rs[keep], cs[keep], ce[keep]
    lines: list[list[int]] = []          # [row0, row1, c0, c1]
    active: list[list[int]] = []
    for r, a, b in zip(rs.tolist(), cs.tolist(), ce.tolist()):
        active = [ln for ln in active if ln[1] >= r - 2]
        for ln in active:
            ov = min(b, ln[3]) - max(a, ln[2])
            if ov > 0.5 * min(b - a, ln[3] - ln[2]):
                ln[1] = r
                ln[2], ln[3] = min(a, ln[2]), max(b, ln[3])
                break
        else:
            ln = [r, r, a, b]
            lines.append(ln)
            active.append(ln)
    out = []
    for r0, r1, c0, c1 in lines:
        thick = r1 - r0 + 1
        if thick <= max_thick and (c1 - c0) >= 8 * thick:
            out.append((c0, r0, c1, r1 + 1) if horizontal else (r0, c0, r1 + 1, c1))
    return out


def _touches(a: Rect, b: Rect, tol: float) -> bool:
    return not (a[2] + tol < b[0] or b[2] + tol < a[0] or a[3] + tol < b[1] or b[3] + tol < a[1])


def estimate_skew(gray: Image.Image, max_angle: float = 3.0) -> float:
    """Угол наклона скана (градусы) методом профиля проекции: при верном угле строки
    текста дают самые «резкие» горизонтальные полосы."""
    import numpy as np
    factor = max(1, round(min(gray.size) / 700))
    small = gray.reduce(factor) if factor > 1 else gray
    mask = Image.fromarray(((np.asarray(small) < 160) * 255).astype("uint8"))

    def score(angle: float) -> float:
        prof = np.asarray(mask.rotate(angle, resample=Image.NEAREST, fillcolor=0)).sum(axis=1, dtype=np.float64)
        return float(np.sum(np.diff(prof) ** 2))

    coarse = np.arange(-max_angle, max_angle + 1e-9, 0.25)
    best = max(coarse, key=score)
    fine = np.arange(best - 0.25, best + 0.25 + 1e-9, 0.05)
    best = max(fine, key=score)
    base = score(0.0)
    if base <= 0 or score(best) < base * 1.02:
        return 0.0
    return float(round(best, 2))


@dataclass
class Rules:
    """Линии скана (пиксели): сетка таблиц, одиночные горизонтали, клетки «номера в клетках»."""
    h: list[Rect] = field(default_factory=list)
    v: list[Rect] = field(default_factory=list)
    free: list[Rect] = field(default_factory=list)     # линия под полем формы, разделитель
    combs: list[Rect] = field(default_factory=list)    # клетки «РФ-[5][9]-[2]…» — стираются, не рисуются


FREE_RULE_MIN = 0.5      # дюйма: короче — не линия поля и не разделитель


def detect_table_rules(gray: Image.Image, dpi: float) -> tuple[list[Rect], list[Rect]]:
    """Горизонтальные и вертикальные линии, образующие сетку таблиц (в пикселях)."""
    r = detect_rules(gray, dpi)
    return r.h, r.v


PHOTO_DARK = 185            # порог «тёмного» на выровненном снимке: тонкие линии там серые
PHOTO_DRIFT = 0.02          # дюйма: линия на снимке гуляет (бумага не плоская, края листа выгнуты)


def detect_rules(gray: Image.Image, dpi: float, photo: bool = False, ignore=None) -> Rules:
    """ignore — маска пикселей, которые не бывают линиями (цветные чернила печати на снимке)."""
    import numpy as np
    arr = np.asarray(gray)
    dark = arr < (PHOTO_DARK if photo else 150)
    if ignore is not None:
        dark = dark & ~ignore
    tol = max(2, int(dpi / 75))
    h_mask = _erode(_dilate(dark, tol, 1), tol, 1)       # закрыть разрывы в отсканированных линиях
    v_mask = _erode(_dilate(dark, tol, 0), tol, 0)
    max_thick = max(3, int(dpi * 0.03))
    if photo:
        # линия снимка — «лесенка» со сдвигом на пиксель: размываем поперёк, чтобы ступени слились
        drift = max(2, int(dpi * PHOTO_DRIFT))
        h_mask = _erode(_dilate(_dilate(dark, drift, 0), tol, 1), tol, 1)
        v_mask = _erode(_dilate(_dilate(dark, drift, 1), tol, 0), tol, 0)
        max_thick += 2 * drift           # размытие + изгиб линии к краям листа
    h = _segments(h_mask, int(dpi * 0.35), max_thick, True)
    v = _segments(v_mask, int(dpi * 0.18), max_thick, False)
    combs = find_comb_boxes(dark, dpi, photo)
    if combs:
        h = [a for a in h if not _covered_by(a, combs, tol * 2)]
        v = [b for b in v if not _covered_by(b, combs, tol * 2)]
    # оставляем только линии, образующие сетку: каждая касается хотя бы двух перпендикулярных
    t = tol * 2
    h_keep = [a for a in h if sum(_touches(a, b, t) for b in v) >= 2]
    v_keep = [b for b in v if sum(_touches(a, b, t) for a in h_keep) >= 2]
    h_keep = [a for a in h_keep if sum(_touches(a, b, t) for b in v_keep) >= 2]
    grid = set(h_keep)
    tables = _grid_boxes(h_keep, v_keep, t)
    free = [a for a in h if a not in grid and a[2] - a[0] >= dpi * FREE_RULE_MIN
            and sum(_touches(a, b, t) for b in v_keep) < 2
            and not _covered_by(a, tables, 2 * t)]       # обрывки линий таблицы на скане
    return Rules(h_keep, v_keep, merge_rules(free, tol), combs)


def merge_rules(rules: list[Rect], tol: float) -> list[Rect]:
    """Склеить куски одной горизонтали (скан рвёт линию, и она находится дважды внахлёст)."""
    out: list[list[float]] = []
    for r in sorted(rules, key=lambda r: (round((r[1] + r[3]) / 2 / max(tol, 1)), r[0])):
        cy = (r[1] + r[3]) / 2
        for o in out:
            if abs((o[1] + o[3]) / 2 - cy) <= tol and r[0] <= o[2] + tol and o[0] <= r[2] + tol:
                o[0], o[1], o[2], o[3] = min(o[0], r[0]), min(o[1], r[1]), max(o[2], r[2]), max(o[3], r[3])
                break
        else:
            out.append(list(r))
    return [tuple(o) for o in out]


def _grid_boxes(h: list[Rect], v: list[Rect], tol: float) -> list[Rect]:
    """Рамки таблиц: связные группы касающихся линий сетки."""
    items = list(h) + list(v)
    parent = list(range(len(items)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(len(h)):
        for j in range(len(h), len(items)):
            if _touches(items[i], items[j], tol):
                parent[find(i)] = find(j)
    groups: dict[int, list[Rect]] = {}
    for i, r in enumerate(items):
        groups.setdefault(find(i), []).append(r)
    return [(min(r[0] for r in g), min(r[1] for r in g), max(r[2] for r in g), max(r[3] for r in g))
            for g in groups.values() if len(g) >= 3]


def _covered_by(r: Rect, boxes: list[Rect], tol: float) -> bool:
    return any(b[0] - tol <= r[0] and r[2] <= b[2] + tol and b[1] - tol <= r[1] and r[3] <= b[3] + tol
               for b in boxes)


COMB_MIN, COMB_MAX = 0.10, 0.45      # дюйма: высота клетки «номера в клетках»


def find_comb_boxes(dark, dpi: float, photo: bool = False) -> list[Rect]:
    """Клетки для букв и цифр («РФ-[5][9]-[2]-[0][0]…» в номере ГПЗУ, ИНН, КПП).

    Это отдельные маленькие замкнутые рамки (несколько клеток подряд — одна рамка с перегородками)
    в строку, и почти в каждой есть знак. Tesseract читает их стенки как «[», «|», а сетка
    таблиц собирала из них крошечные таблицы. Возвращает рамки (пиксели)."""
    import cv2
    import numpy as np
    lo, hi = dpi * COMB_MIN, dpi * COMB_MAX
    mask = dark.astype(np.uint8)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    cells: list[tuple[Rect, int, int]] = []          # рамка, число клеток, клеток со знаком
    for k in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[k])
        if not (lo <= h <= hi) or w < 0.55 * h or w > 14 * h:
            continue
        ring = labels[y:y + h, x:x + w] == k
        band = max(2, int(dpi * 0.012))
        if photo:
            band += int(0.025 * w)       # на снимке рамка чуть наклонена: верх и низ гуляют на пиксели
        sides = (ring[:band].any(axis=0).mean(), ring[-band:].any(axis=0).mean(),
                 ring[:, :band].any(axis=1).mean(), ring[:, -band:].any(axis=1).mean())
        if min(sides[2:]) < 0.75 or min(sides[:2]) < (0.5 if photo else 0.75):
            continue                                  # стенки скана бывают с разрывами (на снимке — верх и низ)
        if area > 0.45 * w * h:                       # залитый прямоугольник, а не рамка
            continue
        # перегородки: столбцы, закрашенные почти на всю высоту
        cols = ring.mean(axis=0) >= 0.75
        edges = np.flatnonzero(np.diff(np.concatenate(([0], cols.astype(np.int8), [0]))) != 0)
        walls = [(int(a), int(b)) for a, b in zip(edges[::2], edges[1::2])]
        if len(walls) < 2:
            continue
        inner = dark[y:y + h, x:x + w] & ~ring
        filled = sum(1 for (a0, a1), (b0, _) in zip(walls, walls[1:])
                     if inner[band:-band, a1:b0].sum() > 0.02 * h * max(1, b0 - a1))
        if max(b0 - a1 for (_, a1), (b0, _) in zip(walls, walls[1:])) > 1.8 * h:
            continue                                  # широкая клетка — это таблица, а не клетки для букв
        cells.append(((x, y, x + w, y + h), len(walls) - 1, filled))
    # клетки идут строкой: у строки хотя бы три клетки, и в большинстве — знаки
    out: list[Rect] = []
    for r, n_cells, filled in cells:
        cy, hh = (r[1] + r[3]) / 2, r[3] - r[1]
        row = [c for c in cells if abs((c[0][1] + c[0][3]) / 2 - cy) < 0.3 * hh
               and abs((c[0][3] - c[0][1]) - hh) < 0.25 * hh]
        total = sum(c[1] for c in row)
        if total >= 3 and sum(c[2] for c in row) >= 0.5 * total:
            out.append(r)
    return out


INK_MAX_SHARE = 0.35     # цветных чернил больше — это цветной текст документа, а не печать
INK_MAX_AREA = 0.25      # цветная область больше четверти листа — карта, фото, а не печать
INK_MIN_HEIGHT = 0.3     # дюйма: ниже — строка цветного текста (до ~20 pt); печать 1–2", подпись 0,5–1"


def colored_ink(rgb: Image.Image):
    """Маска «цветных чернил» (синие печати и подписи, красные штампы): насыщенный, не светлый цвет."""
    import cv2
    import numpy as np
    arr = np.asarray(rgb.convert("RGB"))
    hsv = cv2.cvtColor(arr, cv2.COLOR_RGB2HSV)
    sat, val = hsv[..., 1], hsv[..., 2]
    return (sat > 70) & (val < 235) & (val > 30)


def find_ink_marks(rgb: Image.Image, dpi: float, mask=None) -> list[Rect]:
    """Печати, подписи, рукописные пометки цветными чернилами (пиксели).

    Tesseract читает текст печати по кругу как мусор («одия То р ектуры ,,9%»), а подпись
    теряется. Их место — картинка поверх текста; текст документа под печатью распознаётся
    отдельно (цветные чернила со скана для распознавания стираются)."""
    import cv2
    import numpy as np
    if rgb.mode not in ("RGB", "RGBA", "P", "CMYK"):
        return []
    col = colored_ink(rgb) if mask is None else mask
    n_col = int(col.sum())
    if n_col < (dpi * 0.1) ** 2:
        return []
    dark = np.asarray(rgb.convert("L")) < 150
    if n_col / max(1, n_col + int((dark & ~col).sum())) > INK_MAX_SHARE:
        return []
    m = cv2.morphologyEx(col.astype(np.uint8) * 255, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))
    # по горизонтали склеиваем сильнее: цветная строка печатного текста становится одной полосой
    kx, ky = max(3, int(dpi * 0.2)), max(3, int(dpi * 0.06))
    merged = cv2.dilate(m, np.ones((ky, kx), np.uint8))
    n, _, stats, _ = cv2.connectedComponentsWithStats(merged, connectivity=8)
    page = col.shape[0] * col.shape[1]
    out: list[Rect] = []
    for i in range(1, n):
        x, y, w, h, _ = (int(v) for v in stats[i])
        x0, y0, x1, y1 = x + kx // 2, y + ky // 2, x + w - kx // 2, y + h - ky // 2
        bw, bh = x1 - x0, y1 - y0
        if bw * bh > INK_MAX_AREA * page:
            continue
        if bh < dpi * INK_MIN_HEIGHT or (bh < dpi * 0.45 and bw > 6 * bh):
            continue                                  # цветной печатный текст (заголовок, слово) — распознаём
        if m[y0:y1, x0:x1].sum() / 255 < (dpi * 0.1) ** 2 * 0.5:
            continue                                  # редкие цветные точки
        out.append((x0, y0, x1, y1))
    return out


RULE_MAX_PT = 1.5


def erase_line_crumbs(gray: Image.Image, rules: list[Rect], dpi: float, level: int) -> Image.Image:
    """Обрывки изогнутых линий снимка, не попавшие в полосу линии: тонкие горизонтальные штрихи
    рядом с линией таблицы. Tesseract читал их как «- - - - -» в ячейках."""
    import cv2
    import numpy as np
    if not rules:
        return gray
    arr = np.asarray(gray).copy()
    n, labels, stats, _ = cv2.connectedComponentsWithStats((arr < level).astype(np.uint8), connectivity=8)
    thin = max(3, int(dpi * 0.02))
    near = max(4, int(dpi * 0.05))
    for k in range(1, n):
        x, y, w, h, _ = (int(v) for v in stats[k])
        if h > thin or w < 3 * h:
            continue
        cy = y + h / 2
        if any(r[0] - near <= x and x + w <= r[2] + near and r[1] - near <= cy <= r[3] + near for r in rules):
            arr[labels == k] = 255
    return Image.fromarray(arr)


def erase_rules(gray: Image.Image, rules: list[Rect], pad: int = 2) -> Image.Image:
    """Стирает линии таблиц со скана — Tesseract не путает их с буквами «|», «_»."""
    if not rules:
        return gray
    from PIL import ImageDraw
    out = gray.copy()
    dr = ImageDraw.Draw(out)
    for x0, y0, x1, y1 in rules:
        dr.rectangle((x0 - pad, y0 - pad, x1 + pad, y1 + pad), fill=255)
    return out


# --------------------------------------------------------------------------- hOCR

@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    conf: float


@dataclass
class Line:
    bbox: tuple[float, float, float, float]
    baseline: float          # y базовой линии (px)
    x_size: float            # оценка кегля (px)
    words: list[Word]


@dataclass
class Paragraph:
    lines: list[Line] = field(default_factory=list)


@dataclass
class OcrResult:
    paragraphs: list[Paragraph]
    photos: list[tuple[float, float, float, float]]


_XHTML = "{http://www.w3.org/1999/xhtml}"
_LINE_CLASSES = {"ocr_line", "ocr_header", "ocr_textfloat", "ocr_caption"}


def _title(el: ET.Element) -> dict[str, list[str]]:
    res = {}
    for part in (el.get("title") or "").split(";"):
        bits = part.strip().split()
        if bits:
            res[bits[0]] = bits[1:]
    return res


def _bbox(t: dict) -> tuple[float, float, float, float]:
    b = [float(v) for v in t.get("bbox", ["0", "0", "0", "0"])]
    return b[0], b[1], b[2], b[3]


def parse_hocr(data: bytes | str) -> OcrResult:
    if isinstance(data, str):
        data = data.encode("utf-8")
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        # Tesseract.js отдаёт фрагмент (один или несколько <div class='ocr_page'>) без XHTML-обёртки
        root = ET.fromstring(b"<root>" + re.sub(rb"<\?xml[^>]*>|<!DOCTYPE[^>]*>", b"", data) + b"</root>")
    paragraphs: list[Paragraph] = []
    photos = []
    for el in root.iter():
        cls = el.get("class")
        if cls == "ocr_photo":
            photos.append(_bbox(_title(el)))
        elif cls == "ocr_par":
            par = Paragraph()
            for ln in el.iter():
                if ln.get("class") not in _LINE_CLASSES:
                    continue
                t = _title(ln)
                bb = _bbox(t)
                slope, offset = (float(v) for v in (t.get("baseline") or ["0", "0"])[:2])
                mid = (bb[2] - bb[0]) / 2
                baseline = bb[3] + offset + slope * mid
                x_size = float((t.get("x_size") or [bb[3] - bb[1]])[0])
                words = []
                for w in ln.iter():
                    if w.get("class") != "ocrx_word":
                        continue
                    text = "".join(w.itertext()).strip()
                    if not text:
                        continue
                    wt = _title(w)
                    conf = float((wt.get("x_wconf") or ["0"])[0])
                    words.append(Word(text, *_bbox(wt), conf))
                if words:
                    par.lines.append(Line(bb, baseline, x_size, words))
            if par.lines:
                paragraphs.append(par)
    return OcrResult(paragraphs, photos)


# --------------------------------------------------------------------------- кегль по буквам

CAP_EM = 0.67        # высота прописных и цифр в долях кегля (Times/Arial на скане — 0,66…0,72)
FLAT_EM = 0.475      # высота строчных без выносных элементов вверх (0,45…0,52)
SIZE_TRUST = 0.12    # оценка Tesseract расходится с оценкой по буквам больше — берём по буквам
_CAP_CHARS = set("АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789бфbdfhkl")
_FLAT_CHARS = set("авгдежзиклмнопрстухцчшщъыьэюяacemnorsuvwxzgpqy")
_TALL_CHARS = set("()[]{}|/\\\"«»'`!?йёЙЁ№%°^*")


def word_em(w: Word, baseline: float) -> float | None:
    """Кегль слова (px) по высоте букв над базовой линией: прописные и цифры — 0,67 кегля,
    строчные без выносных — 0,475. Слова со скобками, кавычками, «й», «№» не годятся: эти знаки
    выше букв."""
    t = w.text
    if any(c in _TALL_CHARS for c in t) or not any(c.isalnum() for c in t):
        return None
    h = baseline - w.y0
    if h <= 1:
        return None
    if any(c in _CAP_CHARS for c in t):
        return h / CAP_EM
    letters = [c for c in t if c.isalpha()]
    if len(letters) >= 2 and all(c in _FLAT_CHARS for c in letters):
        return h / FLAT_EM
    return None


def refine_sizes(res: OcrResult) -> int:
    """Поправляет x_size строк, где Tesseract ошибся с кеглем.

    На жирных строках скана Tesseract даёт x_size ~7 pt при настоящих 10 pt (заголовки разделов
    ГПЗУ выходили мелким шрифтом), на коротких строках со скобками — наоборот, завышает. Кегль по
    высоте самих букв (медиана по словам строки) устойчив к обоим случаям."""
    fixed = 0
    for par in res.paragraphs:
        for ln in par.lines:
            est = [e for w in ln.words if (e := word_em(w, ln.baseline)) is not None]
            if not est:
                continue
            em = statistics.median(est)
            xs_em = ln.x_size / X_SIZE_TO_PT
            if xs_em <= 0 or abs(em / xs_em - 1) > SIZE_TRUST:
                ln.x_size = em * X_SIZE_TO_PT
                fixed += 1
    return fixed


# --------------------------------------------------------------------------- сборка страницы

_STD_SIZES = [6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72]


def snap_size(pt: float) -> float:
    best = min(_STD_SIZES, key=lambda s: abs(s - pt))
    if abs(best - pt) / max(pt, 1) <= 0.07:
        return float(best)
    return max(4.0, round(pt * 2) / 2)


_JUNK_EDGE = "|_‘`'¦"


def _clean_word(w: Word) -> Word | None:
    """Убирает артефакты сканирования: «|», «_», одиночные штрихи, мусор по краям слова."""
    text = w.text.strip(_JUNK_EDGE) if len(w.text) > 1 else w.text
    if not text:
        return None
    if text in ("—", "–", "‒", "-") and w.conf < 80:
        return None  # «тире» в широком пробеле выравнивания по ширине — артефакт распознавания
    core = re.sub(r"[\W_]+", "", text)
    if not core:
        if text in ("|", "_", "‘", "`", "'", "¦", "~", "^", "°") or w.conf < 60 or len(text) > 3:
            return None
    elif w.conf < 25 and len(core) <= 2 and not core.isdigit():
        return None
    if text != w.text:
        w = Word(text, w.x0, w.y0, w.x1, w.y1, w.conf)
    return w


def _phrases(words: list[Word], fs_px: float, v_rules: list[Rect] = ()) -> list[list[Word]]:
    """Разбивает строку на фрагменты: по вертикальным линиям таблицы и по «столбцовым» промежуткам."""
    if len(words) < 2:
        return [words]
    cy = (min(w.y0 for w in words) + max(w.y1 for w in words)) / 2

    def ruled(a: Word, b: Word) -> bool:
        return any(a.x1 - 2 <= (r[0] + r[2]) / 2 <= b.x0 + 2 and r[1] <= cy <= r[3] for r in v_rules)

    gaps = [b.x0 - a.x1 for a, b in zip(words, words[1:])]
    med = statistics.median(gaps)
    # обычный пробел — по малым промежуткам: в строке «М.П. ……… /Т.Я.Сухарева/» из двух слов
    # медиана — это и есть разрыв между колонками
    space = sorted(gaps)[len(gaps) // 4] if len(gaps) >= 3 else min(gaps[0], fs_px * 0.5)
    limit = max(2.5 * max(space, 1.0), 1.6 * fs_px)
    # строка по ширине: все промежутки растянуты одинаково — не делим. Сравниваем с малым
    # промежутком, не с медианой: в «М.П. (подпись) (расшифровка подписи)» половина промежутков —
    # разрывы между колонками, медиана с ними сама большая, и строка сходила за «по ширине»
    uniform = len(gaps) >= 2 and max(gaps) < 2.5 * max(min(med, space), 1.0)
    out, cur = [], [words[0]]
    for g, a, b in zip(gaps, words, words[1:]):
        if ruled(a, b) or (not uniform and g > limit):
            out.append(cur)
            cur = [b]
        else:
            cur.append(b)
    out.append(cur)
    return out


def _inside(cx: float, cy: float, rects) -> bool:
    return any(r[0] <= cx <= r[2] and r[1] <= cy <= r[3] for r in rects)


_BULLET_LOOKALIKES = {"®", "©", "¢", "*", "°", "e", "o", "«", "»", "+", "•", "·", "●", "@", "€", "е", "о"}


def _fix_bullet(words: list[Word], fs_px: float) -> list[Word]:
    """Tesseract часто читает маркер списка «•» как «®», «©», «e» — исправляем по форме строки."""
    if len(words) >= 2 and words[0].text in _BULLET_LOOKALIKES:
        a, b = words[0], words[1]
        if b.x0 - a.x1 > 0.3 * fs_px and b.text[:1].isupper() and (a.x1 - a.x0) < 0.8 * fs_px:
            words = [Word("•", a.x0, a.y0, a.x1, a.y1, a.conf), *words[1:]]
    return words


BOLD_RATIO = 1.25        # штрих толще обычного на четверть — жирный
BOLD_WORD_RATIO = 1.35   # отдельное слово в обычной строке — с запасом: пикселей меньше


class InkAnalyzer:
    """Оценивает по пикселям скана жирность и цвет слов.

    Толщина штриха — по скелету букв: удвоенное расстояние до края штриха в точках средней линии
    (75-й процентиль по слову). На скане 200–300 dpi она почти не зависит от кегля (её задаёт
    разрешение), поэтому сравнивается в пикселях с обычным текстом страницы; только для текста
    крупнее основного «обычная» толщина растёт пропорционально кеглю."""

    SKELETON_STEPS = 12      # штрих толще 24 px — уже не буква, дальше не утоняем

    def __init__(self, gray: Image.Image, rgb: Image.Image):
        import cv2
        import numpy as np
        self.np = np
        self.dark = np.asarray(gray) < 140
        self.rgb = np.asarray(rgb)
        self.h, self.w = self.dark.shape
        m = self.dark.astype(np.uint8)
        self.dist = cv2.distanceTransform(m, cv2.DIST_L2, 5)
        self.skel = self._skeleton(m)
        self.ref: float | None = None
        self.body: float | None = None

    def _skeleton(self, m):
        import cv2
        np = self.np
        k = cv2.getStructuringElement(cv2.MORPH_CROSS, (3, 3))
        skel = np.zeros_like(m)
        cur = m
        for _ in range(self.SKELETON_STEPS):
            if not cur.any():
                break
            er = cv2.erode(cur, k)
            skel |= cur & ~cv2.dilate(er, k)
            cur = er
        return skel.astype(bool)

    def _box(self, w: Word):
        x0, y0 = max(0, int(w.x0)), max(0, int(w.y0))
        x1, y1 = min(self.w, int(math.ceil(w.x1))), min(self.h, int(math.ceil(w.y1)))
        return x0, y0, x1, y1

    def _samples(self, words: list[Word]):
        vals = []
        for w in words:
            if sum(c.isalpha() for c in w.text) < 2:
                continue
            x0, y0, x1, y1 = self._box(w)
            vals.append(self.dist[y0:y1, x0:x1][self.skel[y0:y1, x0:x1]])
        return self.np.concatenate(vals) if vals else self.np.zeros(0)

    def stroke_px(self, words: list[Word]) -> float | None:
        """Толщина штриха группы слов (px) или None, если букв слишком мало."""
        v = self._samples(words)
        if v.size < 25:
            return None
        return 2.0 * float(self.np.percentile(v, 75))

    def stroke(self, w: Word, x_size: float) -> float | None:
        return self.stroke_px([w])

    def calibrate(self, lines: list[tuple[list[Word], float]]) -> None:
        """lines: (слова строки, кегль строки). Обычная толщина — нижняя четверть по строкам
        основного кегля (на листе может быть больше половины жирных заголовков)."""
        sized = [(px, size, sum(len(w.text) for w in ws)) for ws, size in lines
                 if size > 0 and (px := self.stroke_px(ws)) is not None]
        if len(sized) < 3:
            self.ref = None
            return
        buckets: dict[float, int] = {}
        for _, size, n in sized:
            buckets[round(size)] = buckets.get(round(size), 0) + n
        body = max(buckets.items(), key=lambda kv: kv[1])[0]
        near = sorted(px for px, size, _ in sized if abs(size - body) <= 0.2 * body)
        if len(near) < 3:
            self.ref = None
            return
        self.ref = near[len(near) // 4]
        self.body = float(body)

    def expected(self, size: float) -> float:
        return self.ref * max(1.0, size / self.body) if (self.ref and self.body) else 0.0

    def ratio(self, w: Word, x_size: float) -> float:
        return self.ratio_words([w], x_size)

    def ratio_words(self, words: list[Word], x_size: float) -> float:
        v = self.stroke_px(words)
        exp = self.expected(x_size)
        return v / exp if (v and exp) else 1.0

    def color(self, w: Word) -> tuple[float, float, float] | None:
        x0, y0, x1, y1 = self._box(w)
        mask = self.dark[y0:y1, x0:x1]
        if mask.sum() < 12:
            return None
        px = self.rgb[y0:y1, x0:x1][mask].astype(self.np.float32)
        lum = px.mean(axis=1)
        px = px[lum <= self.np.percentile(lum, 50)]          # самые «насыщенные» пиксели штрихов
        mean = px.mean(axis=0)
        if mean.max() - mean.min() < 45:
            return None                                       # серый/чёрный текст
        return tuple(round(float(c) / 255 / 0.1) * 0.1 for c in mean)


def build_text_page(out_doc: pymupdf.Document, res: OcrResult, img: Image.Image, dpi: float,
                    h_rules: list[Rect] = (), v_rules: list[Rect] = (),
                    ink: InkAnalyzer | None = None, free_rules: list[Rect] = (),
                    marks: list[Rect] = ()) -> int:
    """Добавляет в out_doc страницу с распознанным текстом, иллюстрациями и линиями таблиц.
    Возвращает число слов."""
    s = 72.0 / dpi
    w_pt, h_pt = img.width * s, img.height * s
    page = out_doc.new_page(width=w_pt, height=h_pt)
    page_area = img.width * img.height
    refine_sizes(res)

    # иллюстрации (графики, фото) — вырезаем из скана как есть; печать и подпись — отдельно, ниже
    photos = []
    for x0, y0, x1, y1 in res.photos:
        area = (x1 - x0) * (y1 - y0)
        if area < page_area * 0.004 or area > page_area * 0.6:
            continue
        if any(_overlap((x0, y0, x1, y1), m) > 0.3 * area for m in marks):
            continue
        crop = img.crop((int(x0), int(y0), int(math.ceil(x1)), int(math.ceil(y1))))
        buf = io.BytesIO()
        crop.convert("RGB").save(buf, format="JPEG", quality=88)
        page.insert_image(pymupdf.Rect(x0 * s, y0 * s, x1 * s, y1 * s), stream=buf.getvalue())
        photos.append((x0, y0, x1, y1))

    # линии таблиц — векторными отрезками: по ним pdf2docx строит настоящую таблицу Word
    def _in_photo(r: Rect) -> bool:
        return _inside((r[0] + r[2]) / 2, (r[1] + r[3]) / 2, photos)

    h_rules = [r for r in h_rules if not _in_photo(r)]
    v_rules = [r for r in v_rules if not _in_photo(r)]
    # толщина линии — не толще RULE_MAX_PT: на снимке полоса линии шире самой линии (изгиб листа),
    # и толстая линия задевала первую строку ячейки — pdf2docx выносил её из таблицы
    for x0, y0, x1, y1 in h_rules:
        y = (y0 + y1) / 2 * s
        page.draw_line((x0 * s, y), (x1 * s, y), color=(0, 0, 0), width=min(RULE_MAX_PT, max(0.5, (y1 - y0) * s)))
    for x0, y0, x1, y1 in v_rules:
        x = (x0 + x1) / 2 * s
        page.draw_line((x, y0 * s), (x, y1 * s), color=(0, 0, 0), width=min(RULE_MAX_PT, max(0.5, (x1 - x0) * s)))
    # линии полей формы («Российская Федерация, Пермский край» над «(субъект РФ)») и разделители:
    # дальше они становятся границей абзаца (vectorfix)
    for x0, y0, x1, y1 in free_rules:
        if _in_photo((x0, y0, x1, y1)):
            continue
        y = (y0 + y1) / 2 * s
        page.draw_line((x0 * s, y), (x1 * s, y), color=(0, 0, 0), width=min(RULE_MAX_PT, max(0.5, (y1 - y0) * s)))

    # кегль: медиана по абзацу, затем выравнивание к основному кеглю страницы
    raw_sizes = []
    for par in res.paragraphs:
        xs = [ln.x_size for ln in par.lines if ln.x_size > 0]
        raw = (statistics.median(xs) if xs else 12 / s) * s / X_SIZE_TO_PT
        weight = sum(len(w.text) for ln in par.lines for w in ln.words)
        raw_sizes.append((raw, weight))
    body = None
    if raw_sizes:
        buckets: dict[float, int] = {}
        for raw, wgt in raw_sizes:
            buckets[snap_size(raw)] = buckets.get(snap_size(raw), 0) + wgt
        body = max(buckets.items(), key=lambda kv: kv[1])[0]

    regular = pymupdf.Font(fontfile=ocr_font_file())
    bold_file = ocr_bold_font_file()
    bold = pymupdf.Font(fontfile=bold_file) if bold_file else regular
    if ink is not None:
        ink.calibrate([(ln.words, ln.x_size) for p in res.paragraphs for ln in p.lines])

    # Текст пишется в порядке чтения; при смене цвета начинается новый TextWriter —
    # так порядок в PDF совпадает с порядком строк, и pdf2docx не переставляет фрагменты.
    state = {"tw": None, "color": None}

    def writer(color) -> pymupdf.TextWriter:
        key = color or (0, 0, 0)
        if state["tw"] is None or state["color"] != key:
            if state["tw"] is not None:
                state["tw"].write_text(page)
            state["tw"], state["color"] = pymupdf.TextWriter(page.rect, color=key), key
        return state["tw"]

    # строки абзаца разного кегля (значение поля 10 pt и подпись под ним 8 pt) — кегль у каждой свой
    spread = {}
    for par in res.paragraphs:
        xs = [ln.x_size for ln in par.lines if ln.x_size > 0]
        spread[id(par)] = bool(xs) and max(xs) > 1.12 * min(xs)
    n_words = 0
    for par, (raw, _) in zip(res.paragraphs, raw_sizes):
        par_fs = body if body and abs(raw - body) / body <= 0.1 else snap_size(raw)
        for ln in par.lines:
            # значение поля и подпись под ним Tesseract относит к одному абзацу: кегль строки свой,
            # если он заметно отличается от кегля абзаца
            line_raw = ln.x_size * s / X_SIZE_TO_PT
            if line_raw > 0 and spread[id(par)] and abs(line_raw - raw) / raw > 0.05:
                fs = body if body and abs(line_raw - body) / body <= 0.1 else snap_size(line_raw)
            else:
                fs = par_fs
            words = [cw for k, w in enumerate(ln.words)
                     if (cw := _clean_word(w) or _lone_dash(ln.words, k, fs / s, v_rules)) is not None]
            if not words:
                continue
            cx, cy = (ln.bbox[0] + ln.bbox[2]) / 2, (ln.bbox[1] + ln.bbox[3]) / 2
            if _inside(cx, cy, photos):
                continue  # подписи внутри графика остаются частью картинки
            words = fix_homoglyphs(_fix_bullet(words, fs / s))

            # стиль слов: жирность по плотности «чернил», цвет по пикселям штрихов
            styles = [(False, None)] * len(words)
            if ink is not None and ink.ref:
                alnum = [any(c.isalnum() for c in w.text) for w in words]
                line_bold = ink.ratio_words(words, ln.x_size) >= BOLD_RATIO
                if line_bold:
                    # «Кадастровый номер земельного участка (при наличии) или…» — жирное начало
                    ratios = [ink.ratio(w, ln.x_size) for w in words]
                    plain = sum(1 for r, w in zip(ratios, words) if r < 1.08 and len(w.text) >= 4)
                    if plain >= 3 and plain >= 0.3 * len(words):
                        line_bold = False
                ratios = [ink.ratio(w, ln.x_size) for w in words] if not line_bold else [1.0] * len(words)
                styles = [((line_bold or (r >= BOLD_WORD_RATIO and len(w.text) >= 4)) and a, ink.color(w))
                          for w, r, a in zip(words, ratios, alnum)]

            y = ln.baseline * s
            prev_end = -1e9
            phrases = _phrases(words, fs / s, v_rules)
            idx = 0
            for phrase in phrases:
                # внутри фрагмента группируем подряд идущие слова одного стиля
                runs: list[tuple[tuple, list[Word]]] = []
                for w in phrase:
                    st = styles[idx]
                    idx += 1
                    if runs and runs[-1][0] == st:
                        runs[-1][1].append(w)
                    else:
                        runs.append((st, [w]))
                x = max(phrase[0].x0 * s, prev_end + regular.text_length(" ", fontsize=fs))
                for (is_bold, color), ws in runs:
                    # пробел в конце каждого фрагмента: строки абзаца и ячеек не слипаются в Word
                    text = " ".join(w.text for w in ws) + " "
                    f = bold if is_bold else regular
                    writer(color).append((x, y), text, font=f, fontsize=fs)
                    x += f.text_length(text, fontsize=fs)
                prev_end = x
                n_words += len(phrase)
    if state["tw"] is not None:
        state["tw"].write_text(page)
    # печати и подписи — поверх текста, только цветные пиксели (текст под ними остаётся текстом)
    for rect in marks:
        try:
            png = _ink_png(img, rect)
        except Exception:
            continue
        if png:
            x0, y0, x1, y1 = rect
            page.insert_image(pymupdf.Rect(x0 * s, y0 * s, x1 * s, y1 * s), stream=png)
    return n_words


def _overlap(a: Rect, b: Rect) -> float:
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(0.0, min(a[3], b[3]) - max(a[1], b[1]))


def _ink_png(img: Image.Image, rect: Rect) -> bytes | None:
    """Вырезка печати/подписи: цветные пиксели как есть, остальное прозрачно."""
    import numpy as np
    box = tuple(int(round(v)) for v in rect)
    crop = img.convert("RGB").crop(box)
    alpha = colored_ink(crop)
    if not alpha.any():
        return None
    from PIL import ImageFilter
    a = Image.fromarray((alpha * 255).astype(np.uint8)).filter(ImageFilter.MaxFilter(3))
    rgba = crop.copy()
    rgba.putalpha(a)
    buf = io.BytesIO()
    rgba.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


_LAT2CYR = str.maketrans("ABCEHKMOPTXYaceopxy", "АВСЕНКМОРТХУасеорху")
_HOMOGLYPHS = set("ABCEHKMOPTXYaceopxy")


def _is_cyr(c: str) -> bool:
    return "\u0400" <= c <= "\u04ff"


def fix_homoglyphs(words: list[Word]) -> list[Word]:
    """Латинские двойники русских букв в русской строке → кириллица («OOO» → «ООО», «B границах»
    → «В границах», «Cухаревой» → «Сухаревой»): Tesseract с rus+eng путает их, а в Word такие слова
    не находятся поиском. Слова с латинскими буквами без двойников (Bilt, MPa) не трогаем."""
    letters = [c for w in words for c in w.text if c.isalpha()]
    if not letters:
        return words
    cyr_share = sum(_is_cyr(c) for c in letters) / len(letters)
    out = []
    for w in words:
        lat = [c for c in w.text if c.isalpha() and not _is_cyr(c)]
        if lat and all(c in _HOMOGLYPHS for c in lat):
            mixed = any(_is_cyr(c) for c in w.text)
            if mixed or cyr_share >= 0.85:
                w = Word(w.text.translate(_LAT2CYR), w.x0, w.y0, w.x1, w.y1, w.conf)
        out.append(w)
    return out


_DASH_WORDS = ("—", "–", "‒", "-", "−")


def _lone_dash(words: list[Word], k: int, fs_px: float, v_rules: list[Rect]) -> Word | None:
    """Прочерк «-» один в ячейке таблицы — это «нет значения», его не выбрасываем как артефакт
    широкого пробела строки по ширине."""
    w = words[k]
    if w.text not in _DASH_WORDS:
        return None
    cy = (w.y0 + w.y1) / 2

    def apart(a: Word, b: Word) -> bool:
        ruled = any(a.x1 - 2 <= (r[0] + r[2]) / 2 <= b.x0 + 2 and r[1] <= cy <= r[3] for r in v_rules)
        return ruled or b.x0 - a.x1 > 3.0 * fs_px

    alone = len(words) == 1
    left = k == 0 or apart(words[k - 1], w)
    right = k == len(words) - 1 or apart(w, words[k + 1])
    in_cell = any(r[1] <= cy <= r[3] for r in v_rules)
    return w if (alone and in_cell) or (left and right and not alone) else None


# --------------------------------------------------------------------------- точка входа

PHOTO_PAGE_MIN = 1370.0     # pt: длинная сторона больше A3 (1191 pt) с запасом
PHOTO_MAX_PPI = 100.0       # снимок, у которого пиксель ≈ пункт листа: размер листа ничего не значит
PHOTO_LONG_IN = 842 / 72    # такой снимок — лист A4 (длинная сторона, дюймы)
PHOTO_UPSCALE = 1.5         # не увеличивать снимок больше чем в полтора раза
PHOTO_EDGE = 0.02           # доля стороны снимка у края: стол и тень за краем листа (поля листа шире)


def photo_page_inches(page: pymupdf.Page) -> tuple[float, float] | None:
    """Снимок листа с телефона (iOS «Сканировать документы», Quartz PDFContext): страница — одна
    картинка, а размер страницы равен её размеру в пикселях (72 «dpi»), 86 × 60 см вместо A4.

    По такому размеру распознавание рендерило 40 млн пикселей (2 ГБ памяти, в браузере — отказ),
    пороги в дюймах промахивались в три раза, а в Word выходил лист 86 × 60 см с текстом 30 pt.
    Возвращает настоящий размер листа в дюймах (A4 с пропорциями снимка) или None."""
    r = page.rect
    if max(r.width, r.height) < PHOTO_PAGE_MIN or page.get_text("text").strip():
        return None
    try:
        images = page.get_images(full=True)
        if len(images) != 1:
            return None
        bbox = page.get_image_bbox(images[0])
        width_px = images[0][2]
    except Exception:
        return None
    if bbox.is_empty or bbox.width * bbox.height < 0.9 * r.width * r.height:
        return None
    ppi = width_px / (bbox.width / 72)
    if ppi > PHOTO_MAX_PPI:
        return None
    long_in = PHOTO_LONG_IN
    short_in = long_in * min(r.width, r.height) / max(r.width, r.height)
    return (long_in, short_in) if r.width >= r.height else (short_in, long_in)


def render_page(page: pymupdf.Page, dpi: int) -> tuple[Image.Image, float]:
    w_in, h_in = page.rect.width / 72, page.rect.height / 72
    phys = photo_page_inches(page)
    if phys is not None:
        # рендер в пикселях настоящего листа: dpi листа → масштаб страницы, но не крупнее снимка × 1,5
        scale = min(dpi * phys[0] / w_in, PHOTO_UPSCALE * _native_ppi(page))
        pix = page.get_pixmap(dpi=max(int(scale), 1), alpha=False, colorspace=pymupdf.csRGB)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return img, pix.width / phys[0]
    dpi = min(dpi, int(math.sqrt(MAX_RENDER_PIXELS / max(w_in * h_in, 0.01))))
    pix = page.get_pixmap(dpi=dpi, alpha=False, colorspace=pymupdf.csRGB)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    # фактический dpi по ширине (get_pixmap округляет размеры)
    return img, pix.width / w_in


def _native_ppi(page: pymupdf.Page) -> float:
    img = page.get_images(full=True)[0]
    return img[2] / (page.get_image_bbox(img).width / 72)


@dataclass
class PreparedScan:
    """Страница-скан, подготовленная к распознаванию."""
    img: Image.Image          # цветная, повёрнутая и выровненная
    clean: Image.Image        # серая, без линий таблиц — её распознаёт Tesseract
    dpi: float
    h_rules: list[Rect]
    v_rules: list[Rect]
    free_rules: list[Rect] = field(default_factory=list)   # линии полей формы, разделители
    marks: list[Rect] = field(default_factory=list)        # печати и подписи цветными чернилами
    combs: list[Rect] = field(default_factory=list)        # «номер в клетках»: рамки групп клеток


def flatten_light(img: Image.Image) -> Image.Image:
    """Снимок листа телефоном: свет неровный (тень от руки, серый угол, зелёный стол по краю), тонкие
    линии таблиц светлее порога «тёмного». Каждый канал делится на фон — его даёт морфологическое
    закрытие в уменьшенной копии (текст и линии тоньше окна, фон — нет): бумага становится белой,
    чернила сохраняют цвет, поле за краем листа исчезает."""
    import cv2
    import numpy as np
    arr = np.asarray(img).astype(np.float32)
    h, w = arr.shape[:2]
    f = max(1, min(h, w) // 400)
    small = cv2.resize(arr, (max(1, w // f), max(1, h // f)), interpolation=cv2.INTER_AREA)
    k = max(7, (min(small.shape[:2]) // 25) | 1)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k))
    bg = cv2.morphologyEx(small, cv2.MORPH_CLOSE, kernel)
    bg = cv2.GaussianBlur(bg, (0, 0), k / 3)
    bg = cv2.resize(bg, (w, h), interpolation=cv2.INTER_LINEAR)
    out = np.clip(arr / np.maximum(bg, 1.0) * 255.0, 0, 255).astype(np.uint8)
    # край снимка — стол, тень, край листа: текста там нет (поля листа шире), заливаем белым
    by, bx = max(2, int(h * PHOTO_EDGE)), max(2, int(w * PHOTO_EDGE))
    out[:by], out[-by:], out[:, :bx], out[:, -bx:] = 255, 255, 255, 255
    return Image.fromarray(out)


DEWARP_MIN_SPREAD = 0.3     # градусов: наклон линий меняется по листу — бумага изогнута
DEWARP_MAX_ANGLE = 4.0


def _line_angles(dark, dpi: float, vertical: bool):
    """Длинные почти горизонтальные (или вертикальные) отрезки: (координата середины поперёк, угол)."""
    import cv2
    import numpy as np
    m = dark.T if vertical else dark
    found = cv2.HoughLinesP(np.ascontiguousarray(m), 1, np.pi / 1800, threshold=int(dpi),
                            minLineLength=int(dpi * 1.5), maxLineGap=int(dpi * 0.06))
    out = []
    if found is None:
        return out
    for x1, y1, x2, y2 in found.reshape(-1, 4).tolist():
        if x2 == x1:
            continue
        a = math.degrees(math.atan2(y2 - y1, x2 - x1))
        if abs(a) <= DEWARP_MAX_ANGLE:
            out.append(((y1 + y2) / 2, a))
    return out


def _angle_profile(samples, length: int, step: int):
    """Наклон как функция координаты: медиана по полосам, между полосами — линейно."""
    import numpy as np
    if len(samples) < 4:
        return None
    angles = [a for _, a in samples]
    if max(angles) - min(angles) < DEWARP_MIN_SPREAD:
        return None
    centers, values = [], []
    for start in range(0, length, step):
        band = [a for c, a in samples if start <= c < start + step]
        if band:
            centers.append(start + step / 2)
            values.append(float(np.median(band)))
    if len(centers) < 2:
        return None
    return np.interp(np.arange(length), centers, values)


def dewarp_photo(img: Image.Image, gray: Image.Image, dpi: float) -> tuple[Image.Image, Image.Image]:
    """Снимок изогнутого листа: линии таблицы наверху наклонены на 2°, в середине ровные, внизу —
    в другую сторону. Общий поворот этого не исправит, и линии таблиц не находились (каждая —
    «лесенка» толще допуска). Наклон горизонталей измеряется по высоте листа, вертикалей — по
    ширине, и снимок выпрямляется сдвигом строк и столбцов (cv2.remap)."""
    import cv2
    import numpy as np
    arr = np.asarray(gray)
    dark = ((arr < PHOTO_DARK) * 255).astype(np.uint8)
    h, w = arr.shape
    step = max(1, int(dpi))
    rows = _angle_profile(_line_angles(dark, dpi, False), h, step)
    cols = _angle_profile(_line_angles(dark, dpi, True), w, step)
    if rows is None and cols is None:
        return img, gray
    ys, xs = np.mgrid[0:h, 0:w].astype(np.float32)
    map_y = ys + ((xs - w / 2) * np.tan(np.radians(rows))[:, None]).astype(np.float32) if rows is not None else ys
    map_x = xs + ((ys - h / 2) * np.tan(np.radians(cols))[None, :]).astype(np.float32) if cols is not None else xs
    del ys, xs
    out_gray = cv2.remap(arr, map_x, map_y, cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT, borderValue=255)
    out_img = cv2.remap(np.asarray(img), map_x, map_y, cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT,
                        borderValue=(255, 255, 255))
    return Image.fromarray(out_img), Image.fromarray(out_gray)


TRACK_STEP = 8          # px: шаг отслеживания линии
TRACK_WIN = 3           # px: на сколько линия может сместиться за шаг
TRACK_DARK = 200        # средняя яркость окна на линии
TRACK_MISS = 6          # шагов подряд без линии — конец линии
TRACK_GRID = 16         # px: шаг сетки поля смещений
TRACK_MIN_SHARE = 0.35  # доля ширины листа: короче — подчёркивание, а не линия таблицы


def _track_line(arr, x: int, y: int, direction: int):
    """Проходит тёмную горизонтальную линию от (x, y) влево или вправо: [(x, y)]."""
    h, w = arr.shape
    pts, miss = [], 0
    while 0 <= x < w and miss <= TRACK_MISS:
        y0, y1 = max(0, y - TRACK_WIN), min(h, y + TRACK_WIN + 1)
        col = arr[y0:y1, max(0, x - TRACK_STEP // 2):x + TRACK_STEP // 2 + 1].mean(axis=1)
        k = int(col.argmin())
        if col[k] < TRACK_DARK:
            y, miss = y0 + k, 0
            pts.append((x, y))
        else:
            miss += 1
        x += direction * TRACK_STEP
    return pts


def curve_field(gray: Image.Image, dpi: float):
    """Поле сдвигов по вертикали, выпрямляющее изогнутые горизонтали (край листа загнут: линия
    таблицы к правому краю уходит на 20 пикселей вверх, и наклоном это не исправить). Каждая
    длинная линия прослеживается по всей длине, её точки встают на общую высоту; между линиями —
    линейно. None — линий мало."""
    import cv2
    import numpy as np
    arr = np.asarray(gray).astype(np.float32)
    h, w = arr.shape
    dark = ((arr < PHOTO_DARK) * 255).astype(np.uint8)
    seeds = sorted(_line_seeds(dark, dpi), key=lambda t: t[1])
    tracks = []
    for x, y in seeds:
        if any(abs(t_y(t, x) - y) < 8 for t in tracks):
            continue
        pts = sorted(_track_line(arr, x, y, -1) + _track_line(arr, x + TRACK_STEP, y, 1))
        if pts and pts[-1][0] - pts[0][0] >= TRACK_MIN_SHARE * w:
            tracks.append(pts)
    if len(tracks) < 3:
        return None
    gx = np.arange(0, w + TRACK_GRID, TRACK_GRID)
    gy = np.arange(0, h + TRACK_GRID, TRACK_GRID)
    rows = []
    for pts in tracks:
        xs = np.array([p[0] for p in pts], np.float32)
        ys = np.array([p[1] for p in pts], np.float32)
        ref = float(np.median(ys))
        rows.append((ref, np.interp(gx, xs, ys), xs[0], xs[-1]))
    rows.sort(key=lambda r: r[0])
    grid = np.empty((len(gy), len(gx)), np.float32)
    for j, x in enumerate(gx):
        # только линии, что доходят до этого столбца: за концом линии её высота ничего не говорит
        here = [(r[0], r[1][j]) for r in rows if r[2] - TRACK_GRID <= x <= r[3] + TRACK_GRID]
        if not here:
            grid[:, j] = gy
            continue
        refs = np.array([a for a, _ in here])
        off = np.array([b - a for a, b in here])
        grid[:, j] = gy + np.interp(gy, refs, off)
    grid = cv2.GaussianBlur(grid - gy[:, None].astype(np.float32), (0, 0), 2) + gy[:, None]
    return cv2.resize(grid.astype(np.float32), (w, h), interpolation=cv2.INTER_LINEAR)


def t_y(track, x: float) -> float:
    xs = [p[0] for p in track]
    ys = [p[1] for p in track]
    import numpy as np
    return float(np.interp(x, xs, ys))


def _line_seeds(dark, dpi: float):
    """Середины длинных почти горизонтальных отрезков — откуда отслеживать линии."""
    import cv2
    import numpy as np
    found = cv2.HoughLinesP(dark, 1, np.pi / 1800, threshold=int(dpi), minLineLength=int(dpi * 1.5),
                            maxLineGap=int(dpi * 0.06))
    out = []
    if found is None:
        return out
    for x1, y1, x2, y2 in found.reshape(-1, 4).tolist():
        if x2 != x1 and abs(y2 - y1) <= abs(x2 - x1) * 0.07:
            out.append(((x1 + x2) // 2, (y1 + y2) // 2))
    return out


def straighten_rows(img: Image.Image, gray: Image.Image, dpi: float) -> tuple[Image.Image, Image.Image]:
    import cv2
    import numpy as np
    field = curve_field(gray, dpi)
    if field is None:
        return img, gray
    h, w = field.shape
    map_x = np.tile(np.arange(w, dtype=np.float32), (h, 1))
    out_gray = cv2.remap(np.asarray(gray), map_x, field, cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT,
                         borderValue=255)
    out_img = cv2.remap(np.asarray(img), map_x, field, cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT,
                        borderValue=(255, 255, 255))
    return Image.fromarray(out_img), Image.fromarray(out_gray)


def prepare_scan(page: pymupdf.Page, dpi: int | None = None, detect_orientation: bool = True) -> PreparedScan:
    from PIL import ImageFilter

    img, real_dpi = render_page(page, dpi or settings.ocr_dpi)
    photo = photo_page_inches(page) is not None
    if photo:
        try:
            img = flatten_light(img)
        except Exception:
            photo = False
    gray = img.convert("L").filter(ImageFilter.MedianFilter(3))  # убирает «соль и перец» скана
    rotation = detect_rotation(gray) if detect_orientation and tesseract_available() else 0
    if rotation:
        img = img.rotate(-rotation, expand=True)
        gray = gray.rotate(-rotation, expand=True)
    skew = estimate_skew(gray)
    if abs(skew) >= 0.1:
        img = img.rotate(skew, resample=Image.BICUBIC, fillcolor=(255, 255, 255))
        gray = gray.rotate(skew, resample=Image.BICUBIC, fillcolor=255)
    if photo:
        try:
            img, gray = dewarp_photo(img, gray, real_dpi)
            img, gray = straighten_rows(img, gray, real_dpi)
        except Exception:
            pass                         # выпрямление — улучшение, не повод терять страницу
    ink = None
    if photo:
        try:
            ink = colored_ink(img)       # серый порог снимка выше — печать и подпись сошли бы за линии
        except Exception:
            ink = None
    rules = detect_rules(gray, real_dpi, photo=photo, ignore=ink)
    clean = erase_rules(gray, rules.h + rules.v + rules.free)
    if photo:
        try:
            clean = erase_line_crumbs(clean, rules.h + rules.free, real_dpi, PHOTO_DARK)
        except Exception:
            pass
    if rules.combs:
        clean = erase_comb_frames(clean, rules.combs, PHOTO_DARK if photo else 150)
    marks: list[Rect] = []
    try:
        ink = colored_ink(img)
        marks = find_ink_marks(img, real_dpi, ink)
        if marks:
            clean = _erase_ink(clean, img, ink, marks, real_dpi)
    except Exception:          # печати — улучшение, не повод терять страницу
        marks = []
    return PreparedScan(img, clean, real_dpi, rules.h, rules.v, rules.free, marks, rules.combs)


def erase_comb_frames(gray: Image.Image, combs: list[Rect], level: int = 150) -> Image.Image:
    """Стереть рамки клеток, оставив знаки в них: в каждой вырезке убирается самая большая связная
    область — сама рамка с перегородками."""
    import cv2
    import numpy as np
    arr = np.asarray(gray).copy()
    for x0, y0, x1, y1 in combs:
        x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
        dark = (arr[y0:y1, x0:x1] < level).astype(np.uint8)
        n, labels, stats, _ = cv2.connectedComponentsWithStats(dark, connectivity=8)
        if n < 2:
            continue
        frame = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
        grow = 3 if level <= 150 else 7          # на снимке у рамки светлая кайма — стираем шире
        ring = cv2.dilate((labels == frame).astype(np.uint8), np.ones((grow, grow), np.uint8)) > 0
        arr[y0:y1, x0:x1][ring] = 255
    return Image.fromarray(arr)


def _erase_ink(gray: Image.Image, rgb: Image.Image, ink, marks: list[Rect], dpi: float,
               black_level: int = 110) -> Image.Image:
    """Стереть печати и подписи со скана для распознавания.

    Края штрихов цветных чернил светлее и почти серые — по насыщенности их не видно, а оставшиеся
    «тени» Tesseract принимает за картинку и не читает чёрный текст под печатью («М.П.», «Дата
    выдачи»). Поэтому рядом с цветными чернилами стирается всё, кроме тёмного ненасыщенного —
    чёрного текста."""
    import cv2
    import numpy as np
    arr = np.asarray(gray).copy()
    hsv = cv2.cvtColor(np.asarray(rgb.convert("RGB")), cv2.COLOR_RGB2HSV)
    k = max(3, int(dpi * 0.012)) | 1
    near = cv2.dilate(ink.astype(np.uint8), np.ones((k, k), np.uint8)) > 0
    black = (hsv[..., 2] < black_level) & (hsv[..., 1] < 70)   # на снимке «чёрный» светлее
    for x0, y0, x1, y1 in marks:
        x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
        zone = near[y0:y1, x0:x1] & ~black[y0:y1, x0:x1]
        arr[y0:y1, x0:x1][zone] = 255
        # бледные следы печати (светлее текста) во всей её рамке
        sub = arr[y0:y1, x0:x1]
        sub[(sub > 150) & ~black[y0:y1, x0:x1]] = 255
    return Image.fromarray(arr)


def scan_png(prep: PreparedScan) -> bytes:
    return _png(prep.clean)


PICTURE_MIN_WORDS = 8         # меньше осмысленных слов — распознавать нечего
PICTURE_PHOTO_SHARE = 0.5     # фото занимает полстраницы и больше…
PICTURE_PHOTO_WORDS = 40      # …а текста как на обложке, а не на странице документа
PICTURE_GARBLED_SHARE = 0.35  # осмысленных слов меньше трети (на тексте — 80–90 %)…
PICTURE_GARBLED_WORDS = 60    # …и немного: это чертёж или карта с повёрнутыми подписями


def looks_like_picture(res: OcrResult, width: float, height: float) -> bool:
    """Страница — иллюстрация (обложка, фотоколлаж), а не скан текста.

    Такую страницу нельзя заменять распознанным текстом: фото выпало бы (оно больше 60 % листа),
    а вместо него остались бы несколько слов, прочитанных с картинки.
    """
    words = [w for p in res.paragraphs for ln in p.lines for w in ln.words]
    good = sum(1 for w in words if w.conf >= 60 and sum(c.isalpha() for c in w.text) >= 2)
    if good < PICTURE_MIN_WORDS:
        return True
    # чертёж, карта, схема: «слова» — это повёрнутые подписи и штриховка, осмысленных почти нет
    meaningful = sum(1 for w in words if w.conf >= 60 and (sum(c.isalpha() for c in w.text) >= 2
                                                          or sum(c.isdigit() for c in w.text) >= 2))
    if meaningful < PICTURE_GARBLED_SHARE * len(words) and meaningful < PICTURE_GARBLED_WORDS:
        return True
    area = max(width * height, 1.0)
    photo = max(((x1 - x0) * (y1 - y0) for x0, y0, x1, y1 in res.photos), default=0.0)
    return photo / area >= PICTURE_PHOTO_SHARE and good < PICTURE_PHOTO_WORDS


def add_cell_dashes(res: OcrResult, prep: PreparedScan) -> int:
    """Прочерки «-» в пустых ячейках таблиц: Tesseract на всей странице такие одиночные штрихи
    не выдаёт, а в таблице это «нет значения» (ГПЗУ: строка «- | - | -»)."""
    import cv2
    import numpy as np
    if not prep.h_rules or not prep.v_rules:
        return 0
    dpi = prep.dpi
    tables = _grid_boxes(prep.h_rules, prep.v_rules, max(4, dpi / 37))
    if not tables:
        return 0
    words = [w for p in res.paragraphs for ln in p.lines for w in ln.words]
    sizes = [ln.x_size for p in res.paragraphs for ln in p.lines if ln.x_size > 0]
    x_size = statistics.median(sizes) if sizes else dpi * 10 / 72 * X_SIZE_TO_PT
    dark = (np.asarray(prep.clean) < 140).astype(np.uint8)
    n, _, stats, _ = cv2.connectedComponentsWithStats(dark, connectivity=8)
    added = 0
    for k in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[k])
        if not (dpi * 0.025 <= w <= dpi * 0.15 and h <= dpi * 0.025 and w >= 2.5 * h and area >= 0.6 * w * h):
            continue
        cx, cy = x + w / 2, y + h / 2
        if not any(t[0] < cx < t[2] and t[1] < cy < t[3] for t in tables):
            continue
        if any(wd.x0 - 2 <= cx <= wd.x1 + 2 and wd.y0 - 2 <= cy <= wd.y1 + 2 for wd in words):
            continue
        near = dpi * 0.045
        if any(r[0] - 2 <= cx <= r[2] + 2 and abs((r[1] + r[3]) / 2 - cy) < near for r in prep.h_rules) \
                or any(t[0] < cx < t[2] and min(abs(cy - t[1]), abs(cy - t[3])) < near for t in tables):
            continue                      # огрызок стёртой линии таблицы у самой границы ячейки
        em = x_size / X_SIZE_TO_PT
        word = Word("-", float(x), cy - 0.3 * em, float(x + w), cy + 0.1 * em, 90.0)
        line = Line((word.x0, word.y0, word.x1, word.y1), cy + 0.25 * em, x_size, [word])
        res.paragraphs.append(Paragraph([line]))
        added += 1
    return added


COMB_CHARS = "0123456789АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ"
_COMB_DIGITS = str.maketrans({"О": "0", "о": "0", "З": "3", "з": "3", "Б": "6", "Ч": "4"})


def _comb_clean(text: str) -> str:
    text = re.sub(r"[^0-9A-ZА-ЯЁ]", "", text.upper())
    if any(c.isdigit() for c in text) or (text and set(text) <= set("ОЗ")):
        text = text.translate(_COMB_DIGITS)
    return text


def _comb_cells(rect: Rect) -> int | None:
    ratio = (rect[2] - rect[0]) / max(rect[3] - rect[1], 1)
    return round(ratio) if abs(ratio - round(ratio)) <= 0.3 else None


def read_combs(prep: PreparedScan, lang: str) -> dict[Rect, list[str]]:
    """Каждая группа клеток («[5][9]», «[2][0][2][6]») читается отдельно, как одно слово, —
    несколькими способами (слово/строка, 1× и 2×); выбор — в put_combs голосованием вместе с общим
    проходом.

    Знаки в клетках стоят вразрядку, и на снимке Tesseract по всей странице читал их со стенками и
    разрывами («[519]», «92/904)»), а на скане общий проход бывает верен, а отдельное чтение
    ошибается («00» → «60»). Только с локальным Tesseract (в браузере группы читает общий проход)."""
    out: dict[Rect, list[str]] = {}
    for rect in prep.combs:
        x0, y0, x1, y1 = rect
        inset = max(3, int((y1 - y0) * 0.08))
        crop = prep.clean.crop((int(x0) + inset, int(y0) + inset, int(x1) - inset, int(y1) - inset))
        bw = crop.point(lambda v: 0 if v < 170 else 255)
        found = []
        for psm, scale in (("8", 1), ("8", 2), ("7", 1), ("7", 2)):
            im = bw.resize((bw.width * scale, bw.height * scale), Image.LANCZOS) if scale > 1 else bw
            pad = Image.new("L", (im.width + 40, im.height + 40), 255)
            pad.paste(im, (20, 20))
            try:
                raw = _run_tesseract(_png(pad), ["-l", lang, "--dpi", str(int(prep.dpi * scale)), "--psm", psm,
                                                 "-c", f"tessedit_char_whitelist={COMB_CHARS}"])
            except Exception:
                continue
            found.append(_comb_clean(raw.decode("utf-8", "ignore")))
        out[rect] = found
    return out


COMB_PRIMARY = 3          # общий проход и чтение «словом» (1× и 2×); «строкой» — запасное:
                          # оно чаще путает последнюю цифру («0076» → «0070»)


def vote_comb(rect: Rect, candidates: list[str]) -> str | None:
    """Чтение группы: самое частое из прочтений с числом знаков, равным числу клеток (при равенстве —
    раньше в списке: общий проход, затем «слово 1×»); прочтения «строкой» — только если среди
    основных подходящих нет."""
    cells = _comb_cells(rect)

    def fits(c: str) -> bool:
        return bool(c) and (cells is None or len(c) == cells)

    for group in (candidates[:COMB_PRIMARY], candidates):
        good = [c for c in group if fits(c)]
        if good:
            return max(good, key=lambda c: (good.count(c), -good.index(c)))
    rest = [c for c in candidates if c]
    return rest[0] if rest else None


def comb_dashes(prep: PreparedScan) -> list[Rect]:
    """Чёрточки между группами клеток («РФ - 59 - 2»): короткий тёмный штрих посередине промежутка."""
    import numpy as np
    arr = np.asarray(prep.clean)
    out = []
    rows: list[list[Rect]] = []
    for r in sorted(prep.combs, key=lambda r: r[0]):
        for row in rows:
            if abs((row[-1][1] + row[-1][3]) / 2 - (r[1] + r[3]) / 2) < 0.4 * (r[3] - r[1]):
                row.append(r)
                break
        else:
            rows.append([r])
    for row in rows:
        for a, b in zip(row, row[1:]):
            gx0, gx1 = int(a[2]) + 2, int(b[0]) - 2
            if gx1 - gx0 < 4 or gx1 - gx0 > 3 * (a[3] - a[1]):
                continue
            gy0, gy1 = int(max(a[1], b[1])), int(min(a[3], b[3]))
            dark = arr[gy0:gy1, gx0:gx1] < 170
            ys, xs = np.nonzero(dark)
            if len(xs) < 4:
                continue
            w, hgt = xs.max() - xs.min() + 1, ys.max() - ys.min() + 1
            if w >= 1.5 * hgt and hgt <= 0.25 * (gy1 - gy0):
                out.append((gx0 + xs.min(), gy0 + ys.min(), gx0 + xs.max() + 1, gy0 + ys.max() + 1))
    return out


def put_combs(res: OcrResult, combs: dict[Rect, list[str] | str]) -> None:
    """Слова, распознанные внутри групп клеток, заменяются прочитанными по группам (голосование
    отдельных прочтений и общего прохода); «-» между группами — как есть."""
    if not combs:
        return

    def within(w: Word, r: Rect) -> bool:
        cx, cy = (w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2
        return r[0] - 4 <= cx <= r[2] + 4 and r[1] - 4 <= cy <= r[3] + 4

    words = [w for p in res.paragraphs for ln in p.lines for w in ln.words]
    boxes = [r for r, c in combs.items() if not isinstance(c, str)]
    # общий проход: слово «-59-2-00-0-» на несколько групп делится по чёрточкам между группами
    full: dict[Rect, str] = {}
    for w in sorted(words, key=lambda w: w.x0):
        cy = (w.y0 + w.y1) / 2
        over = sorted((r for r in boxes if r[1] - 4 <= cy <= r[3] + 4 and min(w.x1, r[2]) - max(w.x0, r[0]) > 0),
                      key=lambda r: r[0])
        if not over:
            continue
        parts = [p for p in re.split(r"[^0-9A-Za-zА-Яа-яЁё]+", w.text) if p]
        if len(over) == 1:
            full[over[0]] = full.get(over[0], "") + "".join(parts)
        elif len(parts) == len(over):
            for r, part in zip(over, parts):
                full[r] = full.get(r, "") + part
    chosen: dict[Rect, str] = {}
    for r, cand in combs.items():
        if isinstance(cand, str):
            chosen[r] = cand
            continue
        text = vote_comb(r, [_comb_clean(full.get(r, ""))] + list(cand))
        if text:
            chosen[r] = text
    combs = chosen
    if not combs:
        return
    rects = list(combs)

    def inside(w: Word) -> bool:
        """Слово внутри группы или поперёк нескольких групп («-2-0О-»)."""
        cy = (w.y0 + w.y1) / 2
        over = [min(w.x1, r[2]) - max(w.x0, r[0]) for r in rects if r[1] - 4 <= cy <= r[3] + 4]
        over = [o for o in over if o > 0]
        return any(within(w, r) for r in rects) or len(over) >= 2 or sum(over) >= 0.4 * max(w.x1 - w.x0, 1)

    placed = set()
    for para in res.paragraphs:
        for ln in para.lines:
            hit = [r for r in rects if r not in placed and min(ln.bbox[3], r[3]) - max(ln.bbox[1], r[1]) > 0.4 * (r[3] - r[1])]
            ln.words = [w for w in ln.words if not inside(w)]
            for r in hit:
                ln.words.append(Word(combs[r], r[0], r[1], r[2], r[3], 95.0))
                placed.add(r)
            if hit:
                ln.words.sort(key=lambda w: w.x0)
    left = [r for r in rects if r not in placed]
    if left:                                   # строку номера Tesseract мог не найти совсем
        y0, y1 = min(r[1] for r in left), max(r[3] for r in left)
        words = [Word(combs[r], r[0], r[1], r[2], r[3], 95.0) for r in sorted(left)]
        res.paragraphs.append(Paragraph([Line((min(r[0] for r in left), y0, max(r[2] for r in left), y1),
                                              y1, (y1 - y0) * 0.5, words)]))


MARK_MIN_CONF = 70        # под печатью и подписью неуверенное слово — обрывки штрихов («Pte,», «axe»)


def drop_unsure_in_marks(res: OcrResult, marks: list[Rect]) -> None:
    if not marks:
        return
    for para in res.paragraphs:
        for ln in para.lines:
            ln.words = [w for w in ln.words if w.conf >= MARK_MIN_CONF or not _inside(
                (w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2, marks)]


def apply_hocr(doc: pymupdf.Document, index: int, prep: PreparedScan, hocr: bytes | str,
               combs: dict | None = None) -> int:
    """Заменяет страницу index на собранную из hOCR. Возвращает число слов;
    0 — страница оказалась иллюстрацией и оставлена как есть (картинкой)."""
    res = parse_hocr(hocr)
    try:
        put_combs(res, combs or {})
    except Exception:
        pass
    drop_unsure_in_marks(res, prep.marks)
    if looks_like_picture(res, prep.img.width, prep.img.height):
        return 0
    try:
        add_cell_dashes(res, prep)
    except Exception:
        pass
    tmp = pymupdf.open()
    n = build_text_page(tmp, res, prep.img, prep.dpi, prep.h_rules, prep.v_rules, InkAnalyzer(prep.clean, prep.img),
                        free_rules=prep.free_rules, marks=prep.marks)
    doc.insert_pdf(tmp, start_at=index)
    doc.delete_page(index + 1)
    return n


def ocr_replace_page(doc: pymupdf.Document, index: int, lang: str, dpi: int | None = None) -> int:
    """Распознаёт страницу локальным Tesseract и заменяет её. Возвращает число слов."""
    prep = prepare_scan(doc[index], dpi)
    hocr = _run_tesseract(scan_png(prep), ["-l", lang, "--dpi", str(int(prep.dpi)), "--psm", "3", "hocr"])
    try:
        combs = read_combs(prep, lang) if prep.combs else {}
        if combs:
            for dash in comb_dashes(prep):
                combs[dash] = "-"
    except Exception:
        combs = {}
    return apply_hocr(doc, index, prep, hocr, combs)


# --------------------------------------------------------------------------- номера страниц

PAGE_ZONE = 0.09            # доля высоты листа у верхнего и нижнего края, где ищется номер
PAGE_MARK_MAX = 30          # символов в надписи с номером
PAGE_MARK_GAP = 6.5         # pt от текста листа (furniture.MIN_BAND_GAP — 6)
_PAGE_WORDS = re.compile(r"(?i)\b(?:стр|страница|лист|листов|из|page|of|с)\b\.?")
_NUMBERS = re.compile(r"\d+")


@dataclass
class PageMark:
    """Надпись с номером страницы у края распознанного листа."""
    page: int
    rect: tuple[float, float, float, float]
    text: str
    numbers: list[int]
    digits: str
    size: float
    top: bool


def page_mark_candidates(page: pymupdf.Page, index: int) -> list[PageMark]:
    """Короткие строки из цифр и знаков («- 4/9 -», «19», «Стр. 3») у верхнего и нижнего края."""
    H = page.rect.height
    out = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", ()):
            spans = [sp for sp in ln["spans"] if sp["text"].strip()]
            if not spans:
                continue
            text = " ".join(" ".join(sp["text"].split()) for sp in spans).strip()
            x0, y0, x1, y1 = ln["bbox"]
            top, bottom = y1 <= PAGE_ZONE * H, y0 >= (1 - PAGE_ZONE) * H
            if not (top or bottom) or len(text) > PAGE_MARK_MAX:
                continue
            rest = _PAGE_WORDS.sub("", text)
            if not any(c.isdigit() for c in rest) or any(c.isalpha() for c in rest):
                continue
            nums = [int(v) for v in _NUMBERS.findall(rest)]
            out.append(PageMark(index, (x0, y0, x1, y1), text, nums, "".join(_NUMBERS.findall(rest)),
                                max(sp["size"] for sp in spans), top))
    return out


def plan_page_marks(cands: dict[int, list[PageMark]], width: dict[int, float],
                    height: dict[int, float]) -> dict | None:
    """Общий вид номера страницы по большинству листов: шаблон «- #/# -», сдвиг нумерации, всего
    листов, место. Номер, прочитанный с ошибкой («19» вместо «1/9»), узнаётся по цифрам."""
    from collections import Counter
    pages = sorted(cands)
    if len(pages) < 3:
        return None
    totals = Counter(m.numbers[1] for ms in cands.values() for m in ms if len(m.numbers) == 2)
    total = totals.most_common(1)[0][0] if totals and totals.most_common(1)[0][1] >= 2 else None
    offsets = Counter(m.numbers[0] - (i + 1) for i, ms in cands.items() for m in ms if m.numbers)
    if not offsets:
        return None
    shift = offsets.most_common(1)[0][0]

    def matches(m: PageMark, i: int) -> bool:
        n = i + 1 + shift
        if m.numbers and m.numbers[0] == n and (total is None or len(m.numbers) == 1 or m.numbers[1] == total):
            return True
        return total is not None and m.digits == f"{n}{total}"

    found = {i: next((m for m in cands[i] if matches(m, i)), None) for i in pages}
    hits = [m for m in found.values() if m is not None]
    if len(hits) < max(3, 0.6 * len(pages)):
        return None
    want = 2 if total is not None else 1
    forms = Counter(_NUMBERS.sub("#", m.text) for m in hits if len(m.numbers) == want)
    if not forms:
        return None
    form = forms.most_common(1)[0][0]
    top = Counter(m.top for m in hits).most_common(1)[0][0]
    hits = [m for m in hits if m.top == top]
    right = [width[m.page] - m.rect[2] for m in hits]
    left = [m.rect[0] for m in hits]
    center = [(m.rect[0] + m.rect[2]) / 2 - width[m.page] / 2 for m in hits]
    spread = {k: max(v) - min(v) for k, v in (("right", right), ("left", left), ("center", center))}
    align = min(spread, key=spread.get)
    edge = (statistics.median(m.rect[3] for m in hits) if top
            else statistics.median(height[m.page] - m.rect[3] for m in hits))
    return {"found": found, "form": form, "shift": shift, "total": total, "top": top, "align": align,
            "pos": statistics.median({"right": right, "left": left, "center": center}[align]),
            "edge": edge, "size": statistics.median(m.size for m in hits)}


def normalize_page_numbers(doc: pymupdf.Document, pages: list[int]) -> int:
    """Номера страниц распознанных листов («- 1/9 -» вверху справа) — одной надписью в одном месте.

    На сканах они стоят с разбросом в несколько пунктов и читаются с ошибками («- 19», «-20-»),
    поэтому повтор на соседних листах не находился и номер оставался текстом страницы. После
    выравнивания его находит общий поиск колонтитулов, а номер становится полем PAGE.
    Возвращает число листов с номером."""
    cands = {i: page_mark_candidates(doc[i], i) for i in pages}
    plan = plan_page_marks(cands, {i: doc[i].rect.width for i in pages}, {i: doc[i].rect.height for i in pages})
    if plan is None:
        return 0
    size = plan["size"]
    # от текста листа — не ближе 6,5 pt, иначе это не колонтитул, а первая строка (furniture)
    edge = plan["edge"]
    for i in pages:
        mark = plan["found"][i]
        H = doc[i].rect.height
        words = [w for w in doc[i].get_text("words")
                 if mark is None or not pymupdf.Rect(w[:4]).intersects(pymupdf.Rect(mark.rect))]
        if plan["top"]:
            below = [w[1] for w in words if w[1] >= (mark.rect[1] if mark else 0) and w[1] < 0.2 * H]
            if below:
                edge = min(edge, min(below) - PAGE_MARK_GAP)
        else:
            above = [H - w[3] for w in words if w[3] <= (mark.rect[3] if mark else H) and w[3] > 0.8 * H]
            if above:
                edge = min(edge, min(above) - PAGE_MARK_GAP)
    if edge < size + 2:
        return 0
    plan["edge"] = edge
    done = 0
    for i in pages:
        page = doc[i]
        font = pymupdf.Font(fontfile=ocr_font_file())     # один Font на несколько страниц не встраивается
        mark = plan["found"][i]
        if mark is None and any(m.top == plan["top"] for m in cands[i]):
            continue                      # у края что-то похожее, но не номер — не трогаем
        if mark is not None:
            page.add_redact_annot(pymupdf.Rect(mark.rect) + (-0.5, -0.5, 0.5, 0.5))
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                  graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
        n = i + 1 + plan["shift"]
        numbers = iter([n, plan["total"]])
        text = re.sub("#", lambda _: str(next(numbers)), plan["form"])
        W, H = page.rect.width, page.rect.height
        w = font.text_length(text, fontsize=size)
        x = {"right": W - plan["pos"] - w, "left": plan["pos"], "center": W / 2 + plan["pos"] - w / 2}[plan["align"]]
        bottom = plan["edge"] if plan["top"] else H - plan["edge"]
        tw = pymupdf.TextWriter(page.rect)
        tw.append((x, bottom - 0.22 * size), text, font=font, fontsize=size)
        tw.write_text(page)
        done += 1
    return done
