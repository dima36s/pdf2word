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
    return shutil.which(settings.tesseract_cmd) is not None


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


def detect_table_rules(gray: Image.Image, dpi: float) -> tuple[list[Rect], list[Rect]]:
    """Горизонтальные и вертикальные линии, образующие сетку таблиц (в пикселях)."""
    import numpy as np
    arr = np.asarray(gray)
    dark = arr < 150
    tol = max(2, int(dpi / 75))
    h_mask = _erode(_dilate(dark, tol, 1), tol, 1)       # закрыть разрывы в отсканированных линиях
    v_mask = _erode(_dilate(dark, tol, 0), tol, 0)
    max_thick = max(3, int(dpi * 0.03))
    h = _segments(h_mask, int(dpi * 0.35), max_thick, True)
    v = _segments(v_mask, int(dpi * 0.18), max_thick, False)
    # оставляем только линии, образующие сетку: каждая касается хотя бы двух перпендикулярных
    t = tol * 2
    h_keep = [a for a in h if sum(_touches(a, b, t) for b in v) >= 2]
    v_keep = [b for b in v if sum(_touches(a, b, t) for a in h_keep) >= 2]
    h_keep = [a for a in h_keep if sum(_touches(a, b, t) for b in v_keep) >= 2]
    return h_keep, v_keep


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


def parse_hocr(data: bytes) -> OcrResult:
    root = ET.fromstring(data)
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
    limit = max(2.5 * max(med, 1.0), 1.6 * fs_px)
    uniform = len(gaps) >= 2 and max(gaps) < 2.5 * max(med, 1.0)
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


class InkAnalyzer:
    """Оценивает по пикселям скана жирность и цвет слов."""

    def __init__(self, gray: Image.Image, rgb: Image.Image):
        import numpy as np
        self.np = np
        self.dark = np.asarray(gray) < 140
        self.integral = np.pad(self.dark.astype(np.int32).cumsum(0).cumsum(1), ((1, 0), (1, 0)))
        self.rgb = np.asarray(rgb)
        self.h, self.w = self.dark.shape
        self.ref: float | None = None

    def _box(self, w: Word):
        x0, y0 = max(0, int(w.x0)), max(0, int(w.y0))
        x1, y1 = min(self.w, int(math.ceil(w.x1))), min(self.h, int(math.ceil(w.y1)))
        return x0, y0, x1, y1

    def stroke(self, w: Word, x_size: float) -> float | None:
        """Средняя толщина штриха (длина горизонтальной серии тёмных пикселей) относительно кегля."""
        x0, y0, x1, y1 = self._box(w)
        m = self.dark[y0:y1, x0:x1]
        n = int(m.sum())
        if n < 15 or x_size <= 0:
            return None
        runs = int(self.np.count_nonzero(self.np.diff(self.np.pad(m, ((0, 0), (1, 0))).astype(self.np.int8), axis=1) == 1))
        return n / max(runs, 1) / x_size

    def calibrate(self, items: list[tuple[Word, float]]) -> None:
        vals = [v for w, xs in items if len(w.text) >= 3 and (v := self.stroke(w, xs))]
        self.ref = statistics.median(vals) if len(vals) >= 5 else None

    def ratio(self, w: Word, x_size: float) -> float:
        v = self.stroke(w, x_size)
        return v / self.ref if (v and self.ref) else 1.0

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
                    ink: InkAnalyzer | None = None) -> int:
    """Добавляет в out_doc страницу с распознанным текстом, иллюстрациями и линиями таблиц.
    Возвращает число слов."""
    s = 72.0 / dpi
    w_pt, h_pt = img.width * s, img.height * s
    page = out_doc.new_page(width=w_pt, height=h_pt)
    page_area = img.width * img.height

    # иллюстрации (графики, фото, печати) — вырезаем из скана как есть
    photos = []
    for x0, y0, x1, y1 in res.photos:
        area = (x1 - x0) * (y1 - y0)
        if area < page_area * 0.004 or area > page_area * 0.6:
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
    for x0, y0, x1, y1 in h_rules:
        y = (y0 + y1) / 2 * s
        page.draw_line((x0 * s, y), (x1 * s, y), color=(0, 0, 0), width=max(0.5, (y1 - y0) * s))
    for x0, y0, x1, y1 in v_rules:
        x = (x0 + x1) / 2 * s
        page.draw_line((x, y0 * s), (x, y1 * s), color=(0, 0, 0), width=max(0.5, (x1 - x0) * s))

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
        ink.calibrate([(w, ln.x_size) for p in res.paragraphs for ln in p.lines for w in ln.words])

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

    n_words = 0
    for par, (raw, _) in zip(res.paragraphs, raw_sizes):
        fs = body if body and abs(raw - body) / body <= 0.1 else snap_size(raw)
        for ln in par.lines:
            words = [cw for w in ln.words if (cw := _clean_word(w)) is not None]
            if not words:
                continue
            cx, cy = (ln.bbox[0] + ln.bbox[2]) / 2, (ln.bbox[1] + ln.bbox[3]) / 2
            if _inside(cx, cy, photos):
                continue  # подписи внутри графика остаются частью картинки
            words = _fix_bullet(words, fs / s)

            # стиль слов: жирность по плотности «чернил», цвет по пикселям штрихов
            styles = [(False, None)] * len(words)
            if ink is not None and ink.ref:
                alnum = [any(c.isalnum() for c in w.text) for w in words]
                ratios = [ink.ratio(w, ln.x_size) for w in words]
                core = [r for r, a in zip(ratios, alnum) if a] or [1.0]
                line_bold = statistics.median(core) >= 1.25
                styles = [((line_bold or (r >= 1.4 and len(w.text) >= 3)) and a, ink.color(w))
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
    return n_words


# --------------------------------------------------------------------------- точка входа

def render_page(page: pymupdf.Page, dpi: int) -> tuple[Image.Image, float]:
    w_in, h_in = page.rect.width / 72, page.rect.height / 72
    dpi = min(dpi, int(math.sqrt(MAX_RENDER_PIXELS / max(w_in * h_in, 0.01))))
    pix = page.get_pixmap(dpi=dpi, alpha=False, colorspace=pymupdf.csRGB)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    # фактический dpi по ширине (get_pixmap округляет размеры)
    return img, pix.width / w_in


def ocr_replace_page(doc: pymupdf.Document, index: int, lang: str, dpi: int | None = None) -> int:
    """Заменяет страницу index документа doc на распознанную. Возвращает число слов."""
    from PIL import ImageFilter

    img, real_dpi = render_page(doc[index], dpi or settings.ocr_dpi)
    gray = img.convert("L").filter(ImageFilter.MedianFilter(3))  # убирает «соль и перец» скана
    rotation = detect_rotation(gray)
    if rotation:
        img = img.rotate(-rotation, expand=True)
        gray = gray.rotate(-rotation, expand=True)
    skew = estimate_skew(gray)
    if abs(skew) >= 0.1:
        img = img.rotate(skew, resample=Image.BICUBIC, fillcolor=(255, 255, 255))
        gray = gray.rotate(skew, resample=Image.BICUBIC, fillcolor=255)
    h_rules, v_rules = detect_table_rules(gray, real_dpi)
    clean = erase_rules(gray, h_rules + v_rules)
    hocr = _run_tesseract(_png(clean), ["-l", lang, "--dpi", str(int(real_dpi)), "--psm", "3", "hocr"])
    res = parse_hocr(hocr)
    tmp = pymupdf.open()
    n = build_text_page(tmp, res, img, real_dpi, h_rules, v_rules, InkAnalyzer(clean, img))
    doc.insert_pdf(tmp, start_at=index)
    doc.delete_page(index + 1)
    return n
