"""Работа с PDF: открытие, проверка пароля, анализ страниц, превью, диапазоны страниц."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pymupdf

pymupdf.TOOLS.mupdf_display_errors(False)
pymupdf.TOOLS.mupdf_display_warnings(False)

MIN_TEXT_CHARS = 25        # меньше видимых символов — страница считается «без текста»
SCAN_IMAGE_COVERAGE = 0.5  # картинка закрывает не меньше половины страницы — похоже на скан
VECTOR_TEXT_PATHS = 400    # столько мелких путей без текстового слоя — текст нарисован линиями (САПР)


class ConversionError(Exception):
    """Ошибка, текст которой можно показать пользователю."""


# --------------------------------------------------------------------------- открытие

def open_pdf(path: Path | str, password: str | None = None) -> pymupdf.Document:
    try:
        doc = pymupdf.open(path)
    except Exception as exc:  # повреждённый или не-PDF файл
        raise ConversionError("Не удалось открыть файл: он повреждён или это не PDF.") from exc
    if not doc.is_pdf:
        doc.close()
        raise ConversionError("Файл не является PDF-документом.")
    if doc.needs_pass:
        if not password:
            doc.close()
            raise ConversionError("Документ защищён паролем — укажите пароль.")
        if not doc.authenticate(password):
            doc.close()
            raise ConversionError("Неверный пароль к документу.")
        # Сразу расшифровываем в память: в PyMuPDF 2.x повторное обращение к needs_pass
        # после authenticate() сбрасывает расшифровку, а pdf2docx его проверяет.
        decrypted = editable_copy(doc)
        doc.close()
        doc = decrypted
    if doc.page_count == 0:
        doc.close()
        raise ConversionError("В документе нет страниц.")
    return doc


def editable_copy(doc: pymupdf.Document) -> pymupdf.Document:
    """Расшифрованная копия документа, в которой можно заменять страницы."""
    try:
        data = doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_NONE, garbage=0)
        return pymupdf.open("pdf", data)
    except Exception:
        out = pymupdf.open()
        out.insert_pdf(doc)
        return out


# --------------------------------------------------------------------------- анализ страниц

def _image_coverage(page: pymupdf.Page) -> float:
    area = abs(page.rect) or 1.0
    covered = 0.0
    try:
        infos = page.get_image_info()
    except Exception:
        return 0.0
    for info in infos:
        r = pymupdf.Rect(info["bbox"]) & page.rect
        if not r.is_empty:
            covered += abs(r)
    return min(covered / area, 1.0)


def _vector_text(page: pymupdf.Page) -> bool:
    """Текст нарисован линиями (шрифты SHX из AutoCAD/nanoCAD и т.п.): текстового слоя нет, зато
    тысячи мелких путей. pdf2docx вставил бы такой лист картинкой — его нужно распознавать."""
    try:
        drawings = page.get_cdrawings()
    except Exception:
        return False
    if len(drawings) < VECTOR_TEXT_PATHS:
        return False
    page_area = abs(page.rect) or 1.0
    small = 0
    for d in drawings:
        x0, y0, x1, y1 = d.get("rect", (0, 0, 0, 0))
        if (x1 - x0) * (y1 - y0) < page_area * 2e-4 and max(x1 - x0, y1 - y0) < 40:
            small += 1
    return small >= VECTOR_TEXT_PATHS


def classify_page(page: pymupdf.Page) -> str:
    """'text' — обычная страница с текстом; 'scan' — картинка без текста;
    'ocred' — скан с невидимым распознанным слоем (такой слой pdf2docx игнорирует);
    'vector' — текст нарисован линиями (выгрузка из САПР), распознаётся как скан."""
    visible = hidden = 0
    try:
        spans = page.get_texttrace()
    except Exception:
        spans = []
    for s in spans:
        n = sum(1 for ch in s.get("chars", ()) if ch[0] > 32)
        if s.get("type") == 3 or s.get("opacity", 1) == 0:
            hidden += n
        else:
            visible += n
    if visible >= MIN_TEXT_CHARS:
        return "text"
    if _image_coverage(page) >= SCAN_IMAGE_COVERAGE:
        return "ocred" if hidden >= MIN_TEXT_CHARS else "scan"
    if _vector_text(page):
        return "vector"
    return "text"


@dataclass
class PdfInfo:
    pages: int | None
    encrypted: bool
    scanned_pages: int
    title: str | None


def inspect_pdf(path: Path, thumb_path: Path | None = None, thumb_width: int = 360) -> PdfInfo:
    """Быстрый анализ файла при загрузке: число страниц, пароль, доля сканов, превью первой страницы."""
    try:
        doc = pymupdf.open(path)
    except Exception as exc:
        raise ConversionError("Не удалось открыть файл: он повреждён или это не PDF.") from exc
    with doc:
        if not doc.is_pdf:
            raise ConversionError("Файл не является PDF-документом.")
        if doc.needs_pass:
            return PdfInfo(pages=None, encrypted=True, scanned_pages=0, title=None)
        n = doc.page_count
        if n == 0:
            raise ConversionError("В документе нет страниц.")
        sample = range(min(n, 10))
        scanned = sum(1 for i in sample if classify_page(doc[i]) != "text")
        scanned = round(scanned * n / len(sample))
        if thumb_path is not None:
            render_thumbnail(doc[0], thumb_path, thumb_width)
        title = (doc.metadata or {}).get("title") or None
        return PdfInfo(pages=n, encrypted=False, scanned_pages=scanned, title=title)


def render_thumbnail(page: pymupdf.Page, out: Path, width: int = 360) -> None:
    zoom = width / max(page.rect.width, 1)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    pix.save(out)


# --------------------------------------------------------------------------- диапазоны

def parse_page_range(spec: str | None, total: int) -> list[int] | None:
    """'1-3, 5, 8-' → [0,1,2,4,7,...]. Пустая строка → None (все страницы)."""
    if not spec or not spec.strip():
        return None
    spec = re.sub(r"\s*[-–—]\s*", "-", spec.strip())
    pages: set[int] = set()
    for part in re.split(r"[,;\s]+", spec):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)-(\d*)", part)
        if m and (m.group(1) or m.group(2)):
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else total
        elif part.isdigit():
            a = b = int(part)
        else:
            raise ConversionError(f"Неверный диапазон страниц: «{part}». Пример: 1-3, 5, 8-10")
        if a < 1 or b < a:
            raise ConversionError(f"Неверный диапазон страниц: «{part}».")
        if a > total:
            raise ConversionError(f"В документе {total} стр., страницы {a} нет.")
        pages.update(range(a - 1, min(b, total)))
    if not pages:
        return None
    return sorted(pages)


# --------------------------------------------------------------------------- фоновые плашки

def _light(color) -> bool:
    return color is not None and len(color) >= 3 and min(color[:3]) >= 0.82


def find_background_panels(page: pymupdf.Page, max_fills: int = 400) -> list[pymupdf.Rect]:
    """Светлые изолированные «карточки» с текстом (типично для PDF из браузера).

    pdf2docx превращает такие плашки в цепочки таблиц 1×1 и ломает колонки, поэтому фон
    убираем, а текст оставляем. Заливки ячеек таблиц (сетка соседних прямоугольников) не трогаем.
    """
    try:
        drawings = page.get_cdrawings()
    except Exception:
        return []
    fills = [(d, pymupdf.Rect(d["rect"])) for d in drawings
             if d.get("fill") is not None and d.get("type") in ("f", "fs")]
    if len(fills) > max_fills:
        return []  # чертёж/диаграмма — не трогаем
    pw = page.rect.width
    panels = []
    for d, r in fills:
        if not _light(d["fill"]) or r.width < 0.15 * pw or r.height < 30 or r.width > 0.98 * pw:
            continue
        if d.get("type") == "fs" and d.get("color") is not None and not _light(d.get("color")):
            continue  # заметная рамка — пусть остаётся таблицей
        grow = pymupdf.Rect(r.x0 - 2, r.y0 - 2, r.x1 + 2, r.y1 + 2)
        if any(o is not d and grow.intersects(orr) and not r.contains(orr) for o, orr in fills):
            continue  # соседние заливки — это ячейки таблицы
        words = page.get_text("words", clip=r)
        if len({round(w[3]) for w in words}) >= 2:
            panels.append(r)
    return panels


def remove_background_panels(page: pymupdf.Page, panels: list[pymupdf.Rect] | None = None) -> int:
    panels = find_background_panels(page) if panels is None else panels
    if not panels:
        return 0
    for r in panels:
        page.add_redact_annot(r + (-0.5, -0.5, 0.5, 0.5))
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                          text=pymupdf.PDF_REDACT_TEXT_NONE)
    return len(panels)


# --------------------------------------------------------------------------- выравнивание по ширине

SHRINK_RATIO = 0.97       # пробел уже 97 % своей ширины — Word сжал его, чтобы слово вошло в строку
EXPAND_RATIO = 1.03
MIN_GAPS = 5              # строки с меньшим числом пробелов ни о чём не говорят


def space_ratios(chars) -> list[float]:
    """Отношение промежутка «пробел → следующая буква» к ширине пробела по шрифту."""
    out = []
    for a, b in zip(chars, chars[1:]):
        if a.get("c") == " " and b.get("c", "").strip():
            natural = a["bbox"][2] - a["bbox"][0]
            if natural > 0.5:
                out.append((b["origin"][0] - a["origin"][0]) / natural)
    return out


def squeezes_spaces(doc: pymupdf.Document, pages=None, max_pages: int = 60, enough: int = 2) -> bool | None:
    """Сжимал ли Word пробелы в строках с выравниванием по ширине (Word 2013 и новее так делает).

    Тогда документ Word должен быть в режиме Word 2013+ (compatibilityMode 15): в режиме Word 2010
    пробелы не сжимаются, строки абзацев становятся длиннее, и плотные листы переезжают на
    следующую страницу. True — сжатые строки есть; False — растянутых строк много, сжатых нет
    (документ из старого Word или другой программы); None — по тексту не понять.
    """
    import statistics

    shrunk = expanded = 0
    for i in list(pages if pages is not None else range(doc.page_count))[:max_pages]:
        try:
            blocks = doc[i].get_text("rawdict", flags=0)["blocks"]
        except Exception:
            continue
        for b in blocks:
            for line in b.get("lines", ()):
                if tuple(line.get("dir", (1, 0))) != (1.0, 0.0):
                    continue
                ratios = space_ratios([c for s in line["spans"] for c in s.get("chars", ())])
                if len(ratios) < MIN_GAPS:
                    continue
                med = statistics.median(ratios)
                if 0.6 < med < SHRINK_RATIO:
                    shrunk += 1
                    if shrunk >= enough:
                        return True
                elif med > EXPAND_RATIO:
                    expanded += 1
    if shrunk:
        return True
    return False if expanded >= 20 else None
