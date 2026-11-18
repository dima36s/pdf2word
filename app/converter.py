"""Конвейер конвертации PDF → DOCX.

    открыть (пароль) → определить сканы → OCR сканов → pdf2docx (с исправлениями из textfix)
    → постобработка DOCX (шрифты, метаданные)

Функция convert_pdf синхронная и тяжёлая — её запускают в отдельном процессе (см. tasks.py).
"""
from __future__ import annotations

import copy
import logging
import os
import re
import tempfile
import time
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable

from . import ocr
from .pdftools import (ConversionError, classify_page, editable_copy, find_background_panels, open_pdf,
                       parse_page_range, remove_background_panels)
from .textfix import install_patches, parse_page_safely

log = logging.getLogger(__name__)

ProgressFn = Callable[[float, str], None]

# Свободные шрифты в PDF (LibreOffice, Linux) → метрически совместимые шрифты Word
FONT_MAP = {
    "Liberation Serif": "Times New Roman",
    "LiberationSerif": "Times New Roman",
    "Tinos": "Times New Roman",
    "Times": "Times New Roman",
    "Times-Roman": "Times New Roman",
    "TimesNewRomanPSMT": "Times New Roman",
    "Liberation Sans": "Arial",
    "LiberationSans": "Arial",
    "Arimo": "Arial",
    "Helvetica": "Arial",
    "ArialMT": "Arial",
    "Liberation Mono": "Courier New",
    "LiberationMono": "Courier New",
    "Cousine": "Courier New",
    "Courier": "Courier New",
    "CourierNewPSMT": "Courier New",
    "Carlito": "Calibri",
    "Caladea": "Cambria",
    "OpenSymbol": "Arial",
    "GlyphLessFont": "Times New Roman",
}

# Параметры pdf2docx (поверх значений по умолчанию)
PARSE_SETTINGS = {
    "ignore_page_error": True,
    "multi_processing": False,
    "debug": False,
}


@dataclass
class ConvertOptions:
    password: str | None = None
    pages: str | None = None        # «1-3, 5»; пусто — все
    ocr: str = "auto"               # auto — распознавать сканы; off — оставить сканы картинками
    ocr_lang: str | None = None     # «rus+eng»
    title: str | None = None        # заголовок в свойствах DOCX


@dataclass
class ConvertStats:
    pages_total: int = 0
    pages_converted: int = 0
    ocr_pages: int = 0
    tables: int = 0
    images: int = 0
    seconds: float = 0.0
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


class _PageErrorCollector(logging.Handler):
    """Собирает ошибки отдельных страниц из логов pdf2docx (он их пропускает и продолжает)."""

    def __init__(self) -> None:
        super().__init__(level=logging.ERROR)
        self.pages: list[int] = []

    def emit(self, record: logging.LogRecord) -> None:
        m = re.search(r"Ignore page (\d+)", record.getMessage())
        if m:
            self.pages.append(int(m.group(1)))


def _noop(_p: float, _s: str) -> None:
    pass


def convert_pdf(src: Path, dst: Path, opts: ConvertOptions, progress: ProgressFn = _noop) -> ConvertStats:
    t0 = time.monotonic()
    install_patches()
    from pdf2docx import Converter  # импорт здесь: тяжёлый, нужен только в рабочем процессе
    from pdf2docx.page.Pages import Pages

    stats = ConvertStats()
    progress(1, "Открываю документ")
    doc = open_pdf(src, opts.password)
    try:
        stats.pages_total = doc.page_count
        indices = parse_page_range(opts.pages, doc.page_count)
        targets = indices if indices is not None else list(range(doc.page_count))
        stats.pages_converted = len(targets)

        # ---------------------------------------------------------------- сканы и OCR
        progress(3, "Анализирую страницы")
        kinds = {i: classify_page(doc[i]) for i in targets}
        scans = [i for i in targets if kinds[i] in ("scan", "ocred")]
        work = doc
        if scans and opts.ocr != "off":
            if not ocr.tesseract_available():
                stats.warnings.append("Распознавание текста (OCR) на этом сервере отключено — "
                                      "отсканированные страницы вставлены как изображения.")
            else:
                lang = ocr.resolve_lang(opts.ocr_lang)
                work = editable_copy(doc)
                for k, i in enumerate(scans):
                    progress(5 + 40 * k / len(scans), f"Распознаю текст: страница {i + 1} из {doc.page_count}")
                    try:
                        ocr.ocr_replace_page(work, i, lang)
                        stats.ocr_pages += 1
                    except ocr.OcrError as exc:
                        log.warning("OCR failed on page %s: %s", i + 1, exc)
                        stats.warnings.append(f"Страница {i + 1}: не удалось распознать текст, вставлена картинкой.")
        elif scans:
            stats.warnings.append(f"Найдено отсканированных страниц: {len(scans)}. "
                                  "Они вставлены как изображения (распознавание выключено).")

        # ---------------------------------------------------------------- фоновые плашки
        panels = {}
        for i in targets:
            if kinds[i] == "text":
                try:
                    if found := find_background_panels(doc[i]):
                        panels[i] = found
                except Exception:
                    log.warning("panel detection failed on page %s", i + 1)
        if panels:
            if work is doc:
                work = editable_copy(doc)
            for i, rects in panels.items():
                try:
                    remove_background_panels(work[i], rects)
                except Exception:
                    log.warning("panel removal failed on page %s", i + 1)

        # ---------------------------------------------------------------- pdf2docx
        base = 45 if stats.ocr_pages else 5
        progress(base, "Анализирую структуру документа")
        cv = Converter.__new__(Converter)       # используем уже открытый (и расшифрованный) документ
        cv.filename_pdf = str(src)
        cv.password = ""
        cv._fitz_doc = work
        cv._pages = Pages()
        cfg = {**cv.default_settings, **PARSE_SETTINGS}

        collector = _PageErrorCollector()
        logging.getLogger().addHandler(collector)
        try:
            cv.load_pages(pages=indices)
            cv.parse_document(**cfg)

            to_parse = [p for p in cv.pages if not p.skip_parsing]
            span = (90 - base) * 0.75
            for k, page in enumerate(to_parse):
                progress(base + span * k / max(len(to_parse), 1), f"Разбираю страницу {page.id + 1}")
                try:
                    parse_page_safely(page, cfg)
                except Exception as exc:  # одна страница не должна валить весь документ
                    log.warning("page %s parse error: %s", page.id + 1, exc)
                    collector.pages.append(page.id + 1)

            from docx import Document
            docx = Document()
            ready = [p for p in cv.pages if p.finalized]
            if not ready:
                raise ConversionError("Не удалось разобрать ни одной страницы документа.")
            start = base + span
            for k, page in enumerate(ready):
                progress(start + (95 - start) * k / len(ready), f"Собираю Word: страница {page.id + 1}")
                try:
                    page.make_docx(docx)
                except Exception as exc:
                    log.warning("page %s make_docx error: %s", page.id + 1, exc)
                    collector.pages.append(page.id + 1)
        finally:
            logging.getLogger().removeHandler(collector)

        if collector.pages:
            bad = ", ".join(str(p) for p in sorted(set(collector.pages)))
            stats.warnings.append(f"Не удалось полностью перенести страницы: {bad}.")

        progress(96, "Сохраняю документ")
        _finalize_docx(docx, title=opts.title)
        try:
            _merge_column_sections(docx)
        except Exception:
            log.exception("column section merge failed")
        dst.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=dst.parent)
        os.close(fd)
        tmp = Path(tmp_name)
        try:
            docx.save(tmp)
            stats.tables, stats.images = _postprocess_docx(tmp)
            os.replace(tmp, dst)
        finally:
            tmp.unlink(missing_ok=True)
    finally:
        if work is not doc:
            work.close()
        doc.close()

    stats.seconds = round(time.monotonic() - t0, 2)
    progress(100, "Готово")
    return stats


def _finalize_docx(docx, title: str | None) -> None:
    props = docx.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = "Сконвертировано из PDF"
    if title:
        props.title = title


def _merge_column_sections(docx) -> int:
    """Колонки: разделы «со следующей колонки» → один раздел с явным разрывом колонки.

    pdf2docx кладёт каждую колонку в отдельный раздел с типом начала nextColumn. MS Word это
    понимает, а LibreOffice, R7-Офис и другие редакторы — нет: текст разъезжается по страницам.
    Разрыв колонки внутри одного раздела одинаково работает везде.
    """
    from docx.oxml.ns import qn
    from lxml import etree

    body = docx.element.body
    W_P, W_PPR, W_SECT, W_TYPE, W_COLS, W_R = (qn(t) for t in ("w:p", "w:pPr", "w:sectPr", "w:type", "w:cols", "w:r"))

    def sect_type(sect) -> str:
        t = sect.find(W_TYPE)
        return t.get(qn("w:val")) if t is not None else "nextPage"

    merged = 0
    while True:
        holders = [p for p in body.iterchildren(W_P) if p.find(f"{W_PPR}/{W_SECT}") is not None]
        sects = [h.find(f"{W_PPR}/{W_SECT}") for h in holders]
        final = body.find(W_SECT)
        if final is not None:
            sects.append(final)
        for i, holder in enumerate(holders):
            a, b = sects[i], sects[i + 1] if i + 1 < len(sects) else None
            if b is None or sect_type(b) != "nextColumn":
                continue
            # свойства объединённого раздела: тип начала и колонки — от первой колонки
            for tag in (W_TYPE, W_COLS):
                old_b, old_a = b.find(tag), a.find(tag)
                if old_b is not None:
                    b.remove(old_b)
                if old_a is not None:
                    b.append(copy.deepcopy(old_a))
            # порядок дочерних элементов sectPr важен для Word: type идёт перед pgSz
            t = b.find(W_TYPE)
            if t is not None:
                b.remove(t)
                b.insert(0, t)
            a.getparent().remove(a)
            nxt = holder.getnext()
            if all(child.tag == W_PPR for child in holder):  # служебный пустой абзац python-docx
                body.remove(holder)
            brk_run = etree.SubElement(etree.Element(W_P), W_R)
            etree.SubElement(brk_run, qn("w:br")).set(qn("w:type"), "column")
            if nxt is not None and nxt.tag == W_P:
                ppr = nxt.find(W_PPR)
                nxt.insert(0 if ppr is None else list(nxt).index(ppr) + 1, brk_run)
            else:  # колонка начинается с таблицы — разрыв отдельным абзацем перед ней
                p = etree.Element(W_P)
                p.append(brk_run)
                if nxt is not None:
                    nxt.addprevious(p)
                else:
                    body.append(p)
            merged += 1
            break
        else:
            return merged


_TR_EXACT_RE = re.compile(r'(<w:trHeight\b[^>]*w:hRule=")exact(")')
_FONT_ATTR_RE = re.compile(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")([^"]+)(")')


def _map_font(m: re.Match) -> str:
    name = m.group(2)
    base = re.sub(r"^[A-Z]{6}\+", "", name)  # префикс подмножества шрифта: ABCDEF+Arial
    mapped = FONT_MAP.get(base) or FONT_MAP.get(base.replace(" ", "")) or base
    return f"{m.group(1)}{mapped}{m.group(3)}"


def _postprocess_docx(path: Path) -> tuple[int, int]:
    """Замена свободных шрифтов на шрифты Word; подсчёт таблиц и картинок."""
    tables = images = 0
    out = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if name.startswith("word/") and name.endswith(".xml") and (
                    name == "word/document.xml" or name.startswith(("word/header", "word/footer"))
                    or name in ("word/styles.xml", "word/fontTable.xml")):
                text = data.decode("utf-8")
                text = _FONT_ATTR_RE.sub(_map_font, text)
                if name == "word/document.xml":
                    # «точная» высота строк таблиц обрезает текст, если шрифт в Word чуть шире
                    text = _TR_EXACT_RE.sub(r'\1atLeast\2', text)
                    tables = text.count("<w:tbl>")
                    images = text.count("<pic:pic")
                data = text.encode("utf-8")
            zout.writestr(item, data)
    os.replace(out, path)
    return tables, images
