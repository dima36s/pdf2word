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
from .regionfix import regions
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
    # Word рисует линию таблицы кусками по ячейкам, на стыках — зазоры до 0,7 pt. С допуском
    # по умолчанию (0,5) границы «не сходятся» и весь лист спецификации становился одной ячейкой.
    "connected_border_tolerance": 1.5,
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
    headers: int = 0                # колонтитулов из шапки/штампа листа
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


class Conversion:
    """Конвертация по шагам. На сервере всё делает convert_pdf(); в браузерной версии шаги
    вызываются из JavaScript, потому что распознавание сканов там выполняет Tesseract.js:

        conv = Conversion(src, opts, external_ocr=True)
        for i in conv.ocr_todo:
            png = conv.prepare_ocr(i)          # картинка страницы → JS → hOCR
            conv.apply_ocr(i, hocr)
        stats = conv.finish(dst)
    """

    def __init__(self, src: Path, opts: ConvertOptions, progress: ProgressFn = _noop,
                 external_ocr: bool = False):
        install_patches()
        self.t0 = time.monotonic()
        self.src, self.opts, self.progress = Path(src), opts, progress
        self.stats = ConvertStats()
        self._prepared: dict[int, ocr.PreparedScan] = {}
        progress(1, "Открываю документ")
        self.doc = open_pdf(src, opts.password)
        self.work = self.doc
        try:
            self.stats.pages_total = self.doc.page_count
            self.indices = parse_page_range(opts.pages, self.doc.page_count)
            self.targets = self.indices if self.indices is not None else list(range(self.doc.page_count))
            self.stats.pages_converted = len(self.targets)
            progress(3, "Анализирую страницы")
            self.kinds = {i: classify_page(self.doc[i]) for i in self.targets}
        except BaseException:
            self.close()
            raise
        scans = [i for i in self.targets if self.kinds[i] in ("scan", "ocred", "vector")]
        vector = [i for i in scans if self.kinds[i] == "vector"]
        if vector:
            pages = ", ".join(str(i + 1) for i in vector[:10]) + ("…" if len(vector) > 10 else "")
            self.stats.warnings.append(f"Текст нарисован линиями (шрифт САПР), страницы распознаны как сканы: {pages}.")
        self.ocr_todo: list[int] = []
        if scans and opts.ocr != "off":
            if external_ocr or ocr.tesseract_available():
                self.ocr_todo = scans
            else:
                self.stats.warnings.append("Распознавание текста (OCR) на этом сервере отключено — "
                                           "отсканированные страницы вставлены как изображения.")
        elif scans:
            self.stats.warnings.append(f"Найдено отсканированных страниц: {len(scans)}. "
                                       "Они вставлены как изображения (распознавание выключено).")

    # ------------------------------------------------------------------ OCR
    def _ensure_work(self):
        if self.work is self.doc:
            self.work = editable_copy(self.doc)
        return self.work

    def run_ocr(self) -> None:
        """Распознавание всех сканов локальным Tesseract (серверная версия)."""
        if not self.ocr_todo:
            return
        lang = ocr.resolve_lang(self.opts.ocr_lang)
        work = self._ensure_work()
        for k, i in enumerate(self.ocr_todo):
            self.progress(5 + 40 * k / len(self.ocr_todo),
                          f"Распознаю текст: страница {i + 1} из {self.doc.page_count}")
            try:
                if ocr.ocr_replace_page(work, i, lang):   # 0 — иллюстрация, осталась картинкой
                    self.stats.ocr_pages += 1
            except ocr.OcrError as exc:
                log.warning("OCR failed on page %s: %s", i + 1, exc)
                self.ocr_failed(i)

    def prepare_ocr(self, i: int) -> bytes:
        """Браузер: подготовить страницу (наклон, линии таблиц) и вернуть PNG для Tesseract.js."""
        prep = ocr.prepare_scan(self.doc[i], detect_orientation=False)
        self._prepared[i] = prep
        return ocr.scan_png(prep)

    def apply_ocr(self, i: int, hocr: bytes | str) -> None:
        """Браузер: собрать страницу из hOCR, полученного от Tesseract.js."""
        prep = self._prepared.pop(i)
        try:
            if ocr.apply_hocr(self._ensure_work(), i, prep, hocr):
                self.stats.ocr_pages += 1
        except Exception as exc:
            log.warning("OCR apply failed on page %s: %s", i + 1, exc)
            self.ocr_failed(i)

    def ocr_failed(self, i: int) -> None:
        self._prepared.pop(i, None)
        self.stats.warnings.append(f"Страница {i + 1}: не удалось распознать текст, вставлена картинкой.")

    # ------------------------------------------------------------------ основной разбор
    def finish(self, dst: Path) -> ConvertStats:
        from pdf2docx import Converter  # импорт здесь: тяжёлый
        from pdf2docx.page.Pages import Pages

        stats, progress, doc, opts = self.stats, self.progress, self.doc, self.opts
        warnings_before = list(stats.warnings)

        # фоновые плашки
        panels = {}
        for i in self.targets:
            if self.kinds[i] == "text":
                try:
                    if found := find_background_panels(doc[i]):
                        panels[i] = found
                except Exception:
                    log.warning("panel detection failed on page %s", i + 1)
        if panels:
            work = self._ensure_work()
            for i, rects in panels.items():
                try:
                    remove_background_panels(work[i], rects)
                except Exception:
                    log.warning("panel removal failed on page %s", i + 1)

        # pdf2docx
        base = 45 if self.ocr_todo else 5
        progress(base, "Анализирую структуру документа")
        cv = Converter.__new__(Converter)       # используем уже открытый (и расшифрованный) документ
        cv.filename_pdf = str(self.src)
        cv.password = ""
        cv._fitz_doc = self.work
        cv._pages = Pages()
        cfg = {**cv.default_settings, **PARSE_SETTINGS}

        collector = _PageErrorCollector()
        logging.getLogger().addHandler(collector)
        made: list[int] = []
        plan = None
        try:
            plan = self._plan_furniture(cfg, base)
            cv.load_pages(pages=self.indices)
            with regions(self.work, plan.body_specs if plan else {}):
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
                    made.append(page.id)
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
        try:
            from .docxfix import improve_structure
            improve_structure(docx)   # заголовки → стили Word, содержание с отточием
        except Exception:
            log.exception("structure post-processing failed")
        if plan:
            try:
                stats.headers = self._attach_furniture(docx, plan, made)
            except Exception:
                # шапка и штамп уже убраны из текста страниц — без колонтитулов они бы пропали,
                # поэтому документ собирается заново по-старому
                log.exception("header/footer attach failed, converting again without them")
                self._furniture_off = True
                stats.warnings[:] = warnings_before
                stats.headers = 0
                return self.finish(dst)
        try:
            from .docxfix import fit_pages
            fit_pages(docx)       # лист PDF — одна страница Word (после колонтитулов: они занимают место)
        except Exception:
            log.exception("page fitting failed")
        dst = Path(dst)
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

        stats.seconds = round(time.monotonic() - self.t0, 2)
        progress(100, "Готово")
        return stats

    # ------------------------------------------------------------------ колонтитулы
    def _plan_furniture(self, cfg: dict, base: float) -> "_FurniturePlan | None":
        """Шапка, штампы и рамка листа → отдельные проходы pdf2docx для колонтитулов (app/furniture.py)."""
        from . import furniture as fu
        from . import geometry as geo
        from . import headers
        from .regionfix import Spec

        if getattr(self, "_furniture_off", False):
            return None
        eligible = [i for i in self.targets if self.kinds[i] == "text" and not self.work[i].rotation]
        if not eligible:
            return None
        try:
            elems = [fu.elements_from_fitz(self.work[i]) for i in eligible]
            found = fu.detect(elems)
        except Exception:
            log.warning("page furniture detection failed", exc_info=True)
            return None
        plan = _FurniturePlan()
        for i, pe, f in zip(eligible, elems, found):
            if not f.empty or f.frame is not None:
                plan.pages[i] = (pe, f)
        if not plan.pages:
            return None
        self.progress(base, "Переношу шапку и штамп в колонтитулы")

        def inside(e, boxes):
            return any(geo.contains(b, e.bbox, 1.5) for b in boxes)

        for i, (pe, f) in plan.pages.items():
            # общие поля тела и колонтитулов: по всему содержимому, кроме бокового штампа (он
            # «плавающий») — тогда ни у текста, ни у шапки не будет отрицательных отступов
            content = [e for e in pe.elems if not inside(e, list(f.sides) + list(f.drop))]
            if content:
                x0, x1 = min(e.bbox[0] for e in content), max(e.bbox[2] for e in content)
                plan.lr[i] = (min(72.0, round(max(x0, 0.0), 1)), min(72.0, round(max(pe.width - x1 - 1.0, 0.0), 1)))
            else:
                plan.lr[i] = None

        def boxes_of(f, kind):
            if kind == "header":
                return [f.header] if f.header else []
            if kind == "footer":
                return [f.footer] if f.footer else []
            return list(f.sides)

        for kind in ("header", "footer", "side"):
            specs, alias, seen = {}, {}, {}
            for i, (pe, f) in plan.pages.items():
                boxes = boxes_of(f, kind)
                if not boxes:
                    continue
                inner = sorted((e.kind, e.text, tuple(round(v) for v in e.bbox)) for e in pe.elems if inside(e, boxes))
                sig = (round(pe.width), round(pe.height), tuple(boxes), tuple(inner), plan.lr[i])
                if sig in seen:          # одинаковая шапка на многих листах — разбираем один раз
                    alias[i] = seen[sig]
                    continue
                seen[sig] = i
                # боковой штамп стоит левее рамки и вставляется «плавающим» — общие поля ему не нужны
                # (с ними pdf2docx отбросил бы его как содержимое за полем страницы)
                specs[i] = Spec(keep=boxes, frame=f.frame, lr_margin=None if kind == "side" else plan.lr[i],
                                frame_mode="touched" if kind == "footer" else "drop")
            try:
                res = headers.band_pass(self.work, cfg, specs)
            except Exception:
                log.warning("%s pass failed", kind, exc_info=True)
                res = {}
            for i, j in alias.items():
                if j in res:
                    res[i] = res[j]
            plan.bands[kind] = res

        for i, (pe, f) in plan.pages.items():
            ok_h = bool(f.header) and i in plan.bands["header"]
            ok_f = bool(f.footer) and i in plan.bands["footer"]
            ok_s = bool(f.sides) and i in plan.bands["side"]
            exclude = ([f.header] if ok_h else []) + ([f.body_footer or f.footer] + list(f.drop) if ok_f else []) \
                + (list(f.sides) if ok_s else [])
            plan.body_specs[i] = Spec(exclude=exclude, frame=f.frame,
                                      top_limit=f.header[3] if ok_h else None,
                                      bottom_limit=f.footer[1] if ok_f else None,
                                      lr_margin=plan.lr[i])
            dist = {}
            if ok_h:
                tops = [e.bbox[1] for e in pe.elems if inside(e, [f.header])]
                dist["header"] = (min(tops) if tops else f.header[1]) - 0.5
            elif f.frame is not None or ok_s:
                dist["header"] = 0.0
            if ok_f:
                bottoms = [e.bbox[3] for e in pe.elems if inside(e, [f.footer])]
                dist["footer"] = pe.height - (max(bottoms) if bottoms else f.footer[3]) - 0.5
            plan.distances[i] = dist
        return plan

    def _attach_furniture(self, docx, plan: "_FurniturePlan", made: list[int]) -> int:
        from docx.enum.section import WD_SECTION

        from . import headers

        page_secs = [s for s in docx.sections if s.start_type == WD_SECTION.NEW_PAGE]
        if len(page_secs) != len(made):
            raise RuntimeError(f"sections/pages mismatch ({len(page_secs)}/{len(made)})")
        contents: dict[int, dict] = {}
        for pid in made:
            if pid not in plan.pages:
                continue
            _, f = plan.pages[pid]
            c = {}
            if (r := plan.bands["header"].get(pid)) is not None:
                c["header"] = headers.prepare(r[0])
            side = plan.bands["side"].get(pid)
            if side is not None or f.frame is not None:
                band = c.get("header") or headers.Band([])
                if side is not None:
                    sb = headers.prepare(side[0])
                    tables = [el for el in sb.elements if el.tag == headers.w("tbl")]
                    for el, bbox in zip(tables, side[1]):
                        band.elements.append(headers.textbox_paragraph([el], bbox))
                    band.images.update(sb.images)
                if f.frame is not None:
                    band.elements.append(headers.frame_paragraph(f.frame))
                c["header"] = band
            if (r := plan.bands["footer"].get(pid)) is not None:
                c["footer"] = headers.prepare(r[0])
            contents[pid] = c
        first = headers.page_fields(contents)
        n = headers.attach(list(zip(made, page_secs)), contents, plan.distances)
        if first is not None and first != 1 and made and made[0] in contents:
            headers.set_start_page(page_secs[0], first)
        return n

    def close(self) -> None:
        if self.work is not self.doc:
            self.work.close()
        self.doc.close()


@dataclass
class _FurniturePlan:
    pages: dict = field(default_factory=dict)        # номер страницы → (PageElems, Furniture)
    lr: dict = field(default_factory=dict)           # одинаковые поля слева/справа
    bands: dict = field(default_factory=lambda: {"header": {}, "footer": {}, "side": {}})
    body_specs: dict = field(default_factory=dict)
    distances: dict = field(default_factory=dict)


def convert_pdf(src: Path, dst: Path, opts: ConvertOptions, progress: ProgressFn = _noop) -> ConvertStats:
    """Полная конвертация (серверная версия)."""
    conv = Conversion(src, opts, progress)
    try:
        conv.run_ocr()
        return conv.finish(dst)
    finally:
        conv.close()


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
