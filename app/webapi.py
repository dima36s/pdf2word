"""API конвертера для браузерной версии (Pyodide, см. web/worker.js).

Тот же конвейер, что и на сервере, но по шагам: распознавание сканов в браузере выполняет
Tesseract.js, поэтому Python отдаёт картинку страницы и принимает готовый hOCR.
Файлы живут в памяти браузера (/work в файловой системе Pyodide) и никуда не отправляются.
"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

from .converter import Conversion, ConvertOptions
from .pdftools import inspect_pdf, open_pdf, render_thumbnail

WORK = Path("/work")
_sessions: dict[str, Conversion] = {}


def _report(key: str):
    try:
        from js import pyProgress  # функция из web/worker.js
    except ImportError:            # запуск вне браузера (тесты)
        return lambda pct, stage: None
    return lambda pct, stage: pyProgress(key, pct, stage)


def _bytes(x) -> bytes:
    return x.to_bytes() if hasattr(x, "to_bytes") else bytes(x)


def _read_and_remove(path: Path) -> bytes | None:
    if not path.is_file():
        return None
    data = path.read_bytes()
    path.unlink()
    return data


def warmup() -> None:
    """Импорт тяжёлых модулей заранее, чтобы первая конвертация не ждала."""
    import importlib

    from .textfix import install_patches
    install_patches()
    importlib.import_module("pdf2docx.converter")


def inspect(key: str) -> dict:
    src = WORK / f"{key}.pdf"
    thumb = WORK / f"{key}.png"
    info = inspect_pdf(src, thumb)
    return {"pages": info.pages, "encrypted": info.encrypted, "scanned_pages": info.scanned_pages,
            "thumb": _read_and_remove(thumb)}


def unlock(key: str, password: str) -> dict:
    doc = open_pdf(WORK / f"{key}.pdf", password)
    try:
        thumb = WORK / f"{key}.png"
        render_thumbnail(doc[0], thumb)
        return {"pages": doc.page_count, "thumb": _read_and_remove(thumb)}
    finally:
        doc.close()


def start(key: str, options: dict) -> dict:
    drop(key, keep_source=True)
    opts = ConvertOptions(
        password=options.get("password") or None,
        pages=options.get("pages") or None,
        ocr=options.get("ocr") or "auto",
        ocr_lang=options.get("ocr_lang") or None,
        title=options.get("title") or None,
    )
    conv = Conversion(WORK / f"{key}.pdf", opts, progress=_report(key), external_ocr=True)
    _sessions[key] = conv
    return {"ocr_pages": list(conv.ocr_todo), "pages_total": conv.stats.pages_total}


def prepare(key: str, page: int) -> bytes:
    return _sessions[key].prepare_ocr(int(page))


def apply(key: str, page: int, hocr: str) -> bool:
    _sessions[key].apply_ocr(int(page), hocr)
    return True


def skip(key: str, page: int) -> bool:
    _sessions[key].ocr_failed(int(page))
    return True


def finish(key: str) -> dict:
    conv = _sessions.pop(key)
    out = WORK / f"{key}.docx"
    xlsx = WORK / f"{key}.xlsx"
    try:
        stats = conv.finish(out, xlsx_dst=xlsx)
    finally:
        conv.close()
    return {"docx": _read_and_remove(out), "xlsx": _read_and_remove(xlsx), "stats": stats.to_dict()}


def make_zip(files: list) -> bytes:
    buf = io.BytesIO()
    used: dict[str, int] = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            name = str(name)
            n = used.get(name.lower(), 0)
            used[name.lower()] = n + 1
            if n:
                name = f"{Path(name).stem} ({n + 1}){Path(name).suffix}"
            zf.writestr(name, _bytes(data))
    return buf.getvalue()


def drop(key: str, keep_source: bool = False) -> bool:
    conv = _sessions.pop(key, None)
    if conv is not None:
        conv.close()
    names = [f"{key}.docx", f"{key}.xlsx", f"{key}.png"] + ([] if keep_source else [f"{key}.pdf"])
    for name in names:
        (WORK / name).unlink(missing_ok=True)
    return True
