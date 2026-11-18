"""HTTP API и раздача интерфейса.

Запуск:  uvicorn app.main:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import asyncio
import contextlib
import logging
import re
import tempfile
import time
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import __version__, ocr
from .config import settings
from .pdftools import ConversionError, inspect_pdf, open_pdf, parse_page_range, render_thumbnail
from .tasks import ACTIVE, Runner, TaskStore

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("pdf2word")

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
CHUNK = 1024 * 1024

store = TaskStore(settings.tasks_dir)
runner: Runner | None = None


# --------------------------------------------------------------------------- жизненный цикл

async def _cleanup_loop() -> None:
    while True:
        try:
            removed = store.cleanup(settings.file_ttl_min * 60, runner.running if runner else set())
            if removed:
                log.info("cleanup: removed %d expired tasks", removed)
        except Exception:
            log.exception("cleanup failed")
        await asyncio.sleep(300)


@contextlib.asynccontextmanager
async def lifespan(_app: FastAPI):
    global runner
    store.recover_after_restart()
    runner = Runner(store, settings.workers, settings.convert_timeout)
    cleaner = asyncio.create_task(_cleanup_loop())
    log.info("PDF→Word %s: data=%s workers=%d ocr=%s", __version__, store.root, settings.workers,
             ocr.tesseract_available())
    try:
        yield
    finally:
        cleaner.cancel()
        await runner.shutdown()


app = FastAPI(title="PDF → Word", version=__version__, lifespan=lifespan,
              docs_url="/api/docs", redoc_url=None, openapi_url="/api/openapi.json")


@app.middleware("http")
async def guard(request: Request, call_next):
    # слишком большой файл отклоняем до чтения тела запроса
    if request.method == "POST" and request.url.path == "/api/tasks":
        length = request.headers.get("content-length")
        if length and length.isdigit() and int(length) > settings.max_file_bytes + CHUNK:
            return JSONResponse({"detail": f"Файл больше {settings.max_file_mb} МБ."}, status_code=413)
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    response.headers.setdefault("X-Frame-Options", "DENY")
    return response


# --------------------------------------------------------------------------- helpers

def _output_name(name: str) -> str:
    stem = Path(name).stem or "document"
    stem = re.sub(r'[\x00-\x1f\\/:*?"<>|]+', "_", stem).strip(" .") or "document"
    return stem[:150] + ".docx"


def _public(meta: dict) -> dict:
    task_id = meta["id"]
    out = {k: meta.get(k) for k in ("id", "name", "size", "pages", "encrypted", "scanned_pages",
                                    "status", "error", "stats", "output_size", "created")}
    out["has_thumb"] = (store.dir(task_id) / "thumb.png").is_file()
    out["output_name"] = _output_name(meta["name"])
    out["expires_at"] = meta["created"] + settings.file_ttl_min * 60
    if meta.get("status") == "processing":
        prog = store.read_json(task_id, "progress.json") or {}
        out["progress"] = prog.get("progress", 0)
        out["stage"] = prog.get("stage", "Запуск…")
    elif meta.get("status") == "queued":
        out["progress"] = 0
        pos = runner.queue_position(task_id) if runner else None
        out["stage"] = f"В очереди ({pos})" if pos and pos > 1 else "В очереди"
    elif meta.get("status") == "done":
        out["progress"] = 100
    return out


def _get(task_id: str) -> dict:
    meta = store.load(task_id) if store.valid_id(task_id) else None
    if meta is None:
        raise HTTPException(404, "Файл не найден или уже удалён с сервера.")
    return meta


# --------------------------------------------------------------------------- API

@app.get("/api/config")
def get_config() -> dict:
    langs = ocr.available_languages()
    default = ocr.resolve_lang(None) if langs else None
    return {
        "version": __version__,
        "max_file_mb": settings.max_file_mb,
        "max_pages": settings.max_pages,
        "file_ttl_min": settings.file_ttl_min,
        "notice": settings.demo_notice or None,
        "ocr": {
            "available": ocr.tesseract_available() and bool(langs),
            "default": default,
            "languages": [{"code": c, "name": ocr.LANG_NAMES.get(c, c)}
                          for c in sorted(langs, key=lambda c: ({"rus": 0, "eng": 1}.get(c, 2), ocr.LANG_NAMES.get(c, c)))],
        },
    }


@app.get("/healthz")
def healthz() -> dict:
    return {"ok": True, "running": len(runner.procs) if runner else 0, "queued": len(runner.queue) if runner else 0}


@app.post("/api/tasks", status_code=201)
async def upload(file: UploadFile = File(...)) -> dict:
    name = Path(file.filename or "document.pdf").name
    if not name.lower().endswith(".pdf") and file.content_type not in ("application/pdf", "application/x-pdf"):
        raise HTTPException(415, "Можно загружать только PDF-файлы.")
    meta = store.create(name, 0)
    task_id = meta["id"]
    path = store.dir(task_id) / "input.pdf"
    size = 0
    try:
        with path.open("wb") as fh:
            head = b""
            while chunk := await file.read(CHUNK):
                if not head:
                    head = chunk[:1024]
                    if b"%PDF-" not in head:
                        raise HTTPException(415, "Файл не похож на PDF.")
                size += len(chunk)
                if size > settings.max_file_bytes:
                    raise HTTPException(413, f"Файл больше {settings.max_file_mb} МБ.")
                fh.write(chunk)
        if size == 0:
            raise HTTPException(400, "Файл пустой.")
        try:
            info = await asyncio.to_thread(inspect_pdf, path, store.dir(task_id) / "thumb.png")
        except ConversionError as exc:
            raise HTTPException(422, str(exc)) from exc
        if info.pages and info.pages > settings.max_pages:
            raise HTTPException(413, f"В документе {info.pages} стр. — больше допустимых {settings.max_pages}.")
    except HTTPException:
        store.delete(task_id)
        raise
    except Exception:
        store.delete(task_id)
        log.exception("upload failed")
        raise HTTPException(500, "Не удалось сохранить файл.")
    finally:
        await file.close()
    meta = store.update(task_id, size=size, pages=info.pages, encrypted=info.encrypted,
                        scanned_pages=info.scanned_pages, status="uploaded")
    return _public(meta)


@app.get("/api/tasks/{task_id}")
def status(task_id: str) -> dict:
    return _public(_get(task_id))


@app.get("/api/tasks/{task_id}/thumb")
def thumb(task_id: str):
    _get(task_id)
    path = store.dir(task_id) / "thumb.png"
    if not path.is_file():
        raise HTTPException(404, "Нет превью.")
    return FileResponse(path, media_type="image/png", headers={"Cache-Control": "private, max-age=3600"})


class UnlockRequest(BaseModel):
    password: str = Field(..., max_length=256)


def _unlock(path: Path, password: str, thumb_path: Path) -> int:
    doc = open_pdf(path, password)
    try:
        render_thumbnail(doc[0], thumb_path)
        return doc.page_count
    finally:
        doc.close()


@app.post("/api/tasks/{task_id}/unlock")
async def unlock(task_id: str, body: UnlockRequest) -> dict:
    """Проверяет пароль и строит превью. Пароль на сервере не сохраняется."""
    meta = _get(task_id)
    if not meta.get("encrypted"):
        return _public(meta)
    d = store.dir(task_id)
    try:
        pages = await asyncio.to_thread(_unlock, d / "input.pdf", body.password, d / "thumb.png")
    except ConversionError as exc:
        raise HTTPException(400, str(exc)) from exc
    if pages > settings.max_pages:
        raise HTTPException(413, f"В документе {pages} стр. — больше допустимых {settings.max_pages}.")
    return _public(store.update(task_id, pages=pages))


class ConvertRequest(BaseModel):
    password: str | None = Field(None, max_length=256)
    pages: str | None = Field(None, max_length=200)
    ocr: str = Field("auto", pattern="^(auto|off)$")
    ocr_lang: str | None = Field(None, max_length=60, pattern=r"^[a-z_+]*$")


@app.post("/api/tasks/{task_id}/convert", status_code=202)
async def convert(task_id: str, body: ConvertRequest) -> dict:
    meta = _get(task_id)
    if meta.get("status") in ACTIVE:
        return _public(meta)
    if meta.get("status") == "uploading":
        raise HTTPException(409, "Файл ещё загружается.")
    d = store.dir(task_id)
    if meta.get("encrypted"):
        if not body.password:
            raise HTTPException(400, "Документ защищён паролем — укажите пароль.")
        try:
            pages = await asyncio.to_thread(_unlock, d / "input.pdf", body.password, d / "thumb.png")
        except ConversionError as exc:
            raise HTTPException(400, str(exc)) from exc
        meta = store.update(task_id, pages=pages)
    if body.pages and meta.get("pages"):
        try:
            parse_page_range(body.pages, meta["pages"])
        except ConversionError as exc:
            raise HTTPException(400, str(exc)) from exc
    options = {"password": body.password, "pages": body.pages, "ocr": body.ocr,
               "ocr_lang": body.ocr_lang or None, "title": Path(meta["name"]).stem}
    meta = store.update(task_id, status="queued", error=None, stats=None, queued=time.time(),
                        options={k: v for k, v in options.items() if k != "password"})
    runner.submit(task_id, options)
    return _public(meta)


@app.get("/api/tasks/{task_id}/download")
def download(task_id: str):
    meta = _get(task_id)
    path = store.dir(task_id) / "result.docx"
    if meta.get("status") != "done" or not path.is_file():
        raise HTTPException(409, "Документ ещё не готов.")
    return FileResponse(path, filename=_output_name(meta["name"]),
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.delete("/api/tasks/{task_id}", status_code=204)
def delete(task_id: str):
    _get(task_id)
    runner.cancel(task_id)
    store.delete(task_id)


@app.get("/api/zip")
def download_zip(ids: str = Query(..., max_length=33 * 100)):
    """Все готовые документы одним ZIP-архивом."""
    names: dict[str, int] = {}
    files: list[tuple[Path, str]] = []
    for task_id in [i for i in ids.split(",") if i][:100]:
        meta = store.load(task_id) if store.valid_id(task_id) else None
        path = store.dir(task_id) / "result.docx" if meta else None
        if not meta or meta.get("status") != "done" or not path.is_file():
            continue
        name = _output_name(meta["name"])
        n = names.get(name.lower(), 0)
        names[name.lower()] = n + 1
        if n:
            name = f"{Path(name).stem} ({n + 1}).docx"
        files.append((path, name))
    if not files:
        raise HTTPException(404, "Нет готовых документов.")
    buf = tempfile.SpooledTemporaryFile(max_size=64 * CHUNK)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, name in files:
            zf.write(path, name)
    buf.seek(0)

    def stream():
        try:
            while chunk := buf.read(CHUNK):
                yield chunk
        finally:
            buf.close()

    return StreamingResponse(stream(), media_type="application/zip",
                             headers={"Content-Disposition": "attachment; filename*=UTF-8''pdf-to-word.zip"})


# --------------------------------------------------------------------------- интерфейс

app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
