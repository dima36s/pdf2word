"""Точка входа рабочего процесса: одна конвертация = один процесс.

Отдельный процесс даёт: изоляцию падений и утечек памяти pdf2docx/MuPDF, жёсткий таймаут
(процесс можно убить), параллельность без GIL. Связь с сервером — через файлы в папке задачи.
"""
from __future__ import annotations

import json
import logging
import os
import time
import warnings
from pathlib import Path


def _write_json(path: Path, data: dict) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def _limit_memory() -> None:
    mb = int(os.environ.get("WORKER_MEMORY_MB", "0") or 0)
    if mb <= 0:
        return
    try:
        import resource
        resource.setrlimit(resource.RLIMIT_AS, (mb * 1024 * 1024, mb * 1024 * 1024))
    except Exception:
        pass


def worker_main(task_dir: str, options: dict) -> None:
    warnings.filterwarnings("ignore")
    logging.basicConfig(level=logging.WARNING, format="%(asctime)s %(levelname)s [worker] %(message)s")
    logging.getLogger("fontTools").setLevel(logging.ERROR)
    _limit_memory()

    from .converter import ConvertOptions, convert_pdf
    from .pdftools import ConversionError

    d = Path(task_dir)
    progress_file = d / "progress.json"
    last = {"pct": -1.0, "stage": "", "t": 0.0}

    def progress(pct: float, stage: str) -> None:
        now = time.monotonic()
        if stage == last["stage"] and pct - last["pct"] < 1 and now - last["t"] < 1:
            return
        last.update(pct=pct, stage=stage, t=now)
        try:
            _write_json(progress_file, {"progress": round(min(max(pct, 0), 100), 1), "stage": stage})
        except OSError:
            pass

    try:
        stats = convert_pdf(d / "input.pdf", d / "result.docx", ConvertOptions(**options), progress,
                            xlsx_dst=d / "result.xlsx")
        _write_json(d / "result.json", {"ok": True, "stats": stats.to_dict()})
    except ConversionError as exc:
        _write_json(d / "result.json", {"ok": False, "error": str(exc)})
    except MemoryError:
        _write_json(d / "result.json", {"ok": False, "error": "Документ слишком большой: не хватило памяти."})
    except Exception as exc:  # неожиданная ошибка — в лог, пользователю общий текст
        logging.exception("conversion failed")
        _write_json(d / "result.json", {"ok": False, "error": "Не удалось сконвертировать документ: "
                                                              "внутренняя ошибка обработки PDF.",
                                        "detail": f"{type(exc).__name__}: {exc}"[:500]})
