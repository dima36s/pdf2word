"""Задачи конвертации: хранение на диске и запуск в отдельных процессах.

Каждая задача — папка data/tasks/<id>/:
    input.pdf      загруженный файл
    thumb.png      превью первой страницы
    meta.json      состояние задачи (статус, имя, параметры, итог)
    progress.json  прогресс от рабочего процесса
    result.json    итог от рабочего процесса
    result.docx    готовый документ
    result.xlsx    таблицы и текст в Excel (если было что выгрузить)
Состояние хранится в файлах, поэтому его видят все процессы сервера.
"""
from __future__ import annotations

import asyncio
import json
import logging
import multiprocessing as mp
import os
import re
import shutil
import time
import uuid
from pathlib import Path

from .worker import worker_main

log = logging.getLogger(__name__)

_ID_RE = re.compile(r"^[0-9a-f]{32}$")
ACTIVE = ("queued", "processing")


class TaskStore:
    def __init__(self, root: Path):
        self.root = self._writable(root)

    @staticmethod
    def _writable(root: Path) -> Path:
        """DATA_DIR может оказаться недоступен для записи (хостинг запускает контейнер под другим
        uid, OpenShift и т.п.) — тогда работаем во временном каталоге."""
        try:
            root.mkdir(parents=True, exist_ok=True)
            probe = root / ".write-test"
            probe.write_text("ok")
            probe.unlink()
            return root
        except OSError as exc:
            import tempfile
            fallback = Path(tempfile.gettempdir()) / "pdf2word" / "tasks"
            fallback.mkdir(parents=True, exist_ok=True)
            log.warning("DATA_DIR %s недоступен для записи (%s), использую %s", root, exc, fallback)
            return fallback

    # ------------------------------------------------------------------ пути
    @staticmethod
    def valid_id(task_id: str) -> bool:
        return bool(_ID_RE.match(task_id or ""))

    def dir(self, task_id: str) -> Path:
        if not self.valid_id(task_id):
            raise KeyError(task_id)
        return self.root / task_id

    # ------------------------------------------------------------------ meta
    def create(self, name: str, size: int) -> dict:
        task_id = uuid.uuid4().hex
        self.dir(task_id).mkdir(parents=True)
        meta = {"id": task_id, "name": name, "size": size, "created": time.time(), "status": "uploading"}
        self.save(meta)
        return meta

    def load(self, task_id: str) -> dict | None:
        try:
            return json.loads((self.dir(task_id) / "meta.json").read_text(encoding="utf-8"))
        except (KeyError, FileNotFoundError, json.JSONDecodeError):
            return None

    def save(self, meta: dict) -> None:
        meta["updated"] = time.time()
        path = self.dir(meta["id"]) / "meta.json"
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, path)

    def update(self, task_id: str, **fields) -> dict | None:
        meta = self.load(task_id)
        if meta is None:
            return None
        meta.update(fields)
        self.save(meta)
        return meta

    def read_json(self, task_id: str, name: str) -> dict | None:
        try:
            return json.loads((self.dir(task_id) / name).read_text(encoding="utf-8"))
        except (KeyError, FileNotFoundError, json.JSONDecodeError):
            return None

    def delete(self, task_id: str) -> None:
        try:
            shutil.rmtree(self.dir(task_id), ignore_errors=True)
        except KeyError:
            pass

    def all_ids(self) -> list[str]:
        return [p.name for p in self.root.iterdir() if p.is_dir() and self.valid_id(p.name)]

    # ------------------------------------------------------------------ обслуживание
    def cleanup(self, ttl_seconds: float, running: set[str]) -> int:
        now, removed = time.time(), 0
        for task_id in self.all_ids():
            if task_id in running:
                continue
            meta = self.load(task_id)
            created = (meta or {}).get("created") or (self.dir(task_id).stat().st_mtime)
            if now - created > ttl_seconds:
                self.delete(task_id)
                removed += 1
        return removed

    def recover_after_restart(self) -> None:
        """Задачи, которые были в работе при остановке сервера, помечаются ошибкой."""
        for task_id in self.all_ids():
            meta = self.load(task_id)
            if meta and meta.get("status") in ACTIVE + ("uploading",):
                meta.update(status="error", error="Сервер был перезапущен во время обработки. Запустите конвертацию снова.")
                self.save(meta)


class Runner:
    """Ограничивает число одновременных конвертаций и следит за таймаутом."""

    def __init__(self, store: TaskStore, workers: int, timeout: int):
        self.store = store
        self.timeout = timeout
        self.sem = asyncio.Semaphore(max(1, workers))
        self.ctx = mp.get_context("spawn")
        self.procs: dict[str, mp.process.BaseProcess] = {}
        self.queue: list[str] = []
        self._bg: set[asyncio.Task] = set()

    @property
    def running(self) -> set[str]:
        return set(self.procs) | set(self.queue)

    def queue_position(self, task_id: str) -> int | None:
        try:
            return self.queue.index(task_id) + 1
        except ValueError:
            return None

    def submit(self, task_id: str, options: dict) -> None:
        self.queue.append(task_id)
        t = asyncio.create_task(self._run(task_id, options))
        self._bg.add(t)
        t.add_done_callback(self._bg.discard)

    async def _run(self, task_id: str, options: dict) -> None:
        try:
            async with self.sem:
                if task_id in self.queue:
                    self.queue.remove(task_id)
                meta = self.store.load(task_id)
                if meta is None or meta.get("status") != "queued":
                    return  # удалена или отменена, пока ждала очереди
                await self._execute(task_id, options)
        except Exception:
            log.exception("runner failed for %s", task_id)
            self.store.update(task_id, status="error", error="Внутренняя ошибка сервера.")
        finally:
            if task_id in self.queue:
                self.queue.remove(task_id)

    async def _execute(self, task_id: str, options: dict) -> None:
        d = self.store.dir(task_id)
        for name in ("progress.json", "result.json", "result.docx", "result.xlsx"):
            (d / name).unlink(missing_ok=True)
        self.store.update(task_id, status="processing", started=time.time(), error=None, stats=None)
        proc = self.ctx.Process(target=worker_main, args=(str(d), options), daemon=True)
        proc.start()
        self.procs[task_id] = proc
        try:
            await asyncio.to_thread(proc.join, self.timeout)
            timed_out = proc.is_alive()
            if timed_out:
                proc.terminate()
                await asyncio.to_thread(proc.join, 5)
                if proc.is_alive():
                    proc.kill()
        finally:
            self.procs.pop(task_id, None)

        if self.store.load(task_id) is None:
            return  # задачу удалили во время работы
        result = self.store.read_json(task_id, "result.json")
        if timed_out:
            limit = f"{self.timeout // 60} мин" if self.timeout >= 60 else f"{self.timeout} с"
            self.store.update(task_id, status="error",
                              error=f"Конвертация заняла больше {limit} и была остановлена. "
                                    "Попробуйте указать диапазон страниц или отключить OCR.")
        elif result and result.get("ok") and (d / "result.docx").is_file():
            self.store.update(task_id, status="done", finished=time.time(), stats=result.get("stats"),
                              output_size=(d / "result.docx").stat().st_size)
        else:
            err = (result or {}).get("error") or "Процесс конвертации завершился аварийно (возможно, не хватило памяти)."
            if result and result.get("detail"):
                log.warning("task %s failed: %s", task_id, result["detail"])
            self.store.update(task_id, status="error", error=err)

    def cancel(self, task_id: str) -> None:
        if task_id in self.queue:
            self.queue.remove(task_id)
        proc = self.procs.get(task_id)
        if proc is not None and proc.is_alive():
            proc.terminate()

    async def shutdown(self) -> None:
        for proc in list(self.procs.values()):
            if proc.is_alive():
                proc.terminate()
        for t in list(self._bg):
            t.cancel()
