"""Тесты HTTP API. Запуск: pytest -q tests/test_api.py"""
from __future__ import annotations

import io
import os
import tempfile
import time
import zipfile
from pathlib import Path

import pytest

os.environ.setdefault("DATA_DIR", tempfile.mkdtemp(prefix="pdf2word-test-"))
os.environ.setdefault("CONVERT_WORKERS", "2")

from fastapi.testclient import TestClient  # noqa: E402

from app.main import app, store  # noqa: E402

SAMPLES = Path(__file__).parent / "samples"


@pytest.fixture(scope="module")
def client():
    if not (SAMPLES / "report.pdf").exists():
        pytest.skip("сначала запустите python tests/make_samples.py")
    with TestClient(app) as c:
        yield c


def upload(client, name: str, data: bytes | None = None, ctype="application/pdf"):
    data = data if data is not None else (SAMPLES / name).read_bytes()
    return client.post("/api/tasks", files={"file": (name, io.BytesIO(data), ctype)})


def wait_done(client, task_id: str, timeout: float = 120) -> dict:
    end = time.time() + timeout
    while time.time() < end:
        st = client.get(f"/api/tasks/{task_id}").json()
        if st["status"] in ("done", "error"):
            return st
        time.sleep(0.4)
    raise AssertionError("conversion timeout")


def test_config(client):
    cfg = client.get("/api/config").json()
    assert cfg["max_file_mb"] > 0
    assert "languages" in cfg["ocr"]


def test_full_cycle(client):
    r = upload(client, "Отчёт о проекте.pdf", (SAMPLES / "report.pdf").read_bytes())
    assert r.status_code == 201, r.text
    task = r.json()
    assert task["pages"] == 3 and task["has_thumb"] and task["output_name"] == "Отчёт о проекте.docx"
    assert client.get(f"/api/tasks/{task['id']}/thumb").headers["content-type"] == "image/png"

    r = client.post(f"/api/tasks/{task['id']}/convert", json={"ocr": "auto"})
    assert r.status_code == 202
    st = wait_done(client, task["id"])
    assert st["status"] == "done", st
    assert st["stats"]["tables"] == 2

    r = client.get(f"/api/tasks/{task['id']}/download")
    assert r.status_code == 200
    assert "filename*=utf-8''" in r.headers["content-disposition"].lower()
    assert zipfile.ZipFile(io.BytesIO(r.content)).read("word/document.xml")

    # таблицы и текст — в Excel
    assert st["has_xlsx"] and st["stats"]["xlsx_sheets"] == 3
    r = client.get(f"/api/tasks/{task['id']}/download?format=xlsx")
    assert r.status_code == 200
    assert "spreadsheetml" in r.headers["content-type"]
    assert ".xlsx" in r.headers["content-disposition"]
    assert zipfile.ZipFile(io.BytesIO(r.content)).read("xl/workbook.xml")
    assert client.get(f"/api/tasks/{task['id']}/download?format=pdf").status_code == 422

    r = client.get(f"/api/zip?ids={task['id']}")
    assert r.status_code == 200
    names = zipfile.ZipFile(io.BytesIO(r.content)).namelist()
    assert names == ["Отчёт о проекте.docx"]

    assert client.delete(f"/api/tasks/{task['id']}").status_code == 204
    assert client.get(f"/api/tasks/{task['id']}").status_code == 404


def test_rejects_non_pdf(client):
    r = upload(client, "fake.pdf", b"hello world")
    assert r.status_code == 415
    r = upload(client, "image.png", b"\x89PNG....", "image/png")
    assert r.status_code == 415


def test_password_flow(client):
    task = upload(client, "protected.pdf").json()
    assert task["encrypted"] and task["pages"] is None
    assert client.post(f"/api/tasks/{task['id']}/convert", json={}).status_code == 400
    r = client.post(f"/api/tasks/{task['id']}/unlock", json={"password": "nope"})
    assert r.status_code == 400 and "Неверный" in r.json()["detail"]
    r = client.post(f"/api/tasks/{task['id']}/unlock", json={"password": "1234"})
    assert r.status_code == 200 and r.json()["pages"] == 3
    r = client.post(f"/api/tasks/{task['id']}/convert", json={"password": "1234", "pages": "1"})
    assert r.status_code == 202
    st = wait_done(client, task["id"])
    assert st["status"] == "done" and st["stats"]["pages_converted"] == 1
    # пароль не сохраняется в метаданных задачи
    meta = (store.dir(task["id"]) / "meta.json").read_text(encoding="utf-8")
    assert "1234" not in meta


def test_bad_page_range(client):
    task = upload(client, "report.pdf").json()
    r = client.post(f"/api/tasks/{task['id']}/convert", json={"pages": "7-9"})
    assert r.status_code == 400 and "страницы 7 нет" in r.json()["detail"]
    r = client.post(f"/api/tasks/{task['id']}/convert", json={"ocr": "maybe"})
    assert r.status_code == 422


def test_parallel_queue(client):
    ids = [upload(client, n).json()["id"] for n in ("two_columns.pdf", "borderless_table.pdf", "report.pdf")]
    for i in ids:
        assert client.post(f"/api/tasks/{i}/convert", json={}).status_code == 202
    results = [wait_done(client, i) for i in ids]
    assert all(r["status"] == "done" for r in results), results


def test_bad_ids(client):
    assert client.get("/api/tasks/not-an-id").status_code == 404
    assert client.get("/api/tasks/" + "0" * 32 + "/download").status_code == 404
    assert client.get("/api/zip?ids=../../etc").status_code == 404
