#!/usr/bin/env python3
"""Сборка браузерной версии (статический сайт для GitHub Pages или любого веб-сервера).

    python scripts/build_web.py --out _site

Что попадает в сайт:
  * web/*                 интерфейс и Web Worker
  * pyodide/              Python в WebAssembly + нужные пакеты (PyMuPDF, OpenCV, lxml, numpy, Pillow…)
  * wheels/               python-docx и pdf2docx (чистый Python, с PyPI)
  * app.zip               код конвертера из app/ (тот же, что на сервере)
  * tesseract/            Tesseract.js, ядро WebAssembly и модели rus/eng
  * fonts/                Liberation Serif — шрифт распознанного текста
Источники: релиз Pyodide на GitHub, PyPI, npm. Внешних CDN сайт не использует.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

PYODIDE_VERSION = "0.28.3"            # Python 3.13, в комплекте PyMuPDF 1.26.3
PYODIDE_PACKAGES = ["numpy", "opencv-python", "lxml", "pillow", "pymupdf", "fonttools", "typing-extensions"]
PYODIDE_CORE_FILES = ["pyodide.js", "pyodide.mjs", "pyodide.asm.js", "pyodide.asm.wasm", "python_stdlib.zip",
                      "pyodide-lock.json", "package.json"]
PY_WHEELS = ["python-docx==1.2.0", "pdf2docx==0.5.13"]
TESSERACT_JS = "7.0.0"
TESSERACT_CORES = ["tesseract-core-relaxedsimd-lstm.wasm.js", "tesseract-core-simd-lstm.wasm.js",
                   "tesseract-core-lstm.wasm.js"]
OCR_LANGS = ["rus", "eng"]
APP_MODULES = ["__init__.py", "config.py", "converter.py", "docxfix.py", "furniture.py", "geometry.py", "headers.py", "ocr.py",
               "pdftools.py", "regionfix.py", "textfix.py", "vectorfix.py", "webapi.py"]
FONT_DIRS = ["/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/liberation2",
             "/usr/share/fonts/liberation-serif", "/usr/share/fonts/liberation"]
FONTS = ["LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf"]


def log(msg: str) -> None:
    print(f"[build-web] {msg}", flush=True)


def download(url: str, dst: Path) -> Path:
    if dst.exists():
        return dst
    log(f"download {url}")
    tmp = dst.with_suffix(dst.suffix + ".part")
    with urllib.request.urlopen(url, timeout=600) as r, open(tmp, "wb") as f:
        shutil.copyfileobj(r, f, 1024 * 1024)
    tmp.rename(dst)
    return dst


def pyodide_files(cache: Path, pyodide_dir: Path | None) -> tuple[Path, dict]:
    """Каталог с файлами Pyodide (ядро + пакеты) и lock-файл."""
    if pyodide_dir is None:
        pyodide_dir = cache / f"pyodide-{PYODIDE_VERSION}"
        if not (pyodide_dir / "pyodide-lock.json").exists():
            tar = download(f"https://github.com/pyodide/pyodide/releases/download/{PYODIDE_VERSION}/"
                           f"pyodide-{PYODIDE_VERSION}.tar.bz2", cache / f"pyodide-{PYODIDE_VERSION}.tar.bz2")
            log("extract pyodide")
            pyodide_dir.mkdir(parents=True, exist_ok=True)
            # распаковываем только ядро и наши пакеты (весь архив — ~350 МБ)
            prefixes = tuple(n.replace("-", "_").lower() + "-" for n in PYODIDE_PACKAGES + ["packaging", "micropip"])
            with tarfile.open(tar) as tf:
                for m in tf:
                    name = Path(m.name).name
                    if m.isfile() and (name in PYODIDE_CORE_FILES or name.lower().startswith(prefixes)):
                        m.name = name
                        tf.extract(m, pyodide_dir, filter="data")
    lock = json.loads((pyodide_dir / "pyodide-lock.json").read_text())
    return pyodide_dir, lock


def package_closure(lock: dict, names: list[str]) -> list[dict]:
    pk, seen, order = lock["packages"], set(), []

    def visit(n: str) -> None:
        if n in seen:
            return
        seen.add(n)
        for d in pk[n].get("depends", []):
            visit(d)
        order.append(pk[n])

    for n in names:
        visit(n)
    return order


def build(out: Path, cache: Path, pyodide_dir: Path | None) -> None:
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    cache.mkdir(parents=True, exist_ok=True)

    # 1. интерфейс
    for f in ["index.html", "app.js", "worker.js", "web.css"]:
        shutil.copy(ROOT / "web" / f, out / f)
    shutil.copy(ROOT / "static" / "styles.css", out / "styles.css")
    (out / ".nojekyll").write_text("")

    # 2. Pyodide
    src, lock = pyodide_files(cache, pyodide_dir)
    dst = out / "pyodide"
    dst.mkdir()
    for f in PYODIDE_CORE_FILES:
        shutil.copy(src / f, dst / f)
    pkgs = package_closure(lock, PYODIDE_PACKAGES)
    for p in pkgs:
        if not (src / p["file_name"]).exists():
            sys.exit(f"нет пакета {p['file_name']} в {src}: удалите кэш и запустите сборку заново")
        shutil.copy(src / p["file_name"], dst / p["file_name"])
        got = hashlib.sha256((dst / p["file_name"]).read_bytes()).hexdigest()
        if got != p["sha256"]:
            sys.exit(f"sha256 mismatch: {p['file_name']}")
    log(f"pyodide {PYODIDE_VERSION}: {', '.join(p['name'] for p in pkgs)}")

    # 3. python-docx, pdf2docx
    wheels = out / "wheels"
    wheels.mkdir()
    subprocess.run([sys.executable, "-m", "pip", "download", "-q", "--no-deps", "--only-binary=:all:",
                    "-d", str(wheels), *PY_WHEELS], check=True)

    # 4. код конвертера
    with zipfile.ZipFile(out / "app.zip", "w", zipfile.ZIP_DEFLATED) as zf:
        for m in APP_MODULES:
            zf.write(ROOT / "app" / m, f"app/{m}")

    # 5. Tesseract.js
    tdir = cache / f"tesseract-{TESSERACT_JS}"
    nm = tdir / "node_modules"
    if not (nm / "tesseract.js").exists():
        tdir.mkdir(parents=True, exist_ok=True)
        (tdir / "package.json").write_text('{"name":"tess","version":"1.0.0","private":true}')
        subprocess.run(["npm", "install", "--silent", "--no-audit", "--no-fund", f"tesseract.js@{TESSERACT_JS}",
                        *[f"@tesseract.js-data/{lang}" for lang in OCR_LANGS]], cwd=tdir, check=True)
    tout = out / "tesseract"
    (tout / "core").mkdir(parents=True)
    (tout / "lang").mkdir()
    for f in ["tesseract.min.js", "worker.min.js"]:
        shutil.copy(nm / "tesseract.js" / "dist" / f, tout / f)
    for f in TESSERACT_CORES:
        shutil.copy(nm / "tesseract.js-core" / f, tout / "core" / f)
    for lang in OCR_LANGS:
        shutil.copy(nm / "@tesseract.js-data" / lang / "4.0.0_best_int" / f"{lang}.traineddata.gz",
                    tout / "lang" / f"{lang}.traineddata.gz")

    # 6. шрифты
    fdir = out / "fonts"
    fdir.mkdir()
    for f in FONTS:
        found = next((Path(d) / f for d in FONT_DIRS if (Path(d) / f).is_file()), None)
        if found is None:
            sys.exit(f"не найден шрифт {f}: установите fonts-liberation")
        shutil.copy(found, fdir / f)

    # 7. манифест
    version = os.environ.get("GITHUB_SHA", "")[:7] or "local"
    manifest = {
        "version": version,
        "pyodide_packages": PYODIDE_PACKAGES,
        "wheels": [f"wheels/{w.name}" for w in sorted(wheels.glob("*.whl"))],
        "app_zip": "app.zip",
        "fonts": {f: f"fonts/{f}" for f in FONTS},
    }
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2))

    total = sum(p.stat().st_size for p in out.rglob("*") if p.is_file())
    log(f"готово: {out} ({total / 1e6:.1f} МБ)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="_site", type=Path)
    ap.add_argument("--cache", default=Path(tempfile.gettempdir()) / "pdf2word-web-cache", type=Path)
    ap.add_argument("--pyodide-dir", type=Path, help="готовый каталог Pyodide (иначе скачивается релиз)")
    args = ap.parse_args()
    build(args.out.resolve(), args.cache, args.pyodide_dir)
