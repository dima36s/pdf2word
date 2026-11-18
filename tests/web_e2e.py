"""Сквозная проверка браузерной версии (Playwright + Chrome).

python scripts/build_web.py --out _site
python -m http.server 8080 -d _site &
python tests/web_e2e.py http://127.0.0.1:8080 /tmp/web-shots
"""
from __future__ import annotations

import sys
import time
import zipfile
from pathlib import Path

from playwright.sync_api import expect, sync_playwright

SAMPLES = Path(__file__).parent / "samples"


def main(base: str, out: str) -> None:
    shots = Path(out)
    shots.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", args=["--no-sandbox"])
        page = browser.new_page(viewport={"width": 1360, "height": 900}, locale="ru-RU", accept_downloads=True)
        errors: list[str] = []
        page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
        page.on("console", lambda m: errors.append(f"console: {m.text}") if m.type == "error" else None)

        t0 = time.time()
        page.goto(base)
        page.screenshot(path=shots / "1_loading.png")
        expect(page.locator("#bootBox")).to_have_class("boot is-ready", timeout=180000)
        print(f"конвертер загружен за {time.time() - t0:.1f} с")
        page.screenshot(path=shots / "2_ready.png")

        page.set_input_files("#fileInput", [str(SAMPLES / n) for n in
                                            ("report.pdf", "scan_color.pdf", "two_columns.pdf", "protected.pdf")])
        expect(page.locator(".card")).to_have_count(4)
        card = page.locator(".card", has_text="protected.pdf")
        expect(card.locator(".pw")).to_be_visible(timeout=60000)
        card.locator(".pw input").fill("bad")
        card.locator(".pw button").click()
        expect(card.locator(".card-error")).to_contain_text("Неверный пароль")
        card.locator(".pw input").fill("1234")
        card.locator(".pw button").click()
        expect(card.locator(".pw")).to_be_hidden()
        expect(page.locator("#convertBtn")).to_be_enabled(timeout=60000)
        page.screenshot(path=shots / "3_files.png")

        t1 = time.time()
        page.locator("#convertBtn").click()
        page.wait_for_timeout(4000)
        page.screenshot(path=shots / "4_processing.png")
        expect(page.locator("#doneBanner")).to_be_visible(timeout=300000)
        expect(page.locator(".card.is-done")).to_have_count(4)
        print(f"4 файла сконвертированы за {time.time() - t1:.1f} с")
        page.screenshot(path=shots / "5_done.png")
        for c in page.locator(".card").all():
            print("  ", c.locator(".card-name").inner_text(), "|", c.locator(".card-meta").inner_text())

        with page.expect_download() as dl:
            page.locator(".card", has_text="scan_color.pdf").locator("button.btn-ok").click()
        target = shots / dl.value.suggested_filename
        dl.value.save_as(target)
        assert zipfile.ZipFile(target).read("word/document.xml")
        print("скачан", target.name, target.stat().st_size, "байт")

        with page.expect_download() as dl:
            page.locator("#downloadAllBtn").click()
        ztarget = shots / dl.value.suggested_filename
        dl.value.save_as(ztarget)
        print("архив", ztarget.name, zipfile.ZipFile(ztarget).namelist())

        browser.close()
        bad = [e for e in errors if "status of 4" not in e]
        if bad:
            print("ошибки:", bad)
            sys.exit(1)
        print("web e2e: OK")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8080",
         sys.argv[2] if len(sys.argv) > 2 else "/tmp/web-shots")
