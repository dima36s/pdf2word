"""Сквозная проверка интерфейса в браузере (Playwright + Chrome).

python tests/ui_e2e.py http://127.0.0.1:8000 /tmp/shots
"""
from __future__ import annotations

import sys
from pathlib import Path

from playwright.sync_api import expect, sync_playwright

SAMPLES = Path(__file__).parent / "samples"


def main(base: str, out: str) -> None:
    shots = Path(out)
    shots.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", args=["--no-sandbox"])
        page = browser.new_page(viewport={"width": 1360, "height": 900}, locale="ru-RU")
        errors: list[str] = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        # ответы 4xx на неверный пароль ожидаемы — браузер пишет их в консоль, это не ошибки JS
        page.on("console", lambda m: errors.append(m.text)
                if m.type == "error" and "status of 400" not in m.text else None)

        page.goto(base)
        expect(page.locator("#limits")).to_contain_text("МБ")
        page.screenshot(path=shots / "1_hero.png")

        page.set_input_files("#fileInput", [str(SAMPLES / n) for n in
                                            ("report.pdf", "scan_color.pdf", "two_columns.pdf", "protected.pdf")])
        expect(page.locator(".card")).to_have_count(4)
        expect(page.locator("#convertBtn")).to_be_enabled(timeout=20000)
        page.wait_for_timeout(800)
        page.screenshot(path=shots / "2_uploaded.png")

        # пароль: сначала неверный, затем верный
        card = page.locator(".card", has_text="protected.pdf")
        card.locator(".pw input").fill("bad")
        card.locator(".pw button").click()
        expect(card.locator(".card-error")).to_contain_text("Неверный пароль")
        card.locator(".pw input").fill("1234")
        card.locator(".pw button").click()
        expect(card.locator(".pw")).to_be_hidden()
        expect(card.locator(".pages-input")).to_be_visible()

        # диапазон страниц для отчёта
        page.locator(".card", has_text="report.pdf").locator(".pages-input").fill("1-2")

        page.locator("#convertBtn").click()
        page.wait_for_timeout(1200)
        page.screenshot(path=shots / "3_processing.png")

        expect(page.locator("#doneBanner")).to_be_visible(timeout=120000)
        expect(page.locator(".card.is-done")).to_have_count(4)
        page.screenshot(path=shots / "4_done.png")

        href = page.locator("#downloadAllBtn").get_attribute("href")
        assert href and href.startswith("api/zip?ids="), href
        with page.expect_download() as dl:
            page.locator(".card", has_text="report.pdf").locator("a.btn-ok").click()
        d = dl.value
        target = shots / d.suggested_filename
        d.save_as(target)
        print("downloaded:", target.name, target.stat().st_size, "bytes")

        # мобильная ширина
        page.set_viewport_size({"width": 390, "height": 860})
        page.wait_for_timeout(300)
        page.screenshot(path=shots / "5_mobile.png", full_page=True)

        # «Конвертировать ещё» возвращает на стартовый экран
        page.locator("#restartBtn").click()
        expect(page.locator("#hero")).to_be_visible()

        browser.close()
        if errors:
            print("JS errors:", errors)
            sys.exit(1)
        print("UI e2e: OK")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8000",
         sys.argv[2] if len(sys.argv) > 2 else "/tmp/shots")
