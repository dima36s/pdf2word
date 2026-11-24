"""Генерация тестовых PDF разных типов для проверки качества конвертации.

Запуск:  python tests/make_samples.py  (нужны python-docx, reportlab, matplotlib,
LibreOffice в PATH для report.pdf).
"""
from __future__ import annotations

import io
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf

OUT = Path(__file__).parent / "samples"
OUT.mkdir(parents=True, exist_ok=True)

LOREM = (
    "Цифровая трансформация производственных процессов требует согласованной работы "
    "инженерных и информационных систем. В отчёте рассматриваются результаты пилотного "
    "внедрения, ключевые показатели эффективности и план дальнейших работ."
)


def _chart_png() -> bytes:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(5, 2.6), dpi=150)
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн"]
    ax.bar(months, [12, 18, 15, 22, 27, 31], color="#2b579a")
    ax.set_title("Обработано заявок, шт.")
    ax.spines[["top", "right"]].set_visible(False)
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return buf.getvalue()


def make_report() -> Path:
    """Документ Word → PDF через LibreOffice: заголовки, стили, списки, таблица, картинка, колонтитулы."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor, Cm

    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Liberation Serif"
    st.font.size = Pt(12)

    sec = d.sections[0]
    sec.header.paragraphs[0].text = "ООО «Пример» · Внутренний отчёт"
    sec.footer.paragraphs[0].text = "Конфиденциально"
    sec.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    d.add_heading("Отчёт о пилотном проекте", level=0)
    d.add_heading("1. Введение", level=1)
    p = d.add_paragraph(LOREM + " ")
    r = p.add_run("Жирный фрагмент, ")
    r.bold = True
    r = p.add_run("курсив, ")
    r.italic = True
    r = p.add_run("цветной текст")
    r.font.color.rgb = RGBColor(0xC0, 0x1F, 0x2F)
    p.add_run(" и обычный текст после него.")
    p.paragraph_format.first_line_indent = Cm(1.25)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    d.add_heading("2. Цели проекта", level=1)
    for item in ["Сократить время согласования документации", "Снизить число ошибок нормоконтроля",
                 "Обеспечить прозрачность статусов"]:
        d.add_paragraph(item, style="List Bullet")
    for item in ["Анализ текущего процесса", "Разработка прототипа", "Опытная эксплуатация"]:
        d.add_paragraph(item, style="List Number")

    d.add_heading("3. Результаты", level=1)
    d.add_paragraph("Таблица 1 — Ключевые показатели")
    rows = [("Показатель", "До", "После", "Изменение"),
            ("Время согласования, дн.", "14", "6", "−57%"),
            ("Ошибки на 100 листов", "23", "9", "−61%"),
            ("Доля автоматических проверок", "0%", "72%", "+72 п.п."),
            ("Удовлетворённость пользователей", "3,1", "4,4", "+1,3")]
    tbl = d.add_table(rows=len(rows), cols=4)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            if i == 0:
                cell.paragraphs[0].runs[0].bold = True

    d.add_paragraph()
    d.add_picture(io.BytesIO(_chart_png()), width=Cm(14))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = d.add_paragraph("Рисунок 1 — Динамика обработки заявок")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    d.add_page_break()
    d.add_heading("4. План работ", level=1)
    for _ in range(3):
        d.add_paragraph(LOREM)
    t2 = d.add_table(rows=4, cols=3)
    t2.style = "Table Grid"
    hdr = ["Этап", "Срок", "Ответственный"]
    data = [("Масштабирование", "I кв.", "Иванов И.И."), ("Обучение", "II кв.", "Петрова А.С."),
            ("Сдача в промышленную эксплуатацию", "III кв.", "Сидоров П.П.")]
    for j, h in enumerate(hdr):
        t2.cell(0, j).text = h
        t2.cell(0, j).paragraphs[0].runs[0].bold = True
    for i, row in enumerate(data, start=1):
        for j, val in enumerate(row):
            t2.cell(i, j).text = val
    d.add_heading("5. Заключение", level=1)
    d.add_paragraph(LOREM)

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "report.docx"
        d.save(src)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, capture_output=True, timeout=180)
        dst = OUT / "report.pdf"
        shutil.copy(Path(tmp) / "report.pdf", dst)
    return dst


def make_two_columns() -> Path:
    """Двухколоночная вёрстка (как в журнале) через reportlab."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer

    pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
    body = ParagraphStyle("b", fontName="DejaVu", fontSize=9.5, leading=13, alignment=4)
    head = ParagraphStyle("h", fontName="DejaVu-Bold", fontSize=12, leading=16, spaceBefore=6, spaceAfter=4)

    dst = OUT / "two_columns.pdf"
    doc = BaseDocTemplate(str(dst), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                          topMargin=2 * cm, bottomMargin=2 * cm)
    gap = 0.8 * cm
    w = (doc.width - gap) / 2
    frames = [Frame(doc.leftMargin, doc.bottomMargin, w, doc.height, id="c1"),
              Frame(doc.leftMargin + w + gap, doc.bottomMargin, w, doc.height, id="c2")]
    doc.addPageTemplates([PageTemplate(id="two", frames=frames)])
    story = []
    for k in range(1, 7):
        story.append(Paragraph(f"Раздел {k}. Методика исследования", head))
        for _ in range(2):
            story.append(Paragraph(LOREM + " " + LOREM, body))
            story.append(Spacer(1, 4))
    doc.build(story)
    return dst


def make_borderless_table() -> Path:
    """Таблица без рамок (прайс-лист) — проверка, что «потоковые» таблицы не теряются."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
    body = ParagraphStyle("b", fontName="DejaVu", fontSize=10, leading=14)
    head = ParagraphStyle("h", fontName="DejaVu-Bold", fontSize=14, leading=18, spaceAfter=8)
    dst = OUT / "borderless_table.pdf"
    doc = SimpleDocTemplate(str(dst), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm)
    data = [["Позиция", "Ед.", "Кол-во", "Цена, руб.", "Сумма, руб."],
            ["Сервер приложений", "шт.", "2", "450 000", "900 000"],
            ["Лицензия СУБД", "шт.", "4", "120 000", "480 000"],
            ["Внедрение и настройка", "ч", "160", "3 500", "560 000"],
            ["Обучение персонала", "ч", "24", "4 000", "96 000"],
            ["Итого", "", "", "", "2 036 000"]]
    t = Table(data, colWidths=[6.5 * cm, 1.5 * cm, 2 * cm, 3 * cm, 3.5 * cm])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "DejaVu", 10),
        ("FONT", (0, 0), (-1, 0), "DejaVu-Bold", 10),
        ("FONT", (0, -1), (-1, -1), "DejaVu-Bold", 10),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#2b579a")),
    ]))
    story = [Paragraph("Коммерческое предложение № 15", head),
             Paragraph("Предлагаем следующий состав работ и оборудования для реализации проекта. "
                       "Цены указаны с учётом НДС, срок действия предложения — 30 дней.", body),
             Spacer(1, 12), t, Spacer(1, 12),
             Paragraph("Условия оплаты: 50% предоплата, 50% по факту поставки.", body)]
    doc.build(story)
    return dst


INVOICE_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><style>
body{font-family:'Liberation Sans',Arial;font-size:11pt;color:#222;margin:0}
.top{display:flex;justify-content:space-between;border-bottom:3px solid #2b579a;padding-bottom:12px}
.logo{font-size:20pt;font-weight:bold;color:#2b579a}.req{font-size:9.5pt;line-height:1.45;text-align:right}
h1{font-size:18pt;margin:22px 0 6px}.muted{color:#666}
.parties{display:flex;gap:30px;margin:16px 0}.parties div{flex:1;background:#f3f6fb;padding:10px 12px;border-radius:6px;font-size:10pt;line-height:1.5}
table{width:100%;border-collapse:collapse;margin-top:14px;font-size:10pt}
th{background:#2b579a;color:#fff;text-align:left;padding:7px 8px}td{padding:7px 8px;border-bottom:1px solid #d5dbe5}
tr:nth-child(even) td{background:#f6f8fc}.r{text-align:right}
.tot{margin:12px 0 0 auto;width:45%}.tot div{display:flex;justify-content:space-between;padding:4px 0}
.tot .big{font-weight:bold;font-size:13pt;border-top:2px solid #222}ol{font-size:9.5pt;color:#444}
</style></head><body>
<div class=top><div class=logo>ТехноСервис</div><div class=req>ООО «ТехноСервис»<br>ИНН 7801234567 / КПП 780101001<br>
г. Санкт-Петербург, ул. Примерная, д. 1<br>+7 (812) 000-00-00</div></div>
<h1>Счёт на оплату № 2481 от 18 ноября 2026 г.</h1><div class=muted>Срок оплаты — 5 банковских дней</div>
<div class=parties><div><b>Поставщик</b><br>ООО «ТехноСервис»<br>р/с 40702810000000000001<br>ПАО «Банк», БИК 044030000</div>
<div><b>Покупатель</b><br>АО «Промышленная компания»<br>ИНН 7700000000<br>г. Москва, Ленинский пр-т, д. 10</div></div>
<table><tr><th>№</th><th>Наименование</th><th class=r>Кол-во</th><th>Ед.</th><th class=r>Цена</th><th class=r>Сумма</th></tr>
<tr><td>1</td><td>Ноутбук 15.6" Core i7, 16 ГБ</td><td class=r>5</td><td>шт.</td><td class=r>98 500,00</td><td class=r>492 500,00</td></tr>
<tr><td>2</td><td>Монитор 27" IPS 2560×1440</td><td class=r>5</td><td>шт.</td><td class=r>31 200,00</td><td class=r>156 000,00</td></tr>
<tr><td>3</td><td>Док-станция USB-C</td><td class=r>5</td><td>шт.</td><td class=r>12 900,00</td><td class=r>64 500,00</td></tr>
<tr><td>4</td><td>Настройка рабочих мест</td><td class=r>8</td><td>ч</td><td class=r>2 500,00</td><td class=r>20 000,00</td></tr>
<tr><td>5</td><td>Доставка</td><td class=r>1</td><td>усл.</td><td class=r>3 000,00</td><td class=r>3 000,00</td></tr></table>
<div class=tot><div><span>Итого:</span><span>736 000,00</span></div><div><span>В т.ч. НДС 20%:</span><span>122 666,67</span></div>
<div class=big><span>Всего к оплате:</span><span>736 000,00 ₽</span></div></div>
<p>Всего наименований 5, на сумму <b>736 000,00 руб.</b></p>
<ol><li>Счёт действителен в течение 5 банковских дней.</li><li>При себе иметь доверенность и паспорт.</li></ol>
</body></html>"""


def make_invoice_chrome() -> Path | None:
    """Счёт, напечатанный в PDF браузером Chrome (плашки с фоном, таблица-«зебра»). Нужен Playwright."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("skip invoice_chrome.pdf: playwright не установлен")
        return None
    dst = OUT / "invoice_chrome.pdf"
    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox"])
        page = browser.new_page()
        page.set_content(INVOICE_HTML)
        page.pdf(path=str(dst), format="A4", print_background=True,
                 margin={"top": "18mm", "bottom": "18mm", "left": "16mm", "right": "16mm"})
        browser.close()
    return dst


def make_scan(src: Path) -> Path:
    """«Скан»: первая страница отчёта, растрированная в картинку без текстового слоя."""
    import random
    from PIL import Image, ImageFilter

    s = pymupdf.open(src)
    out = pymupdf.open()
    for pno in range(min(2, s.page_count)):
        pg = s[pno]
        pix = pg.get_pixmap(dpi=200, colorspace=pymupdf.csGRAY)
        img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        img = img.rotate(0.4, fillcolor=255, resample=Image.BICUBIC).filter(ImageFilter.GaussianBlur(0.6))
        px = img.load()
        rnd = random.Random(1)
        for _ in range(6000):
            x, y = rnd.randrange(img.width), rnd.randrange(img.height)
            px[x, y] = rnd.choice((0, 255))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=80)
        np = out.new_page(width=pg.rect.width, height=pg.rect.height)
        np.insert_image(np.rect, stream=buf.getvalue())
    dst = OUT / "scan.pdf"
    out.save(dst, deflate=True)
    return dst


def make_color_scan(src: Path) -> Path:
    """Чистый цветной скан (200 dpi, наклон 1.2°), как с офисного сканера."""
    from PIL import Image

    s = pymupdf.open(src)
    pg = s[0]
    pix = pg.get_pixmap(dpi=200)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img = img.rotate(-1.2, fillcolor=(250, 250, 248), resample=Image.BICUBIC)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    out = pymupdf.open()
    np_ = out.new_page(width=pg.rect.width, height=pg.rect.height)
    np_.insert_image(np_.rect, stream=buf.getvalue())
    dst = OUT / "scan_color.pdf"
    out.save(dst, deflate=True)
    return dst


def make_encrypted(src: Path) -> Path:
    """PDF с паролем на открытие «1234» (AES-256)."""
    dst = OUT / "protected.pdf"
    if shutil.which("qpdf"):
        subprocess.run(["qpdf", "--encrypt", "1234", "owner", "256", "--", str(src), str(dst)], check=True)
    else:
        s = pymupdf.open(src)
        s.save(dst, encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="1234", owner_pw="owner")
    return dst


if __name__ == "__main__":
    from gost_samples import make_gost_spec, make_gost_text
    print("ok", make_gost_text())
    print("ok", make_gost_spec())
    rep = make_report()
    print("ok", rep)
    print("ok", make_two_columns())
    print("ok", make_borderless_table())
    print("ok", make_invoice_chrome())
    print("ok", make_scan(rep))
    print("ok", make_color_scan(rep))
    print("ok", make_encrypted(rep))
    sys.exit(0)
