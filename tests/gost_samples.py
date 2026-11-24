"""Синтетические листы проектной документации (только reportlab) — по мотивам замечаний тестировщиков.

* gost_text.pdf — текстовая часть / опросные листы: шапка (логотип из кривых, шифр, линия под ней),
  «Содержание» с отточиями и уровнями, нумерованные заголовки, таблица, нижний штамп-рамка с годом,
  названием и номером листа, имя файла под ним; третий лист — A3 альбомный.
* gost_spec.pdf — спецификация по ГОСТ 21.110: рамка листа ломаной с «плавающими» координатами
  (как из САПР), таблица спецификации, основная надпись (на 1-м листе — большая, дальше — малая),
  боковой штамп согласования с вертикальным текстом, логотип в штампе.

Запуск: python tests/gost_samples.py (вызывается и из make_samples.py).
"""
from __future__ import annotations

from pathlib import Path

from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

OUT = Path(__file__).parent / "samples"
CODE = "Е110-0162-УКК.25.1087-РД-01-03.03.086-АТХ.ОЛ"
SPEC_CODE = "Е110-0162-УКК.25.1087-РД-01-03.03.086-АТХ.СО"

_FONT_DIRS = ["/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/liberation2",
              "/usr/share/fonts/liberation-sans", "/usr/share/fonts/liberation"]


def _fonts() -> tuple[str, str]:
    for d in _FONT_DIRS:
        reg, bold = Path(d) / "LiberationSans-Regular.ttf", Path(d) / "LiberationSans-Bold.ttf"
        if reg.exists() and bold.exists():
            if "GSans" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("GSans", str(reg)))
                pdfmetrics.registerFont(TTFont("GSans-Bold", str(bold)))
            return "GSans", "GSans-Bold"
    raise RuntimeError("нужны шрифты Liberation Sans (fonts-liberation)")


def _logo(c: canvas.Canvas, x: float, y: float, s: float = 1.0) -> None:
    """Логотип: две «волны» кривыми Безье + надпись (текст)."""
    reg, bold = _fonts()
    c.saveState()
    c.setFillColorRGB(0.12, 0.12, 0.14)
    for dx in (0, 7 * s):
        p = c.beginPath()
        p.moveTo(x + dx, y)
        p.curveTo(x + dx + 4 * s, y + 14 * s, x + dx + 10 * s, y + 18 * s, x + dx + 16 * s, y + 20 * s)
        p.curveTo(x + dx + 10 * s, y + 12 * s, x + dx + 8 * s, y + 6 * s, x + dx + 6 * s, y)
        p.close()
        c.drawPath(p, stroke=0, fill=1)
    c.setFont(bold, 13 * s)
    c.drawString(x + 27 * s, y + 6 * s, "ПроТех")
    c.setFont(reg, 4.5 * s)
    c.drawString(x + 27 * s, y + 1 * s, "ИНЖИНИРИНГ")
    c.restoreState()


def _polyline(c: canvas.Canvas, pts, width: float = 0.5, drift: float = 0.0) -> None:
    """Ломаная одной командой пути; drift — «плавание» координат, как в выгрузке из САПР."""
    p = c.beginPath()
    x, y = pts[0]
    p.moveTo(x, y)
    for k, (x, y) in enumerate(pts[1:], 1):
        p.lineTo(x + (drift if k % 2 else 0), y + (drift if k % 2 == 0 else 0))
    c.setLineWidth(width)
    c.drawPath(p, stroke=1, fill=0)


def _box(c, x0, y0, x1, y1, width=0.5):
    _polyline(c, [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)], width)


# --------------------------------------------------------------------------- текстовая часть / ОЛ

def _text_header(c: canvas.Canvas, w: float, h: float) -> None:
    reg, _ = _fonts()
    left, right = 20 * mm, w - 10 * mm
    _logo(c, left, h - 17 * mm)
    c.setFont(reg, 8)
    c.drawRightString(right, h - 12 * mm, CODE)
    c.setLineWidth(0.6)
    c.line(left, h - 19 * mm, right, h - 19 * mm)


def _text_footer(c: canvas.Canvas, w: float, sheet: int, title: tuple[str, str]) -> None:
    reg, bold = _fonts()
    left, right = 20 * mm, w - 10 * mm
    y0, y1 = 14 * mm, 24 * mm
    _box(c, left, y0, right, y1, 0.7)
    c.line(left + 16 * mm, y0, left + 16 * mm, y1)
    c.line(right - 12 * mm, y0, right - 12 * mm, y1)
    c.setFont(bold, 8)
    c.drawString(left + 3 * mm, y0 + 3.6 * mm, "2026")
    c.setFont(reg, 8)
    c.drawString(left + 18 * mm, y0 + 5.8 * mm, title[0])
    c.drawString(left + 18 * mm, y0 + 2.0 * mm, title[1])
    c.setFont(bold, 9)
    c.drawCentredString(right - 6 * mm, y0 + 3.5 * mm, str(sheet))
    c.setFont(reg, 6.5)
    c.drawRightString(right, y0 - 3.5 * mm, f"{CODE}_0_0_RU_IFC.docx")


TOC = [
    (0, "Обозначения и сокращения", "3"),
    (0, "1 Общие положения", "4"),
    (0, "2 Характеристика принятой технологической схемы производства в целом и характеристика "
        "отдельных параметров технологического процесса", "5"),
    (1, "2.1 Общие характеристики принятой технологической схемы", "5"),
    (1, "2.2 Характеристика отдельных параметров технологического процесса", "6"),
    (2, "2.2.1 Конвейерно-транспортная галерея подачи руды", "7"),
    (1, "2.3 Требования к организации производства", "8"),
    (0, "3 Обоснование потребности в основных видах ресурсов", "9"),
]

BODY = ("Технологическая схема предусматривает переработку руды по схеме с флотационным обогащением. "
        "Оборудование размещается в главном корпусе фабрики, транспортные операции выполняются "
        "конвейерами. Параметры процесса приняты по результатам испытаний и опыту эксплуатации "
        "аналогичных предприятий.")


def _wrap(text: str, font: str, size: float, width: float) -> list[str]:
    words, lines, cur = text.split(), [], ""
    for wd in words:
        trial = f"{cur} {wd}".strip()
        if pdfmetrics.stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            lines.append(cur)
            cur = wd
    if cur:
        lines.append(cur)
    return lines


def make_gost_text() -> Path:
    reg, bold = _fonts()
    dst = OUT / "gost_text.pdf"
    c = canvas.Canvas(str(dst), pagesize=A4)
    w, h = A4
    left, right = 25 * mm, w - 15 * mm
    title = ("Трубопровод технического конденсата от насосной станции 03.03.130 до корпуса",
             "02.20.203. Опросный лист на расходомеры")

    # лист 1 — содержание
    _text_header(c, w, h)
    c.setFont(bold, 12)
    c.drawCentredString(w / 2, h - 32 * mm, "Содержание")
    y = h - 42 * mm
    c.setFont(reg, 11)
    for level, text, page in TOC:
        x = left + level * 7 * mm
        lines = _wrap(text, reg, 11, right - x - 14 * mm)
        for k, ln in enumerate(lines):
            c.drawString(x if k == 0 else x, y, ln)
            if k == len(lines) - 1:
                start = x + pdfmetrics.stringWidth(ln, reg, 11) + 2
                dots = "." * max(0, int((right - 8 * mm - start) / pdfmetrics.stringWidth(".", reg, 11)))
                c.drawString(start, y, dots)
                c.drawRightString(right, y, page)
            y -= 6.2 * mm
    _text_footer(c, w, 2, title)
    c.showPage()

    # лист 2 — заголовки, текст, таблица
    _text_header(c, w, h)
    y = h - 34 * mm
    c.setFont(bold, 12)
    c.drawString(left, y, "2")
    for k, ln in enumerate(_wrap("Характеристика принятой технологической схемы производства в целом "
                                 "и характеристика отдельных параметров технологического процесса",
                                 bold, 12, right - left - 12 * mm)):
        c.drawString(left + 10 * mm, y - k * 6 * mm, ln)
    y -= 20 * mm
    c.setFont(bold, 11)
    c.drawString(left, y, "2.1")
    c.drawString(left + 10 * mm, y, "Общие характеристики принятой технологической схемы")
    y -= 9 * mm
    c.setFont(reg, 11)
    for ln in _wrap(BODY, reg, 11, right - left):
        c.drawString(left, y, ln)
        y -= 5.6 * mm
    y -= 6 * mm
    cols = [left, left + 15 * mm, left + 95 * mm, right]
    rows = [("№", "Наименование", "Значение"), ("1", "Производительность, т/ч", "420"),
            ("2", "Крупность руды, мм", "0–300"), ("3", "Влажность, %", "4,5")]
    top = y
    for r, row in enumerate(rows):
        yy = top - r * 7 * mm
        c.setFont(bold if r == 0 else reg, 10)
        for k, cell in enumerate(row):
            c.drawString(cols[k] + 2 * mm, yy - 5 * mm, cell)
    bottom = top - len(rows) * 7 * mm
    for r in range(len(rows) + 1):
        c.line(cols[0], top - r * 7 * mm, cols[-1], top - r * 7 * mm)
    for x in cols:
        c.line(x, top, x, bottom)
    _text_footer(c, w, 3, title)
    c.showPage()

    # лист 3 — A3 альбомный (опросный лист)
    w3, h3 = landscape(A3)
    c.setPageSize((w3, h3))
    _text_header(c, w3, h3)
    c.setFont(bold, 12)
    c.drawCentredString(w3 / 2, h3 - 32 * mm, "Опросный лист на расходомеры")
    cols = [25 * mm, 45 * mm, 200 * mm, 300 * mm, w3 - 15 * mm]
    head = ("№", "Параметр", "Ед. изм.", "Значение")
    data = [head] + [(str(i), f"Параметр среды {i}", "кгс/см²", f"{i * 1.5:.1f}") for i in range(1, 11)]
    top = h3 - 40 * mm
    for r, row in enumerate(data):
        c.setFont(bold if r == 0 else reg, 10)
        for k, cell in enumerate(row):
            c.drawString(cols[k] + 2 * mm, top - r * 7 * mm - 5 * mm, cell)
    for r in range(len(data) + 1):
        c.line(cols[0], top - r * 7 * mm, cols[-1], top - r * 7 * mm)
    for x in cols:
        c.line(x, top, x, top - len(data) * 7 * mm)
    _text_footer(c, w3, 4, title)
    c.showPage()
    c.save()
    return dst


# --------------------------------------------------------------------------- спецификация

def _side_stamp(c: canvas.Canvas, h: float) -> None:
    """Боковой штамп слева от рамки: «Взам. инв. №», «Подп. и дата», «Инв. № подл.» — вертикально."""
    reg, _ = _fonts()
    x0, x1, xm = 8 * mm, 20 * mm, 13 * mm
    cells = [(5 * mm, 30 * mm, "Инв. № подл."), (30 * mm, 65 * mm, "Подп. и дата"),
             (65 * mm, 90 * mm, "Взам. инв. №")]
    for y0, y1, label in cells:
        _box(c, x0, y0, x1, y1, 0.7)
        c.line(xm, y0, xm, y1)
        c.saveState()
        c.translate(x0 + 3.5 * mm, (y0 + y1) / 2)
        c.rotate(90)
        c.setFont(reg, 7)
        c.drawCentredString(0, 0, label)
        c.restoreState()


def _main_stamp(c: canvas.Canvas, w: float, sheet: int, sheets: int, big: bool) -> float:
    """Основная надпись 185 мм в правом нижнем углу рамки. Возвращает её верхнюю границу."""
    reg, bold = _fonts()
    x1, y0 = w - 5 * mm, 5 * mm
    x0 = x1 - 185 * mm
    hgt = 40 * mm if big else 15 * mm
    y1 = y0 + hgt
    c.setLineWidth(0.8)
    c.line(x0, y1, x1, y1)
    c.line(x0, y0, x0, y1)
    if big:
        for k in range(1, 8):
            c.setLineWidth(0.3)
            c.line(x0, y0 + k * 5 * mm, x0 + 65 * mm, y0 + k * 5 * mm)
        c.setLineWidth(0.8)
        for dx in (7, 17, 40, 55, 65):
            c.line(x0 + dx * mm, y0, x0 + dx * mm, y1)
        c.line(x0 + 65 * mm, y0 + 15 * mm, x1, y0 + 15 * mm)
        c.line(x0 + 65 * mm, y0 + 25 * mm, x1, y0 + 25 * mm)
        c.line(x1 - 50 * mm, y0, x1 - 50 * mm, y0 + 25 * mm)
        c.line(x1 - 35 * mm, y0 + 15 * mm, x1 - 35 * mm, y0 + 25 * mm)
        c.line(x1 - 20 * mm, y0 + 15 * mm, x1 - 20 * mm, y0 + 25 * mm)
        c.setFont(reg, 7)
        for k, lab in enumerate(["Разраб.", "Проверил", "Н. контр.", "ГИП"]):
            c.drawString(x0 + 1 * mm, y0 + (3 - k) * 5 * mm + 1.5 * mm, lab)
        c.drawString(x0 + 18 * mm, y0 + 3 * 5 * mm + 1.5 * mm, "Иванов")
        c.setFont(reg, 9)
        c.drawCentredString(x0 + 125 * mm, y0 + 31 * mm, SPEC_CODE)
        c.setFont(reg, 8)
        c.drawString(x0 + 67 * mm, y0 + 19 * mm, "Спецификация оборудования,")
        c.drawString(x0 + 67 * mm, y0 + 16 * mm, "изделий и материалов")
        c.setFont(reg, 7)
        c.drawCentredString(x1 - 42.5 * mm, y0 + 21.5 * mm, "Стадия")
        c.drawCentredString(x1 - 27.5 * mm, y0 + 21.5 * mm, "Лист")
        c.drawCentredString(x1 - 10 * mm, y0 + 21.5 * mm, "Листов")
        c.setFont(reg, 9)
        c.drawCentredString(x1 - 42.5 * mm, y0 + 16.5 * mm, "Р")
        c.drawCentredString(x1 - 27.5 * mm, y0 + 16.5 * mm, str(sheet))
        c.drawCentredString(x1 - 10 * mm, y0 + 16.5 * mm, str(sheets))
        _logo(c, x1 - 47 * mm, y0 + 3 * mm, 0.8)
    else:
        c.line(x1 - 10 * mm, y0, x1 - 10 * mm, y1)
        c.line(x1 - 10 * mm, y0 + 8 * mm, x1, y0 + 8 * mm)
        c.setFont(reg, 9)
        c.drawCentredString(x0 + 87 * mm, y0 + 6 * mm, SPEC_CODE)
        c.setFont(reg, 7)
        c.drawCentredString(x1 - 5 * mm, y0 + 10 * mm, "Лист")
        c.setFont(reg, 9)
        c.drawCentredString(x1 - 5 * mm, y0 + 2.5 * mm, str(sheet))
    return y1


SPEC_COLS = [("Поз.", 20), ("Наименование и техническая характеристика", 130), ("Тип, марка", 60),
             ("Код", 25), ("Завод-изготовитель", 45), ("Ед. изм.", 20), ("Кол.", 20), ("Масса", 25),
             ("Примечание", 40)]


def make_gost_spec() -> Path:
    reg, bold = _fonts()
    dst = OUT / "gost_spec.pdf"
    w, h = landscape(A3)
    c = canvas.Canvas(str(dst), pagesize=(w, h))
    items = [("1", "Датчик давления, 0–1,6 МПа, 4–20 мА", "Метран-150", "", "Метран", "шт.", "4", "", ""),
             ("2", "Расходомер электромагнитный DN50", "Взлёт ЭР", "", "Взлёт", "шт.", "2", "", ""),
             ("3", "Кабель контрольный 7×1,5", "КВВГнг(А)-LS", "", "Электрокабель", "м", "350", "", "")]
    for sheet in (1, 2):
        fx0, fy0, fx1, fy1 = 20 * mm, 5 * mm, w - 5 * mm, h - 5 * mm
        # рамка — одна ломаная с микроплаванием координат (выгрузка из САПР)
        _polyline(c, [(fx0, fy0), (fx1, fy0), (fx1, fy1), (fx0, fy1), (fx0, fy0)], 0.8, drift=0.03)
        _side_stamp(c, h)
        stamp_top = _main_stamp(c, w, sheet, 2, big=(sheet == 1))
        # таблица спецификации: от верха рамки до штампа, на всю ширину
        scale = (fx1 - fx0) / (sum(wd for _, wd in SPEC_COLS) * mm)
        xs = [fx0]
        for _, wd in SPEC_COLS:
            xs.append(xs[-1] + wd * mm * scale)
        head_bottom = fy1 - 15 * mm
        c.setLineWidth(0.8)
        c.line(fx0, head_bottom, fx1, head_bottom)
        for x in xs[1:-1]:
            c.line(x, fy1, x, stamp_top)
        c.setFont(reg, 8)
        for (name, _), xa, xb in zip(SPEC_COLS, xs, xs[1:]):
            c.drawCentredString((xa + xb) / 2, head_bottom + 6 * mm, name)
        c.setLineWidth(0.3)
        y = head_bottom
        rows = items if sheet == 1 else items[::-1]
        while y - 8 * mm > stamp_top:
            y -= 8 * mm
            c.line(fx0, y, fx1, y)
        for r, row in enumerate(rows):
            yy = head_bottom - (r + 1) * 8 * mm + 2.6 * mm
            for k, cell in enumerate(row):
                if cell:
                    c.drawString(xs[k] + 1.5 * mm, yy, cell)
        c.showPage()
    c.save()
    return dst


if __name__ == "__main__":
    print("ok", make_gost_text())
    print("ok", make_gost_spec())
