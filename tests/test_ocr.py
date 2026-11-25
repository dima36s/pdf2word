"""Распознавание сканов: оценки по пикселям и тексту без запуска Tesseract."""
from __future__ import annotations

import numpy as np
from PIL import Image, ImageDraw

from app import ocr
from app.ocr import Line, OcrResult, Paragraph, Word


def _w(text, x0=0.0, y0=0.0, x1=10.0, y1=10.0, conf=90.0):
    return Word(text, x0, y0, x1, y1, conf)


def test_homoglyphs_in_russian_line():
    words = [_w("OOO"), _w('"ЕвроХим-Усольский'), _w("Калийный"), _w('комбинат"'), _w("B"), _w("Cухаревой")]
    out = [w.text for w in ocr.fix_homoglyphs(words)]
    assert out[0] == "ООО" and out[4] == "В" and out[5] == "Сухаревой"
    # латиница в латинской строке и слова с «настоящей» латиницей не меняются
    assert [w.text for w in ocr.fix_homoglyphs([_w("ACME"), _w("Corp")])] == ["ACME", "Corp"]
    assert ocr.fix_homoglyphs([_w("давление"), _w("10"), _w("MPa")])[2].text == "MPa"


def test_size_from_letter_heights():
    # жирная строка: Tesseract дал x_size ~7 pt (28 px при 300 dpi), буквы — как у 10 pt
    em = 10 / 72 * 300                       # 41.7 px
    base = 200.0
    words = [Word("Местонахождение", 0, base - 0.67 * em, 300, base + 5, 96),
             Word("земельного", 310, base - 0.475 * em, 500, base + 9, 96),
             Word("участка", 510, base - 0.475 * em, 640, base + 9, 96)]
    ln = Line((0, base - 0.67 * em, 640, base + 9), base, 28.0, words)
    res = OcrResult([Paragraph([ln])], [])
    assert ocr.refine_sizes(res) == 1
    assert abs(ln.x_size / ocr.X_SIZE_TO_PT - em) < 1.0
    # совпадает с оценкой Tesseract — не трогаем
    ln2 = Line(ln.bbox, base, em * ocr.X_SIZE_TO_PT, words)
    assert ocr.refine_sizes(OcrResult([Paragraph([ln2])], [])) == 0


def _page(draw):
    img = Image.new("L", (1200, 600), 255)
    draw(ImageDraw.Draw(img))
    return img


def test_comb_boxes_found_and_frames_erased():
    def draw(d):
        # «[5][9]» — две клетки одной рамкой, «[2]» — одна, «[2][0][2][5]» — четыре
        for x0, n in ((100, 2), (300, 1), (500, 4)):
            d.rectangle((x0, 100, x0 + 50 * n, 150), outline=0, width=3)
            for k in range(1, n):
                d.line((x0 + 50 * k, 100, x0 + 50 * k, 150), fill=0, width=3)
            for k in range(n):
                d.rectangle((x0 + 50 * k + 18, 112, x0 + 50 * k + 32, 138), fill=0)   # «знак»
        d.line((100, 400, 1100, 400), fill=0, width=3)                                   # линия поля
    img = _page(draw)
    rules = ocr.detect_rules(img, 300)
    assert len(rules.combs) == 3
    assert rules.free and abs(rules.free[0][1] - 400) < 4 and not rules.h and not rules.v
    clean = np.asarray(ocr.erase_comb_frames(img, rules.combs)) < 140
    assert not clean[100:103, 100:200].any()                 # рамка стёрта
    assert clean[112:138, 118:132].any()                     # знак остался


def test_seal_is_ink_mark_but_colored_heading_is_text():
    img = Image.new("RGB", (2400, 1800), "white")
    d = ImageDraw.Draw(img)
    d.ellipse((200, 600, 650, 1050), outline=(40, 60, 200), width=8)                     # печать
    d.text((300, 800), "SEAL", fill=(40, 60, 200))
    d.rectangle((200, 200, 1400, 260), fill=(30, 80, 190))                               # цветной заголовок
    for k in range(8):                                                                   # чёрный текст
        d.rectangle((700, 300 + 120 * k, 2200, 330 + 120 * k), fill=(0, 0, 0))
    marks = ocr.find_ink_marks(img, 300)
    assert len(marks) == 1 and marks[0][0] <= 205 and marks[0][2] >= 645
    # цветом напечатан весь документ — это его текст, а не печати
    blue = Image.new("RGB", (2400, 1800), "white")
    db = ImageDraw.Draw(blue)
    for k in range(8):
        db.rectangle((200, 300 + 120 * k, 2200, 330 + 120 * k), fill=(30, 80, 190))
    db.ellipse((200, 1300, 650, 1750), outline=(40, 60, 200), width=8)
    assert ocr.find_ink_marks(blue, 300) == []


def test_lone_dash_in_table_cell_is_kept():
    v_rules = [(95, 0, 98, 100), (195, 0, 198, 100), (295, 0, 298, 100)]
    words = [_w("-", 140, 40, 150, 50, 40), _w("-", 240, 40, 250, 50, 40)]
    assert ocr._clean_word(words[0]) is None                          # без таблицы — артефакт
    assert ocr._lone_dash(words, 0, 40, v_rules) is words[0]
    assert ocr._lone_dash([_w("слово", 0, 40, 60, 50), _w("-", 70, 40, 80, 50, 40), _w("ещё", 90, 40, 130, 50)],
                          1, 40, []) is None                          # «тире» в широком пробеле строки


def test_page_marks_plan_reads_misread_number():
    def mark(i, text, nums, digits, x0=514):
        return ocr.PageMark(i, (x0, 33, x0 + 27, 44), text, nums, digits, 10.0, True)
    cands = {0: [mark(0, "- 19", [19], "19")], 1: [mark(1, "- 2/9 -", [2, 9], "29")],
             3: [mark(3, "- 4/9 -", [4, 9], "49")], 4: [mark(4, "- 5/9 -", [5, 9], "59", 748)]}
    width = {0: 595, 1: 595, 3: 595, 4: 842}
    plan = ocr.plan_page_marks(cands, width, {i: 842 for i in width})
    assert plan["form"] == "- #/# -" and plan["total"] == 9 and plan["shift"] == 0
    assert plan["found"][0] is not None and plan["align"] == "right"


def test_garbled_drawing_page_is_picture():
    words = [_w(t, conf=c) for t, c in [("PLZLOUZOZLE", 40), ("HOVU", 61), ("sanupoy", 30), ("wey", 20)] * 20]
    words += [_w("Условные", conf=90), _w("обозначения", conf=92)] * 3
    res = OcrResult([Paragraph([Line((0, 0, 10, 10), 10, 30, words)])], [])
    assert ocr.looks_like_picture(res, 3500, 4900)


def test_caption_row_splits_into_columns():
    # «(при наличии)   (подпись)   (расшифровка подписи)»: половина промежутков — разрывы колонок,
    # медиана с ними большая, и строка сходила за выровненную по ширине — подписи слипались
    from app.ocr import Word, _phrases

    k = 300 / 72
    def w(t, x0, x1):
        return Word(t, x0 * k, 552 * k, x1 * k, 563 * k, 90.0)
    words = [w("(при", 91, 107), w("наличии)", 111, 144), w("(подпись)", 253, 289),
             w("(расшифровка", 426, 480), w("подписи)", 483, 517)]
    parts = _phrases(words, 10 * k)
    assert [[x.text for x in p] for p in parts] == [["(при", "наличии)"], ["(подпись)"], ["(расшифровка", "подписи)"]]

