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


def _photo_pdf(px_w=2436, px_h=1702, ppi=72):
    """Снимок с телефона в PDF: страница размером с картинку в пикселях (iOS Quartz)."""
    import io
    import pymupdf
    img = Image.new("RGB", (px_w, px_h), (240, 240, 235))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    doc = pymupdf.open()
    page = doc.new_page(width=px_w * 72 / ppi, height=px_h * 72 / ppi)
    page.insert_image(page.rect, stream=buf.getvalue())
    return doc


def test_phone_photo_page_is_a4():
    # 86 × 60 см «страница» из пикселей снимка — это лист A4: рендер в пикселях листа, а не 40 Мпикс
    doc = _photo_pdf()
    long_in, short_in = ocr.photo_page_inches(doc[0])
    assert abs(long_in - 842 / 72) < 0.01 and abs(short_in - long_in * 1702 / 2436) < 0.01
    img, dpi = ocr.render_page(doc[0], 300)
    assert max(img.size) <= 2436 * ocr.PHOTO_UPSCALE + 2 and 200 <= dpi <= 300
    # настоящий скан A4 (300 dpi) — не снимок
    assert ocr.photo_page_inches(_photo_pdf(2480, 3508, 300)[0]) is None
    # A3 и меньше по размеру страницы — не трогаем
    assert ocr.photo_page_inches(_photo_pdf(1191, 842)[0]) is None


def test_flatten_light_whitens_paper_keeps_ink():
    # тень от руки: фон от 160 до 250; текст и тонкая серая линия остаются тёмными, поле стола — белеет
    w, h = 1200, 800
    grad = np.tile(np.linspace(160, 250, w, dtype=np.float32), (h, 1))
    arr = np.stack([grad] * 3, axis=-1)
    arr[400:403, 100:1100] = (grad[400:403, 100:1100, None] * 0.55)        # линия таблицы
    arr[200:230, 300:320] = 40                                              # буква
    arr[:, :12] = (60, 120, 60)                                              # зелёный стол у края
    out = np.asarray(ocr.flatten_light(Image.fromarray(arr.astype(np.uint8)).convert("RGB")))
    assert out[100, 100:1100].mean() > 245 and out[700, 1000].mean() > 245
    assert out[401, 150].mean() < 170 and out[401, 1050].mean() < 170
    assert out[215, 310].mean() < 120
    assert out[:, :5].mean() > 250


def test_curved_rows_are_straightened():
    # изогнутый лист: горизонтали к правому краю уходят вверх по дуге (до 20 px)
    w, h = 2400, 1800
    img = Image.new("L", (w, h), 255)
    dr = ImageDraw.Draw(img)
    for y0 in range(300, 1500, 150):
        pts = [(x, y0 - 20 * ((x - 1200) / 1200) ** 2 * (1 if x > 1200 else 0.3)) for x in range(150, 2250, 10)]
        dr.line(pts, fill=60, width=4)
    rgb = img.convert("RGB")
    _, straight = ocr.straighten_rows(rgb, img, 300)
    arr = np.asarray(straight)
    for y0 in range(300, 1500, 300):
        ys = []
        for x in (300, 1200, 2100):
            col = arr[y0 - 40:y0 + 40, x]
            ys.append(int(col.argmin()))
        assert max(ys) - min(ys) <= 4, (y0, ys)


def test_comb_reading_votes_for_cell_count():
    # общий проход «00», отдельное чтение ошибается «60», двойное увеличение потеряло знак
    assert ocr.vote_comb((0, 0, 140, 70), ["00", "60", "0", "00", "0"]) == "00"
    # «строкой» путает последнюю цифру — берётся из основных: «0076»
    assert ocr.vote_comb((0, 0, 280, 70), ["", "0076", "0075", "0070", "0070"]) == "0076"
    assert ocr.vote_comb((0, 0, 140, 70), ["", "", "", "РФ", "И"]) == "РФ"


def test_comb_words_replaced_and_split_by_dashes():
    # общий проход прочёл «-59-2-00-» одним словом через три группы; второе слово — мусор от стенок
    rects = [(100, 0, 240, 70), (300, 0, 370, 70), (430, 0, 570, 70)]
    words = [Word("[519]", 100, 5, 240, 65, 40), Word("-2-0О-", 250, 5, 570, 65, 50)]
    ln = Line((100, 0, 570, 70), 60, 30, words)
    res = OcrResult([Paragraph([ln])], [])
    cand = {rects[0]: ["59", "59", "39", "59"], rects[1]: ["2", "", "2", "З"], rects[2]: ["00", "60", "0", "00"]}
    cand[(245, 30, 260, 34)] = "-"
    ocr.put_combs(res, cand)
    assert [w.text for w in ln.words] == ["59", "-", "2", "00"]


def test_unsure_words_under_seal_dropped():
    words = [Word("Pte,", 300, 1130, 340, 1150, 42), Word("Сухаревой", 1000, 1130, 1200, 1150, 92),
             Word("Дата", 320, 1300, 380, 1330, 88)]
    res = OcrResult([Paragraph([Line((0, 1100, 1300, 1350), 1150, 30, words)])], [])
    ocr.drop_unsure_in_marks(res, [(289, 921, 833, 1373)])
    assert [w.text for w in res.paragraphs[0].lines[0].words] == ["Сухаревой", "Дата"]

