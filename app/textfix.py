"""Исправления поверх pdf2docx.

1. Нормализация сырого текста страницы (до разбора макета):
   * склейка фрагментов одной строки, которые MuPDF разрезал из-за больших пробелов
     (выравнивание по ширине) — иначе pdf2docx принимает абзацы за таблицы;
   * склейка маркера списка (•, 1., а)) с текстом пункта;
   * пробел между соседними фрагментами строки, если в PDF его не было («1.Анализ» → «1. Анализ»);
   * символы из Private Use Area шрифтов Symbol/Wingdings → обычный Unicode (•, ▪, ✓ …).
2. Проверка «потоковых» (безрамочных) таблиц: если найденная таблица похожа на разорванный текст,
   страница разбирается повторно без поиска таких таблиц.
"""
from __future__ import annotations

import copy
import logging
import re

log = logging.getLogger(__name__)

# --------------------------------------------------------------------------- символы

_SYMBOL_PUA = {
    0xF0B7: "•", 0xF0A7: "▪", 0xF02D: "−", 0xF0B0: "°", 0xF0B1: "±", 0xF0B4: "×", 0xF0B8: "÷",
    0xF0A3: "≤", 0xF0B3: "≥", 0xF0B9: "≠", 0xF0BB: "≈", 0xF0AE: "→", 0xF0AC: "←", 0xF0DE: "⇒",
    0xF0D7: "⋅", 0xF0E0: "◊", 0xF06F: "○", 0xF0A8: "□",
}
_WINGDINGS_PUA = {
    0xF06C: "●", 0xF06E: "■", 0xF071: "❑", 0xF075: "◆", 0xF076: "❖", 0xF0A7: "▪", 0xF0A8: "◻",
    0xF0D8: "➢", 0xF0E8: "➔", 0xF0FC: "✓", 0xF0FB: "✗", 0xF0FE: "☑", 0xF09F: "•", 0xF0B7: "•",
    0xF0F0: "⇨", 0xF0E0: "➢", 0xF077: "⬥", 0xF0B2: "✧",
}

_LIST_MARKERS = set("•◦▪▫■□●○‣⁃–—-*·➢✓❖◆➔⇨❑►▸")
_LIST_MARKER_RE = re.compile(r"^(\(?\d{1,3}[.)]|\d{1,2}\.\d{1,2}\.?|[a-zа-яё][.)]|[IVXLC]{1,5}\.)$", re.I)


def _fix_symbol_chars(span: dict) -> None:
    font = (span.get("font") or "").lower()
    if "wingding" in font:
        table = _WINGDINGS_PUA
    elif "symbol" in font:
        table = _SYMBOL_PUA
    else:
        table = None
    changed = False
    for ch in span.get("chars", ()):
        c = ch.get("c", "")
        if len(c) == 1 and 0xE000 <= ord(c) <= 0xF8FF:
            repl = (table or {}).get(ord(c)) or _SYMBOL_PUA.get(ord(c)) or _WINGDINGS_PUA.get(ord(c))
            if repl:
                ch["c"] = repl
                changed = True
    if changed and table is not None:
        # шрифт-символ заменяем обычным, чтобы Word не показывал «квадратики»
        span["font"] = "Arial"


# --------------------------------------------------------------------------- строки

def _horizontal(line: dict) -> bool:
    d = line.get("dir", (1, 0))
    return abs(d[0] - 1) < 1e-3 and abs(d[1]) < 1e-3 and line.get("wmode", 0) == 0


def _chars(line: dict):
    for s in line["spans"]:
        yield from s.get("chars", ())


def _first_char(line: dict) -> dict | None:
    for s in line["spans"]:
        if s.get("chars"):
            return s["chars"][0]
    return None


def _last_char(line: dict) -> dict | None:
    for s in reversed(line["spans"]):
        if s.get("chars"):
            return s["chars"][-1]
    return None


def _line_size(line: dict) -> float:
    return max((s.get("size", 0) for s in line["spans"]), default=0) or 10.0


def _line_text(line: dict) -> str:
    return "".join(c.get("c", "") for c in _chars(line))


def _is_list_marker(text: str) -> bool:
    t = text.strip()
    return bool(t) and (t in _LIST_MARKERS or bool(_LIST_MARKER_RE.match(t)))


def _space_char(after: dict, gap: float, size: float, fill: bool = False) -> dict:
    """Синтетический пробел после символа after. fill=True — пробел закрывает весь промежуток,
    иначе pdf2docx разрежет строку по большому промежутку и пробел потеряется."""
    x0 = after["bbox"][2]
    w = max(0.1, gap if fill else min(gap, size * 0.25))
    return {"c": " ", "origin": (x0, after["origin"][1]),
            "bbox": (x0, after["bbox"][1], x0 + w, after["bbox"][3]), "synthetic": True}


def _union(a, b):
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _try_merge(a: dict, b: dict) -> bool:
    """Склеить строку b в конец строки a, если это продолжение той же строки текста."""
    if not (_horizontal(a) and _horizontal(b)):
        return False
    ca_last, cb_first = _last_char(a), _first_char(b)
    if ca_last is None or cb_first is None:
        return False
    size = max(_line_size(a), _line_size(b))
    if abs(ca_last["origin"][1] - cb_first["origin"][1]) > 0.2 * size:
        return False
    gap = b["bbox"][0] - a["bbox"][2]
    if gap < -0.5 or gap > 15 * size:
        return False
    has_space = ca_last["c"].isspace() or cb_first["c"].isspace()
    if not has_space:
        if not (_is_list_marker(_line_text(a)) and gap < 4 * size):
            return False
        a["spans"][-1]["chars"].append(_space_char(ca_last, gap, size, fill=True))
    a["spans"].extend(b["spans"])
    a["bbox"] = _union(a["bbox"], b["bbox"])
    return True


def _absorb_blank_spans(line: dict) -> None:
    """Фрагменты из одних пробелов pdf2docx отбрасывает — переносим их в предыдущий фрагмент."""
    out = []
    for span in line["spans"]:
        chars = span.get("chars") or []
        if out and chars and all(c.get("c", "").isspace() for c in chars):
            prev = out[-1]
            prev["chars"].extend({**c, "c": " "} for c in chars[:1])
            b = prev["bbox"]
            prev["bbox"] = (b[0], b[1], max(b[2], chars[-1]["bbox"][2]), b[3])
            continue
        out.append(span)
    line["spans"] = out


def _insert_missing_spaces(line: dict) -> None:
    spans = [s for s in line["spans"] if s.get("chars")]
    head = ""
    for prev, nxt in zip(spans, spans[1:]):
        head += "".join(c.get("c", "") for c in prev["chars"])
        c1, c2 = prev["chars"][-1], nxt["chars"][0]
        if c1["c"].isspace() or c2["c"].isspace():
            continue
        size = max(prev.get("size", 10), nxt.get("size", 10))
        gap = c2["bbox"][0] - c1["bbox"][2]
        if gap > 0.2 * size:
            marker = _is_list_marker(head) and gap < 4 * size
            prev["chars"].append(_space_char(c1, gap, size, fill=marker))


_NO_SPACE_AFTER = set("-\u00ad\u2010\u2011/")   # дефисы: перенос слова, пробел не нужен
_NEVER_WORD_START = set("ьъы")                 # с этих букв не начинается ни одно слово


def _is_cjk(c: str) -> bool:
    o = ord(c)
    return 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF


def _add_line_end_space(line: dict, nxt: dict) -> None:
    """pdf2docx склеивает строки абзаца без пробела, если строка в PDF не кончается пробелом."""
    last, first = _last_char(line), _first_char(nxt)
    if last is None or first is None:
        return
    c, n = last["c"], first["c"]
    if not c or c.isspace() or n.isspace() or c in _NO_SPACE_AFTER or _is_cjk(c) or _is_cjk(n[:1] or " "):
        return
    if n in _NEVER_WORD_START and c.isalpha():
        return  # слово разорвано без дефиса (узкая ячейка): «Удовлетворённост|ь»
    size = _line_size(line)
    if first["origin"][1] - last["origin"][1] < 0.5 * size:  # следующая строка не ниже — не перенос
        return
    line["spans"][-1]["chars"].append(_space_char(last, size * 0.25, size))


def normalize_text_blocks(blocks: list[dict]) -> list[dict]:
    for block in blocks:
        lines = block.get("lines")
        if not lines:
            continue
        for line in lines:
            for span in line["spans"]:
                _fix_symbol_chars(span)
        merged = [lines[0]]
        for ln in lines[1:]:
            if not _try_merge(merged[-1], ln):
                merged.append(ln)
        for ln in merged:
            _absorb_blank_spans(ln)
            _insert_missing_spaces(ln)
        for ln, nxt in zip(merged, merged[1:]):
            _add_line_end_space(ln, nxt)
        block["lines"] = merged
        bb = merged[0]["bbox"]
        for ln in merged[1:]:
            bb = _union(bb, ln["bbox"])
        block["bbox"] = bb
    return blocks


# --------------------------------------------------------------------------- патчи pdf2docx

_installed = False


def install_patches() -> None:
    """Подменяет извлечение текста в pdf2docx на версию с нормализацией (идемпотентно)."""
    global _installed
    if _installed:
        return
    from pdf2docx.page.RawPageFitz import RawPageFitz

    original = RawPageFitz._preprocess_text

    def _preprocess_text(self, **settings):
        blocks = original(self, **settings)
        try:
            return normalize_text_blocks(blocks)
        except Exception:  # нормализация не должна ломать конвертацию
            log.exception("text normalization failed")
            return blocks

    RawPageFitz._preprocess_text = _preprocess_text
    _patch_paragraph_split()
    _installed = True


def _row_starts_list_item(text: str) -> bool:
    text = text.lstrip()
    if not text:
        return False
    if text[0] in _LIST_MARKERS and (len(text) == 1 or text[1].isspace()):
        return True
    first = text.split(None, 1)[0]
    return bool(_LIST_MARKER_RE.match(first)) and len(text) > len(first)


def _patch_paragraph_split() -> None:
    """Включает в pdf2docx разбиение абзаца перед пунктом списка (в оригинале отключено: `if 0 and …`)."""
    from pdf2docx.common import constants
    from pdf2docx.text.Lines import Lines

    def split_vertically_by_text(self, line_break_free_space_ratio: float, new_paragraph_free_space_ratio: float):
        rows = self.group_by_physical_rows()
        for row in rows:
            row.sort_in_line_order()
        num = len(rows)
        if num == 1:
            return rows
        W = max(row[-1].bbox[2] - row[0].bbox[0] for row in rows[1:])
        H = sum(row[0].bbox[3] - row[0].bbox[1] for row in rows) / num
        res = []
        lines = Lines()
        punc = tuple(constants.SENTENCE_END_PUNC)
        start_of_para = end_of_para = False
        start_of_sen = end_of_sen = False
        for i, row in enumerate(rows):
            end_of_sen = row[-1].text.strip().endswith(punc)
            w = row[-1].bbox[2] - row[0].bbox[0]
            if i > 0 and _row_starts_list_item("".join(ln.text for ln in row)):
                start_of_para = True
            elif end_of_sen and w / W <= 1.0 - line_break_free_space_ratio:
                end_of_para = True
            elif start_of_sen and (W - w) / H >= new_paragraph_free_space_ratio:
                start_of_para = True
            if end_of_para:
                lines.extend(row)
                res.append(lines)
                lines = Lines()
            elif start_of_para:
                if lines:
                    res.append(lines)
                lines = Lines()
                lines.extend(row)
            else:
                lines.extend(row)
            start_of_sen = end_of_sen
            start_of_para = end_of_para = False
        if lines:
            res.append(lines)
        return res

    Lines.split_vertically_by_text = split_vertically_by_text

    # В оригинале пробел между соседними строками ставится только для ASCII (`c.encode().isalnum()`),
    # поэтому русские слова из разных блоков склеивались: «2026 г.Срок оплаты».
    from pdf2docx.text.TextSpan import TextSpan

    word_punct = set(".,;:!?)]}»”\"'%…№") | set("([{«„“")

    def wordish(c: str) -> bool:
        return bool(c) and not _is_cjk(c[-1]) and (c[-1].isalnum() or c[-1] in word_punct)

    def adjust_last_word(self, delete_end_line_hyphen: bool):
        for i, line in enumerate(self._instances[:-1]):
            end_span, start_span = line.spans[-1], self._instances[i + 1].spans[0]
            if not (isinstance(end_span, TextSpan) and isinstance(start_span, TextSpan)):
                continue
            if not end_span.chars or not start_span.chars:
                continue
            end_char, next_char = end_span.chars[-1], start_span.chars[0]
            if delete_end_line_hyphen and end_char.c.endswith("-") and next_char.c[:1].islower():
                end_char.c = ""
            nc = next_char.c[:1]
            if wordish(end_char.c) and wordish(nc) and not (nc in _NEVER_WORD_START and end_char.c[-1:].isalpha()):
                end_char.c += " "

    Lines.adjust_last_word = adjust_last_word


# --------------------------------------------------------------------------- ложные таблицы

def _iter_tables(blocks):
    for b in blocks:
        if getattr(b, "is_table_block", False):
            yield b
            for row in b:
                for cell in row:
                    if getattr(cell, "blocks", None):
                        yield from _iter_tables(cell.blocks)


def _page_tables(page):
    for section in page.sections:
        for column in section:
            yield from _iter_tables(column.blocks)


def _cell_text(cell) -> str:
    return " ".join(b.text for b in (cell.blocks or []) if hasattr(b, "text")).strip()


def _suspicious_stream_table(table) -> bool:
    """Безрамочная «таблица», которая на деле — разорванный по словам абзац."""
    rows = list(table)
    if not rows:
        return False
    total = placeholders = 0
    texts = []
    for row in rows:
        for cell in row:
            total += 1
            bb = cell.bbox
            if bb.is_empty or bb.width <= 0:
                placeholders += 1
            elif t := _cell_text(cell):
                texts.append(t)
    if len(rows) == 1:
        # одна строка из нескольких ячеек по 1–2 слова — это слова одной строки текста;
        # блок 1×1 (плашка с фоном) — нормальная таблица
        return len(texts) >= 2 and all(len(t.split()) <= 2 for t in texts)
    return total > 0 and placeholders / total > 0.4


def page_has_false_tables(page) -> bool:
    """Откат нужен, только если подозрительных безрамочных таблиц на странице большинство."""
    stream = [t for t in _page_tables(page) if t.is_stream_table_block]
    bad = sum(1 for t in stream if _suspicious_stream_table(t))
    return bad > 0 and bad >= 0.5 * len(stream)


def parse_page_safely(page, settings: dict) -> bool:
    """Разбирает страницу; при ложных безрамочных таблицах — повторно без них.
    Возвращает True, если понадобился повторный разбор."""
    if not settings.get("parse_stream_table", True):
        page.parse(**settings)
        return False
    try:
        # memo: ссылки секций на саму страницу должны остаться ссылками на неё, а не на копию
        backup = copy.deepcopy(page.__dict__, {id(page): page})
    except Exception:
        backup = None
    page.parse(**settings)
    if backup is None or not page_has_false_tables(page):
        return False
    page.__dict__.clear()
    page.__dict__.update(backup)
    page.parse(**{**settings, "parse_stream_table": False})
    return True
