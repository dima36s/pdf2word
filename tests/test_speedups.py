"""Быстрые группировки (app/speedups.py) дают то же, что попарный перебор pdf2docx."""
from __future__ import annotations

import random

import pytest

from app import speedups


def _components(n, connected, exact_order=False):
    """Эталон: попарный перебор + BFS, как Collection.group в pdf2docx."""
    from pdf2docx.common.algorithm import graph_bfs

    graph = [set() for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            if connected(i, j):
                graph[i].add(j)
                graph[j].add(i)
    return [list(g) if exact_order else sorted(g) for g in graph_bfs(graph)]


def _random_intervals(rng, n):
    out = []
    for _ in range(n):
        s = round(rng.uniform(0, 500), rng.choice((0, 1, 2)))
        out.append((s, s + round(rng.uniform(0, 40), rng.choice((0, 1, 2)))))
    return out


@pytest.mark.parametrize("seed", range(20))
def test_interval_components_match_bruteforce(seed):
    rng = random.Random(seed)
    spans = _random_intervals(rng, rng.randint(1, 120))
    spans[0] = None if seed % 3 == 0 else spans[0]   # пустой элемент ни с кем не связан

    def connected(i, j):
        if spans[i] is None or spans[j] is None:
            return False
        (s1, e1), (s2, e2) = spans[i], spans[j]
        return (e1 - s1) + (e2 - s2) - (max(e1, e2) - min(s1, s2)) + 1e-3 >= 0

    assert speedups.interval_components(spans) == _components(len(spans), connected)


def test_interval_components_touching_and_nested():
    spans = [(0, 10), (10, 20), (30, 40), (32, 35), (20.0005, 25)]
    assert speedups.interval_components(spans) == [[0, 1, 4], [2, 3]]


@pytest.mark.parametrize("seed", range(20))
def test_rect_components_match_fitz(seed):
    import pymupdf

    rng = random.Random(100 + seed)
    rects, keys = [], []
    for _ in range(rng.randint(1, 150)):
        x, y = rng.uniform(0, 300), rng.uniform(0, 300)
        if rng.random() < 0.5:   # тонкие «линии» сетки, как рамки ячеек в каталогах
            w, h = (rng.uniform(5, 80), 0.5) if rng.random() < 0.5 else (0.5, rng.uniform(5, 80))
        else:
            w, h = rng.uniform(1, 40), rng.uniform(1, 40)
        rects.append((x, y, x + w, y + h))
        keys.append(rng.choice((0, 0xFFFFFF, 0xD6D6D6)))
    d = 0.5

    def connected(i, j):
        a, b = pymupdf.Rect(rects[i]), pymupdf.Rect(rects[j])
        return keys[i] == keys[j] and a.intersects(b + (-d, -d, d, d))

    # порядок внутри групп тоже совпадает: от него зависит, какая заливка останется после склейки
    assert speedups.rect_components(rects, keys, d) == _components(len(rects), connected, True)


def test_pdf2docx_grouping_unchanged():
    """Патч Collection.group_by_rows/columns на реальных объектах pdf2docx."""
    from pdf2docx.common.Collection import Collection
    from pdf2docx.common.Element import Element

    from app.textfix import install_patches

    install_patches()
    rng = random.Random(7)
    items = []
    for _ in range(80):
        x, y = rng.uniform(0, 400), rng.uniform(0, 600)
        items.append(Element({"bbox": (x, y, x + rng.uniform(0, 60), y + rng.uniform(0, 15))}))
    col = Collection(items)

    index = {id(e): i for i, e in enumerate(items)}

    def ids(groups):
        # порядок внутри группы у pdf2docx — порядок обхода set(); сравниваем состав и порядок групп
        return [sorted(index[id(e)] for e in g) for g in groups]

    def brute(fun, key):
        groups = Collection(items).group(fun)
        groups.sort(key=key)
        return ids(groups)

    rows = brute(lambda a, b: a.horizontally_align_with(b), lambda g: g.bbox[1])
    cols = brute(lambda a, b: a.vertically_align_with(b), lambda g: g.bbox[0])
    assert ids(col.group_by_rows()) == rows
    assert ids(col.group_by_columns()) == cols


def test_docx_cell_cache_follows_merges():
    import docx

    from app.textfix import install_patches

    install_patches()
    doc = docx.Document()
    table = doc.add_table(rows=3, cols=3)
    table.__dict__["_p2w_cells"] = []
    first = table.cell(0, 0)
    assert table.cell(0, 0) is first                     # кэш работает
    first.merge(table.cell(1, 1))
    assert table.cell(1, 1)._tc is table.cell(0, 0)._tc  # после объединения сетка пересчитана
    table.__dict__.pop("_p2w_cells")
    assert table.cell(1, 1)._tc is table.cell(0, 0)._tc
    assert len(table._cells) == 9


@pytest.mark.parametrize("seed", range(10))
def test_overlap_pairs_match_fitz(seed):
    import pymupdf

    rng = random.Random(500 + seed)
    rects = []
    for _ in range(rng.randint(1, 120)):
        x, y = round(rng.uniform(0, 300), 1), round(rng.uniform(0, 300), 1)
        w, h = round(rng.choice((0, rng.uniform(0, 60))), 1), round(rng.choice((0, rng.uniform(0, 15))), 1)
        rects.append((x, y, x + w, y + h))
    rects.append((0, 0, 0, 0))   # строка, убранная pdf2docx (рамка обнулена)
    expect = [(i, j) for i in range(len(rects)) for j in range(i + 1, len(rects))
              if not (pymupdf.Rect(rects[i]) & pymupdf.Rect(rects[j])).is_empty]
    assert sorted(speedups.overlap_pairs(rects)) == expect

