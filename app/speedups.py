"""Ускорение pdf2docx на страницах с тысячами векторных элементов (каталоги из Word, сетки таблиц).

pdf2docx 0.5.13 группирует элементы попарным перебором — O(n²) вызовов медленных методов fitz.Rect.
На листе каталога ~1300 заливок-рамок ячеек, и одна страница разбиралась несколько секунд.
Здесь те же группировки считаются сортировкой с проходом (O(n log n)) и дают тот же результат:

* Collection.group_by_rows / group_by_columns — связные компоненты пересекающихся проекций;
* Shapes._merge_shapes — склейка соприкасающихся заливок одного цвета;
* Table._cells (python-docx) — сетка ячеек пересчитывалась при каждом table.cell(i, j);
  на время сборки таблицы pdf2docx она кэшируется и сбрасывается при объединении ячеек.

Чистые функции (interval_components, rect_components) покрыты тестами на совпадение с попарным
перебором. Если входные данные выходят за проверенные условия, вызывается исходный код pdf2docx.
"""
from __future__ import annotations

import logging
from contextvars import ContextVar

log = logging.getLogger(__name__)

ALIGN_EPS = 1e-3          # допуск pdf2docx в vertically_align_with / horizontally_align_with

# Подмена рамок элементов при группировке строк/колонок листа (см. vectorfix: таблица целиком).
# {id(элемент): (x0, y0, x1, y1)}; действует только внутри RawPage.parse_section.
GROUP_PROXY: ContextVar[dict | None] = ContextVar("p2w_group_proxy", default=None)


# --------------------------------------------------------------------------- чистая геометрия

def _ordered(parent: list[int], n: int) -> list[list[int]]:
    """Компоненты в порядке graph_bfs pdf2docx: по наименьшему индексу, внутри — по возрастанию."""
    def root(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    groups: dict[int, list[int]] = {}
    for i in range(n):
        groups.setdefault(root(i), []).append(i)
    return sorted(groups.values(), key=lambda g: g[0])


def interval_components(intervals: list[tuple[float, float] | None],
                        eps: float = ALIGN_EPS) -> list[list[int]]:
    """Связные компоненты отрезков, у которых перекрытие + eps >= 0 (factor=0 в pdf2docx).

    None — пустой элемент (pdf2docx не связывает его ни с кем). Длины должны быть неотрицательны.
    """
    n = len(intervals)
    parent = list(range(n))
    order = sorted((i for i in range(n) if intervals[i] is not None), key=lambda i: intervals[i][0])
    head, reach = None, 0.0
    for i in order:
        s, e = intervals[i]
        if head is not None and s <= reach + eps:
            parent[i] = head
            reach = max(reach, e)
        else:
            head, reach = i, e
    # сжатие путей не нужно: каждый указывает прямо на голову своей цепочки
    return _ordered(parent, n)


def rect_graph(rects: list[tuple[float, float, float, float]], keys: list,
               dist: float) -> list[set[int]]:
    """Списки смежности: одинаковый ключ и a ∩ (b, расширенный на dist) ≠ ∅, где a — элемент
    с меньшим номером (так вызывает Collection.group).

    Как fitz.Rect.intersects: строгие неравенства, пустой (вырожденный) a не пересекается ни с чем.
    Соседи добавляются по возрастанию — как в Collection.group, поэтому graph_bfs pdf2docx
    обходит граф в том же порядке, и порядок элементов в группах совпадает с исходным.
    """
    n = len(rects)
    neighbours: list[list[int]] = [[] for _ in range(n)]
    empty = [not (r[0] < r[2] and r[1] < r[3]) for r in rects]
    grown_empty = [not (r[2] - r[0] + 2 * dist > 0 and r[3] - r[1] + 2 * dist > 0) for r in rects]
    by_key: dict = {}
    for i, k in enumerate(keys):
        by_key.setdefault(k, []).append(i)
    for members in by_key.values():
        members.sort(key=lambda i: rects[i][0])
        active: list[int] = []
        for j in members:
            x0, y0, x1, y1 = rects[j]
            keep = []
            for i in active:
                a0, b0, a1, b1 = rects[i]
                if a1 + dist <= x0:        # дальше по x уже не дотянется ни до кого
                    continue
                keep.append(i)
                lo, hi = (i, j) if i < j else (j, i)
                if empty[lo] or grown_empty[hi]:
                    continue
                if a0 < x1 + dist and b0 < y1 + dist and y0 - dist < b1:
                    neighbours[i].append(j)
                    neighbours[j].append(i)
            keep.append(j)
            active = keep
    graph = []
    for items in neighbours:
        items.sort()
        graph.append(set(items))
    return graph


def overlap_pairs(rects: list[tuple[float, float, float, float]]) -> list[tuple[int, int]]:
    """Пары (i < j) с пересечением ненулевой площади (строгие неравенства по обеим осям)."""
    order = sorted(range(len(rects)), key=lambda i: rects[i][0])
    pairs: list[tuple[int, int]] = []
    active: list[int] = []
    for j in order:
        x0, y0, x1, y1 = rects[j]
        keep = []
        for i in active:
            a0, b0, a1, b1 = rects[i]
            if a1 <= x0:
                continue
            keep.append(i)
            if b0 < y1 and y0 < b1 and x0 < x1 and y0 < y1 and a0 < a1 and b0 < b1:
                pairs.append((i, j) if i < j else (j, i))
        keep.append(j)
        active = keep
    return pairs


def rect_components(rects: list[tuple[float, float, float, float]], keys: list,
                    dist: float) -> list[list[int]]:
    """Компоненты rect_graph в порядке graph_bfs pdf2docx (порядок внутри группы — как у set)."""
    from pdf2docx.common.algorithm import graph_bfs

    return [list(g) for g in graph_bfs(rect_graph(rects, keys, dist))]


# --------------------------------------------------------------------------- патчи

_installed = False


def install() -> None:
    global _installed
    if _installed:
        return
    _patch_grouping()
    _patch_merge_shapes()
    _patch_docx_cells()
    _patch_overlapped_lines()
    _patch_cell_assignment()
    _patch_fonts_cache()
    _installed = True


def _patch_fonts_cache() -> None:
    """Fonts.extract читает шрифты всего документа; разбор идёт пачками листов и отдельными
    проходами для колонтитулов — шрифты берутся один раз на документ."""
    from pdf2docx.font.Fonts import Fonts

    original = Fonts.extract.__func__

    def extract(cls, fitz_doc):
        try:
            cache = fitz_doc.__dict__.setdefault("_p2w_fonts", {})
        except (AttributeError, TypeError):
            return original(cls, fitz_doc)
        key = fitz_doc.page_count
        if key not in cache:
            cache[key] = original(cls, fitz_doc)
        return cache[key]

    Fonts.extract = classmethod(extract)


def _box(rect) -> tuple[float, float, float, float]:
    return float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)


def _patch_overlapped_lines() -> None:
    """Blocks._remove_overlapped_lines: get_main_bbox пары без пересечения сразу отвергает,
    поэтому проверять нужно только пересекающиеся строки (вместо всех пар)."""
    from pdf2docx.common.algorithm import graph_bfs
    from pdf2docx.layout.Blocks import Blocks

    original = Blocks._remove_overlapped_lines

    def remove_overlapped_lines(self, line_overlap_threshold: float):
        try:
            items = list(self._instances)
            n = len(items)
            neighbours: list[list[int]] = [[] for _ in range(n)]
            for i, j in overlap_pairs([_box(e.bbox) for e in items]):
                if items[i].get_main_bbox(items[j], threshold=line_overlap_threshold):
                    neighbours[i].append(j)
                    neighbours[j].append(i)
            graph = [set(sorted(nb)) for nb in neighbours]   # как в Collection.group: по возрастанию
            groups = [[items[i] for i in g] for g in graph_bfs(graph)]
        except Exception:
            log.debug("fast overlap check failed", exc_info=True)
            return original(self, line_overlap_threshold)
        for group in groups:
            if len(group) < 2:
                continue
            for line in sorted(group, key=lambda line: line.bbox.get_area())[:-1]:
                logging.warning('Ignore Line "%s" due to overlap', line.text)  # как в pdf2docx
                line.update_bbox((0, 0, 0, 0))
        return self

    Blocks._remove_overlapped_lines = remove_overlapped_lines


def _patch_cell_assignment() -> None:
    """TableBlock.assign_blocks: каждая ячейка перебирала все строки страницы (тысячи проверок
    fitz.Rect на таблицу). Ячейке достаются только строки, рамка которых её касается: остальные
    исходный код всё равно отвергает (нет ни вложения, ни пересечения)."""
    from pdf2docx.table.TableBlock import TableBlock

    original = TableBlock.assign_blocks

    def assign_blocks(self, blocks: list):
        try:
            blocks = list(blocks)
            boxes = [_box(b.bbox) for b in blocks]
        except Exception:
            return original(self, blocks)
        for row in self._rows:
            for cell in row:
                if not cell:
                    continue
                cx0, cy0, cx1, cy1 = _box(cell.bbox)
                cell.assign_blocks([b for b, (x0, y0, x1, y1) in zip(blocks, boxes)
                                    if x0 <= cx1 and cx0 <= x1 and y0 <= cy1 and cy0 <= y1])

    TableBlock.assign_blocks = assign_blocks


def _span(bbox, idx: int):
    s, e = bbox[idx], bbox[idx + 2]
    if s == e == 0 and bbox[1 - idx] == bbox[3 - idx] == 0:
        return None            # Rect(0, 0, 0, 0): pdf2docx считает элемент пустым
    return float(s), float(e)


def _patch_grouping() -> None:
    from pdf2docx.common.Collection import Collection

    original_rows = Collection.group_by_rows
    original_cols = Collection.group_by_columns

    def fast_groups(self, idx: int):
        spans = []
        proxy = GROUP_PROXY.get(None) or {}
        for inst in self._instances:
            b = proxy.get(id(inst))
            if b is None:
                b = inst.bbox
                b = (b.x0, b.y0, b.x1, b.y1)
            sp = _span(b, idx)
            if sp is not None and sp[1] < sp[0]:
                return None    # вывернутый прямоугольник — пусть решает исходный код
            spans.append(sp)
        comps = interval_components(spans)
        return [self.__class__([self._instances[i] for i in comp]) for comp in comps]

    def group_by_columns(self, factor: float = 0.0, sorted: bool = True, text_direction: bool = False):
        if factor == 0.0 and not text_direction:
            try:
                groups = fast_groups(self, 0)
            except Exception:
                log.debug("fast column grouping failed", exc_info=True)
                groups = None
            if groups is not None:
                if sorted:
                    groups.sort(key=lambda group: group.bbox[0])
                return groups
        return original_cols(self, factor=factor, sorted=sorted, text_direction=text_direction)

    def group_by_rows(self, factor: float = 0.0, sorted: bool = True, text_direction: bool = False):
        if factor == 0.0 and not text_direction:
            try:
                groups = fast_groups(self, 1)
            except Exception:
                log.debug("fast row grouping failed", exc_info=True)
                groups = None
            if groups is not None:
                if sorted:
                    groups.sort(key=lambda group: group.bbox[1])
                return groups
        return original_rows(self, factor=factor, sorted=sorted, text_direction=text_direction)

    Collection.group_by_columns = group_by_columns
    Collection.group_by_rows = group_by_rows


def _patch_merge_shapes() -> None:
    from pdf2docx.common import constants
    from pdf2docx.common.Collection import Collection
    from pdf2docx.common.share import RectType
    from pdf2docx.shape.Shapes import Shapes

    original = Shapes._merge_shapes

    def merge_shapes(shapes):
        try:
            normal = [s for s in shapes if not s.is_determined]
            rects = []
            for s in normal:
                b = s.bbox
                if b.is_infinite:
                    return original(shapes)
                rects.append((float(b.x0), float(b.y0), float(b.x1), float(b.y1)))
            comps = rect_components(rects, [s.color for s in normal], constants.TINY_DIST)
        except Exception:
            log.debug("fast shape merge failed", exc_info=True)
            return original(shapes)
        merged = []
        for comp in comps:
            group = Collection([normal[i] for i in comp])
            merged_area = group.bbox.get_area()
            sum_area = sum(shape.bbox.get_area() for shape in group)
            # нулевая площадь (линии тоньше 0,05 pt: рамка группы округляется до 0,1) — pdf2docx
            # здесь делил на ноль и ронял разбор всего документа; такие фигуры не склеиваем
            if merged_area and sum_area / merged_area >= constants.FACTOR_ALMOST:
                merged.append(group[0].update_bbox(group.bbox))
            else:
                merged.extend(group)
        merged.extend(s for s in shapes if s.equal_to_type(RectType.HYPERLINK))
        return merged

    Shapes._merge_shapes = staticmethod(merge_shapes)


def _patch_docx_cells() -> None:
    from docx.table import Table, _Cell
    from pdf2docx.table.TableBlock import TableBlock

    original_cells = Table._cells
    original_merge = _Cell.merge
    original_make = TableBlock.make_docx
    key = "_p2w_cells"

    def cells(self):
        cache = self.__dict__.get(key)
        if cache is None:
            return original_cells.fget(self)
        if not cache:
            cache.extend(original_cells.fget(self))
        return cache

    def merge(self, other_cell):
        try:
            return original_merge(self, other_cell)
        finally:
            cache = getattr(self._parent, "__dict__", {}).get(key)
            if cache:
                cache.clear()     # сетка изменилась — пересчитать при следующем обращении

    def make_docx(self, table):
        table.__dict__[key] = []
        try:
            return original_make(self, table)
        finally:
            table.__dict__.pop(key, None)

    Table._cells = property(cells)
    _Cell.merge = merge
    TableBlock.make_docx = make_docx
