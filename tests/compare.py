"""Визуальное сравнение: исходный PDF слева, DOCX (отрендеренный LibreOffice) справа.

python tests/compare.py source.pdf result.docx out.png [страниц]
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw


def docx_to_pdf(docx: Path, outdir: Path) -> Path:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(docx)],
                   check=True, capture_output=True, timeout=300)
    return outdir / (docx.stem + ".pdf")


def render(pdf: Path, n: int, dpi: int = 70) -> list[Image.Image]:
    doc = pymupdf.open(pdf)
    imgs = []
    for i in range(min(n, doc.page_count)):
        pix = doc[i].get_pixmap(dpi=dpi)
        imgs.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    return imgs


def main(src: str, docx: str, out: str, pages: int = 2) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        res_pdf = docx_to_pdf(Path(docx), Path(tmp))
        left, right = render(Path(src), pages), render(res_pdf, pages)
    rows = max(len(left), len(right))
    w = max(i.width for i in left + right)
    h = max(i.height for i in left + right)
    canvas = Image.new("RGB", (w * 2 + 30, (h + 30) * rows), "#888")
    dr = ImageDraw.Draw(canvas)
    for r in range(rows):
        y = r * (h + 30) + 25
        if r < len(left):
            canvas.paste(left[r], (0, y))
        if r < len(right):
            canvas.paste(right[r], (w + 30, y))
        dr.text((5, y - 20), f"PDF p.{r + 1}", fill="white")
        dr.text((w + 35, y - 20), f"DOCX p.{r + 1}", fill="white")
    canvas.save(out)
    print(out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]) if len(sys.argv) > 4 else 2)
