#!/usr/bin/env python3
"""Copy Excel ranges (tables) to bookmarks in a Word document."""

from __future__ import annotations

import argparse
import tempfile
from pathlib import Path

from gslide.excel_reader import _open_excel, _safe_close
from gslide.word_writer import _open_word, paste_image_at_bookmark


XL_PICTURE = -4147
XL_SCREEN = 1


def export_range(wb, sheet: str, cell_range: str) -> str:
    """Capture a range as an EMF file and return the path."""
    sht = wb.Worksheets(sheet)
    sht.Range(cell_range).CopyPicture(Appearance=XL_SCREEN, Format=XL_PICTURE)
    chart = wb.Charts.Add()
    chart.Paste()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".emf") as tmp:
        chart.Export(tmp.name)
    chart.Delete()
    return tmp.name


def main() -> None:
    parser = argparse.ArgumentParser(description="Copy tables from Excel to Word")
    parser.add_argument("--excel", required=True, help="Path to the Excel workbook")
    parser.add_argument("--word", required=True, help="Path to the Word document")
    parser.add_argument(
        "--mapping",
        action="append",
        nargs=3,
        metavar=("sheet", "range", "bookmark"),
        help="Sheet name, range, bookmark",
    )
    args = parser.parse_args()

    excel_path = str(Path(args.excel).expanduser().resolve())
    word_path = str(Path(args.word).expanduser().resolve())
    mappings = args.mapping or []

    xl, wb = _open_excel(excel_path)
    word_app, doc = _open_word(word_path)
    try:
        for sheet, rng, bookmark in mappings:
            img_path = export_range(wb, sheet, rng)
            try:
                paste_image_at_bookmark(word_path, bookmark, img_path, word_app=word_app, doc=doc)
            finally:
                Path(img_path).unlink(missing_ok=True)
        doc.Save()
    finally:
        doc.Close(False)
        word_app.Quit()
        _safe_close(xl, wb)


if __name__ == "__main__":
    main()
