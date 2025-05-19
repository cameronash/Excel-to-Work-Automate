#!/usr/bin/env python3
"""
Optimized script to copy multiple Excel cells into Word bookmarks,
with separate handling for numeric and text (words) fields.
"""
from pathlib import Path
import argparse

from gslide.excel_reader import _open_excel, _safe_close
from gslide.word_writer import _open_word

# number-to-words conversion
try:
    from num2words import num2words
except ImportError:
    raise ImportError(
        "num2words library is required for writing numbers in words. "
        "Install with 'pip install num2words'."
    )

import pandas as pd

print("🔄 running optimized script…")


def load_mappings_from_excel(config_path: str):
    """
    Read an Excel config file with columns:
      Sheet Name | Cell | Bookmark | Formatting
    Returns list of tuples: (sheet, cell, bookmark, fmt_spec)
    """
    df = pd.read_excel(config_path, engine='openpyxl')
    mappings = []
    for _, row in df.iterrows():
        sheet = str(row.get('Sheet Name') or '').strip()
        cell = str(row.get('Cell') or '').strip()
        bookmark = str(row.get('Bookmark') or '').strip()
        fmt_spec = row.get('Formatting')
        fmt = None
        if pd.notna(fmt_spec) and fmt_spec:
            fmt = str(fmt_spec).strip()
        if sheet and cell and bookmark:
            mappings.append((sheet, cell, bookmark, fmt))
    return mappings


def format_number_as_words(value):
    """Convert a number to words, e.g. 6800000 → (Six Million Eight Hundred Thousand Dollars)."""
    if value is None or value == "":
        return "(Not Available)"
    try:
        num = int(float(value))
        words = num2words(num, to='cardinal', lang='en').title() + ' Dollars'
        return f"({words})"
    except Exception:
        return f"({value})"


def format_number(value, fmt_spec: str | None):
    """Format a number according to the provided format spec or default comma-grouping."""
    if value is None or value == "":
        return ""  # return blank for missing values
    try:
        num = float(value)
        if fmt_spec:
            return fmt_spec.format(num)
        # default numeric formatting
        return f"{num:,.0f}"
    except Exception:
        return str(value)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy multiple Excel cells into Word bookmarks (optimized)"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--config', help='Excel config file listing mappings')
    group.add_argument(
        '--mapping', action='append', nargs='+',
        help="sheet cell bookmark [format]"
    )
    parser.add_argument('--excel', required=True, help='Path to Excel workbook')
    parser.add_argument('--word', required=True, help='Path to Word document')
    args = parser.parse_args()

    # resolve paths
    args.excel = str(Path(args.excel).expanduser().resolve())
    args.word = str(Path(args.word).expanduser().resolve())

    # load mappings
    if args.config:
        mappings = load_mappings_from_excel(args.config)
        print(f"🔢 Loaded {len(mappings)} mappings from '{args.config}'")
    else:
        mappings = []
        for m in args.mapping:
            if len(m) not in (3, 4):
                parser.error(f"Invalid mapping {m}")
            mappings.append((m[0], m[1], m[2], m[3] if len(m) == 4 else None))

    # --- open Excel once ---
    xl, wb = _open_excel(args.excel)
    try:
        raw_values: dict[tuple[str, str], object] = {}
        for sheet, cell, bookmark, fmt in mappings:
            try:
                raw_values[(sheet, cell)] = wb.Worksheets(sheet).Range(cell).Value
            except Exception:
                raw_values[(sheet, cell)] = None
    finally:
        _safe_close(xl, wb)

    # --- open Word once ---
    word_app, doc = _open_word(args.word)
    try:
        for sheet, cell, bookmark, fmt in mappings:
            raw = raw_values.get((sheet, cell))
            # decide formatting based on fmt_spec
            if fmt:
                # numeric field (explicit or default)
                formatted = format_number(raw, fmt)
                fmt_type = "numeric"
            else:
                # text field (convert to words)
                formatted = format_number_as_words(raw)
                fmt_type = "text"

            # write to bookmark
            try:
                rng = doc.Bookmarks(bookmark).Range
                rng.Text = formatted
                doc.Bookmarks.Add(bookmark, rng)
                print(f"✅ Wrote {fmt_type} '{formatted}' into '{bookmark}'")
            except Exception as e:
                print(f"❌ Error writing to bookmark '{bookmark}': {e}")

        doc.Save()
    finally:
        doc.Close(False)
        word_app.Quit()


if __name__ == '__main__':
    main()
