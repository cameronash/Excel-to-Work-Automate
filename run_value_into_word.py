# run_value_into_word.py
from pathlib import Path
import argparse

from gslide.excel_reader import get_value
from gslide.word_writer import write_to_bookmark

# number-to-words conversion for MarketRent1_text
try:
    from num2words import num2words
except ImportError:
    raise ImportError("num2words library is required for writing numbers in words. Install with 'pip install num2words'.")

print("🔄 running updated script…")

def load_mappings_from_excel(config_path: str):
    """
    Read an Excel config file with columns:
      Sheet Name | Cell | Bookmark | Formatting
    Returns list of tuples: (sheet, cell, bookmark, fmt_spec)
    """
    try:
        import pandas as pd
    except ImportError:
        raise ImportError("pandas is required to load mappings from Excel. Install with 'pip install pandas'.")

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


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy multiple Excel cells into Word bookmarks"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--config',
        help='Path to Excel file listing mappings (Sheet Name, Cell, Bookmark, Formatting)'
    )
    group.add_argument(
        '--mapping',
        action='append',
        nargs='+',
        help=(
            "One or more manual mappings; each is: sheet cell bookmark [format]\n"
            "e.g. --mapping 'PRP ValSum' F22 MarketRent1 '{:,.0f}'"
        )
    )
    parser.add_argument(
        '--excel', required=True,
        help='Path to the Excel workbook (.xlsx)'
    )
    parser.add_argument(
        '--word', required=True,
        help='Path to the Word document (.docx/.docm)'
    )
    args = parser.parse_args()

    # Make paths absolute
    args.excel = str(Path(args.excel).expanduser().resolve())
    args.word = str(Path(args.word).expanduser().resolve())

    # Determine mappings from config or manual
    if getattr(args, 'config', None):
        cfg_path = str(Path(args.config).expanduser().resolve())
        mappings = load_mappings_from_excel(cfg_path)
        print(f"🔢 Loaded {len(mappings)} mappings from config '{cfg_path}'")
    else:
        mappings = []
        for mapping in args.mapping:
            if len(mapping) not in (3, 4):
                parser.error(f"Invalid mapping: {mapping}. Must be 3 or 4 parts.")
            sheet, cell, bookmark = mapping[:3]
            fmt = mapping[3] if len(mapping) == 4 else None
            mappings.append((sheet, cell, bookmark, fmt))

    # Process each mapping (skip any *_text bookmarks to avoid overwriting special text)
    for sheet, cell, bookmark, fmt_spec in mappings:
        if bookmark.endswith('_text'):
            continue

        raw_value = get_value(args.excel, sheet, cell)
        try:
            num = float(raw_value)
            formatted = fmt_spec.format(num) if fmt_spec else f"{num:,.0f}"
        except Exception:
            formatted = str(raw_value)

        write_to_bookmark(args.word, bookmark, formatted)
        print(f"✅ Wrote numeric {formatted!r} into bookmark '{bookmark}'")

        # special number-to-words for the main market rent
        if bookmark == 'MarketRent1':
            try:
                num_int = int(float(raw_value))
                words = num2words(num_int, to='cardinal', lang='en').title() + ' Dollars'
            except Exception:
                words = str(raw_value)
            text_value = f"({words})"
            write_to_bookmark(args.word, f"{bookmark}_text", text_value)
            print(f"✅ Wrote text {text_value!r} into bookmark '{bookmark}_text'")


if __name__ == '__main__':
    main()
