# run_value_into_word.py
from pathlib import Path
import argparse

from gslide.excel_reader import _open_excel, _safe_close
from gslide.word_writer import _open_word

# number-to-words conversion for text formatting
try:
    from num2words import num2words
except ImportError:
    raise ImportError("num2words library is required for writing numbers in words. Install with 'pip install num2words'.")

print("🔄 running optimized script…")

def load_mappings_from_excel(config_path: str):
    """
    Read an Excel config file with columns:
      Sheet Name | Cell | Bookmark | Formatting
    Returns list of tuples: (sheet, cell, bookmark, fmt_spec)
    
    If Formatting is blank, the value will be converted to text (words).
    """
    import pandas as pd
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
    """Convert a number to words with proper formatting."""
    if value is None or value == "":
        return "(Not Available)"
    
    try:
        num = float(value)
        words = num2words(int(num), to='cardinal', lang='en').title() + ' Dollars'
        
        # Add commas for better readability in large numbers
        if num >= 1000000:
            words = words.replace(" Thousand", ", Thousand")
            words = words.replace(" Million", ", Million")
            words = words.replace(" Billion", ", Billion")
            # Clean up any double commas
            words = words.replace(", ,", ",")
            words = words.replace("  ", " ")
        
        # Wrap in parentheses
        return f"({words})"
    except:
        return f"({str(value)})"


def format_number(value, fmt_spec):
    """Format a number according to the provided format specification."""
    if value is None or value == "":
        return "None"  # or use another placeholder of your choice
    
    try:
        num = float(value)
        return fmt_spec.format(num) if fmt_spec else f"{num:,.0f}"
    except:
        return str(value)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy multiple Excel cells into Word bookmarks (optimized)"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--config', help='Excel config file listing mappings'
    )
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
            mappings.append((m[0], m[1], m[2], m[3] if len(m)==4 else None))

    # --- open Excel once ---
    xl, wb = _open_excel(args.excel)
    try:
        raw_values = {}
        for sheet, cell, bookmark, fmt in mappings:
            try:
                val = wb.Worksheets(sheet).Range(cell).Value
            except Exception:
                val = None
            raw_values[(sheet, cell)] = val
    finally:
        _safe_close(xl, wb)

    # --- open Word once ---
    word_app, doc = _open_word(args.word)
    try:
        for sheet, cell, bookmark, fmt in mappings:
            raw = raw_values.get((sheet, cell))
            
            # Format based on whether formatting is specified
            if fmt is None or fmt == "":
                # Handle as text (words)
                formatted = format_number_as_words(raw)
                format_type = "text"
            else:
                # Handle as numeric with the specified format
                formatted = format_number(raw, fmt)
                format_type = "numeric"
            
            # Write to bookmark
            try:
                rng = doc.Bookmarks(bookmark).Range
                rng.Text = formatted
                doc.Bookmarks.Add(bookmark, rng)
                print(f"✅ Wrote {format_type} '{formatted}' into '{bookmark}'")
            except Exception as e:
                print(f"❌ Error writing to bookmark '{bookmark}': {e}")

        doc.Save()
    finally:
        doc.Close(False)
        word_app.Quit()

if __name__ == '__main__':
    main()