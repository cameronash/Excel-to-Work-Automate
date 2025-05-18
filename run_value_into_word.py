# run_value_into_word.py
from pathlib import Path
import argparse

from gslide.excel_reader import get_value
from gslide.word_writer import write_to_bookmark

# Indicate this is the updated script
print("🔄 running updated script…")

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy one Excel cell into a Word bookmark"
    )
    parser.add_argument("--excel",    required=True, help="Path to workbook (.xlsx)")
    parser.add_argument("--sheet",    required=True, help="Worksheet name")
    parser.add_argument("--cell",     required=True, help="Excel A1 cell (e.g. F22)")
    parser.add_argument("--word",     required=True, help="Path to Word doc (.docx/.docm)")
    parser.add_argument("--bookmark", required=True, help="Bookmark name in Word doc")
    args = parser.parse_args()

    # Make the paths absolute so Excel/Word can always find them
    args.excel = str(Path(args.excel).expanduser().resolve())
    args.word  = str(Path(args.word).expanduser().resolve())

    # Retrieve the value from Excel
    raw_value = get_value(args.excel, args.sheet, args.cell)

    # Format the value with comma as thousands separator
    try:
        formatted_value = f"{raw_value:,}"
    except Exception:
        formatted_value = str(raw_value)

    # Write the formatted value into the Word bookmark
    write_to_bookmark(args.word, args.bookmark, formatted_value)

    print(
        f"✅  Wrote {formatted_value!r} from {args.sheet}!{args.cell} "
        f"into bookmark “{args.bookmark}”."
    )


if __name__ == "__main__":
    main()
