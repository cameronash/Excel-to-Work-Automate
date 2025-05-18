# run_value_into_word.py
from pathlib import Path
import argparse

from gslide.excel_reader import get_value
from gslide.word_writer import write_to_bookmark


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

    # ── make the paths absolute so Excel/Word can always find them ─────────────
    args.excel = str(Path(args.excel).expanduser().resolve())
    args.word  = str(Path(args.word).expanduser().resolve())

    # ── do the work ────────────────────────────────────────────────────────────
    value = get_value(args.excel, args.sheet, args.cell)
    write_to_bookmark(args.word, args.bookmark, value)

    print(
        f"✅  Wrote {value!r} from {args.sheet}!{args.cell} "
        f"into bookmark “{args.bookmark}”."
    )


if __name__ == "__main__":
    main()
