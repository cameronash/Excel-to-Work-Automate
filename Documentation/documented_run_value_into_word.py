#!/usr/bin/env python3
"""
EXCEL-TO-WORD AUTOMATION SCRIPT
==============================
This script automates copying data from Excel spreadsheets into Word documents.
It reads numeric values from specified Excel cells and inserts them into 
Word document bookmarks, with options to format as numbers or convert to words.

MAIN PURPOSE: Replace manual copy-paste workflow with automated data transfer

WORKFLOW:
1. Read configuration (which Excel cells map to which Word bookmarks)  
2. Open Excel file once and extract all needed values
3. Open Word document once and insert all formatted values
4. Save and close both documents

PROGRAMMING CONCEPTS USED:
- Command-line argument parsing (argparse)
- File path handling (pathlib)
- Excel automation (via COM/win32com)
- Word automation (via COM/win32com)  
- Exception handling (try/except blocks)
- Data processing (pandas for config files)
"""

from pathlib import Path  # Modern way to handle file paths (better than os.path)
import argparse          # Built-in library for command-line interfaces

# Import our custom modules (these are in the gslide package)
from gslide.excel_reader import _open_excel, _safe_close  # Excel file operations
from gslide.word_writer import _open_word                 # Word file operations

# Third-party library for converting numbers to words (like 1000 → "One Thousand")
try:
    from num2words import num2words
except ImportError:
    # This is defensive programming - handle the case where required library isn't installed
    raise ImportError(
        "num2words library is required for writing numbers in words. "
        "Install with 'pip install num2words'."
    )

import pandas as pd  # Powerful library for working with Excel/CSV data

# User feedback - let them know the script is starting
print("🔄 running optimized script…")


def load_mappings_from_excel(config_path: str):
    """
    CONFIGURATION LOADER
    ===================
    Reads an Excel file that defines which Excel cells should be copied 
    to which Word bookmarks and how they should be formatted.
    
    Expected Excel columns:
    - Sheet Name: Which Excel worksheet (e.g., "Summary", "Data")
    - Cell: Which cell to read (e.g., "B5", "C10") 
    - Bookmark: Word bookmark name to write to (e.g., "total_revenue")
    - Formatting: How to format the number (optional, e.g., "{:,.2f}")
    
    Args:
        config_path (str): Path to the Excel configuration file
        
    Returns:
        list: List of tuples (sheet, cell, bookmark, format_spec)
        
    PROGRAMMING CONCEPTS:
    - pandas.read_excel(): Reads Excel files into a DataFrame (like a table)
    - .iterrows(): Loops through each row in the DataFrame
    - .get() method: Safely gets values from dictionary-like objects
    - .strip(): Removes whitespace from beginning/end of strings
    - List comprehension alternative: Building lists with loops
    """
    # Read the Excel config file into a pandas DataFrame (think: smart table)
    df = pd.read_excel(config_path, engine='openpyxl')
    
    mappings = []  # Empty list to store our configuration data
    
    # Loop through each row in the configuration spreadsheet
    for _, row in df.iterrows():  # _ means "ignore the index, just give me the row"
        # Extract and clean each column value
        # .get() returns None if column doesn't exist (safer than direct access)
        # str() converts to string, .strip() removes extra spaces
        sheet = str(row.get('Sheet Name') or '').strip()
        cell = str(row.get('Cell') or '').strip() 
        bookmark = str(row.get('Bookmark') or '').strip()
        
        # Handle optional formatting column
        fmt_spec = row.get('Formatting')
        fmt = None  # Default: no special formatting
        
        # Check if formatting was provided and isn't empty
        if pd.notna(fmt_spec) and fmt_spec:  # pd.notna() checks for not-null/not-NaN
            fmt = str(fmt_spec).strip()
            
        # Only add mapping if we have all required fields
        if sheet and cell and bookmark:  # All must be non-empty strings
            mappings.append((sheet, cell, bookmark, fmt))
            
    return mappings


def format_number_as_words(value):
    """
    NUMBER-TO-WORDS CONVERTER
    ========================
    Converts numeric values to written words, typically for financial documents.
    Example: 6800000 → "(Six Million Eight Hundred Thousand Dollars)"
    
    Args:
        value: The number to convert (can be int, float, or string)
        
    Returns:
        str: Formatted string with number written as words
        
    PROGRAMMING CONCEPTS:
    - Exception handling: try/except blocks to handle errors gracefully
    - Type conversion: int(float(value)) handles both "123" and "123.45"
    - String formatting: f-strings for clean string building
    - Third-party APIs: Using num2words library
    """
    # Handle missing or empty values
    if value is None or value == "":
        return "(Not Available)"
        
    try:
        # Convert to number: first to float (handles decimals), then int (whole numbers)
        num = int(float(value))
        
        # Use num2words library to convert number to English words
        # to='cardinal' means regular numbers (not ordinal like "first", "second") 
        # .title() capitalizes first letter of each word
        words = num2words(num, to='cardinal', lang='en').title() + ' Dollars'
        
        # Wrap in parentheses for document formatting
        return f"({words})"
        
    except Exception:
        # If conversion fails for any reason, just return the original value in parentheses
        # This is defensive programming - don't crash, just do something reasonable
        return f"({value})"


def format_number(value, fmt_spec: str | None):
    """
    NUMERIC FORMATTER
    ================
    Formats numbers for display in documents (adds commas, decimal places, etc.)
    
    Args:
        value: The number to format
        fmt_spec: Python format specification string (e.g., "{:,.2f}" for comma-separated with 2 decimals)
        
    Returns:
        str: Formatted number string
        
    PROGRAMMING CONCEPTS:
    - String formatting: Python's powerful format() method
    - Default parameters: fmt_spec can be None
    - Type annotations: str | None means "string or None type"
    - Default behavior: If no format specified, use comma grouping
    """
    # Handle empty/missing values by returning blank (not zero)
    if value is None or value == "":
        return ""
        
    try:
        num = float(value)  # Convert to decimal number
        
        if fmt_spec:
            # Use custom formatting if provided (e.g., "{:,.2f}".format(1234.5) → "1,234.50")
            return fmt_spec.format(num)
            
        # Default formatting: comma-separated with no decimals (e.g., 1234 → "1,234")
        return f"{num:,.0f}"  # f-string with format specification
        
    except Exception:
        # If number conversion fails, just return as string
        return str(value)


def main() -> None:
    """
    MAIN PROGRAM LOGIC
    ==================
    This is the primary function that coordinates the entire process.
    It handles command-line arguments, manages file operations, and 
    orchestrates the data transfer from Excel to Word.
    
    WORKFLOW:
    1. Parse command-line arguments (what files to use, how to configure)
    2. Load configuration (which cells map to which bookmarks)
    3. Extract all data from Excel (open once, read all values, close)
    4. Insert all data into Word (open once, write all values, save, close)
    
    PROGRAMMING CONCEPTS:
    - Command-line interfaces: Using argparse for professional CLI tools
    - Context managers: try/finally blocks for reliable cleanup
    - Batch processing: Open files once, do all operations, then close
    - Error handling: Graceful handling of missing cells/bookmarks
    """
    
    # STEP 1: SET UP COMMAND-LINE INTERFACE
    # ====================================
    # argparse makes our script professional - users can run it with different options
    parser = argparse.ArgumentParser(
        description="Copy multiple Excel cells into Word bookmarks (optimized)"
    )
    
    # Create mutually exclusive options (user must choose one or the other)
    group = parser.add_mutually_exclusive_group(required=True)
    
    # Option 1: Use Excel config file to define mappings
    group.add_argument('--config', help='Excel config file listing mappings')
    
    # Option 2: Define mappings directly on command line
    group.add_argument(
        '--mapping', action='append', nargs='+',
        help="sheet cell bookmark [format]"  # e.g., --mapping Sheet1 B5 revenue_bookmark
    )
    
    # Required arguments (user must always provide these)
    parser.add_argument('--excel', required=True, help='Path to Excel workbook')
    parser.add_argument('--word', required=True, help='Path to Word document')
    
    # Parse the actual command-line arguments the user provided
    args = parser.parse_args()

    # STEP 2: RESOLVE FILE PATHS
    # =========================
    # Convert user-provided paths to absolute paths (handles ~, .., etc.)
    args.excel = str(Path(args.excel).expanduser().resolve())
    args.word = str(Path(args.word).expanduser().resolve())

    # STEP 3: LOAD CONFIGURATION
    # =========================
    # Figure out which Excel cells should go to which Word bookmarks
    if args.config:
        # Load from Excel configuration file
        mappings = load_mappings_from_excel(args.config)
        print(f"🔢 Loaded {len(mappings)} mappings from '{args.config}'")
    else:
        # Build from command-line arguments
        mappings = []
        for m in args.mapping:
            # Each mapping should have 3 or 4 parts: sheet, cell, bookmark, [format]
            if len(m) not in (3, 4):
                parser.error(f"Invalid mapping {m}")
            mappings.append((m[0], m[1], m[2], m[3] if len(m) == 4 else None))

    # STEP 4: EXTRACT ALL DATA FROM EXCEL
    # ===================================
    # OPTIMIZATION: Open Excel once, read all values, then close
    # This is much faster than opening/closing for each cell
    
    xl, wb = _open_excel(args.excel)  # xl = Excel application, wb = workbook
    try:
        # Dictionary to store all the values we read from Excel
        # Key: (sheet_name, cell_address), Value: the actual cell value
        raw_values: dict[tuple[str, str], object] = {}
        
        # Loop through all our mappings and extract the values
        for sheet, cell, bookmark, fmt in mappings:
            try:
                # Access: Workbook → Worksheet → Cell → Value
                raw_values[(sheet, cell)] = wb.Worksheets(sheet).Range(cell).Value
            except Exception:
                # If we can't read a cell (sheet doesn't exist, cell is invalid, etc.)
                raw_values[(sheet, cell)] = None
                
    finally:
        # CRITICAL: Always close Excel, even if something goes wrong
        # This prevents Excel processes from staying open in background
        _safe_close(xl, wb)

    # STEP 5: INSERT ALL DATA INTO WORD
    # ================================
    # OPTIMIZATION: Open Word once, write all values, then save and close
    
    word_app, doc = _open_word(args.word)
    try:
        # Process each mapping and write to Word
        for sheet, cell, bookmark, fmt in mappings:
            # Get the raw value we extracted from Excel
            raw = raw_values.get((sheet, cell))
            
            # BUSINESS LOGIC: Decide how to format based on whether format is specified
            if fmt:
                # Explicit formatting means treat as numeric field
                formatted = format_number(raw, fmt)
                fmt_type = "numeric"
            else:
                # No formatting means convert to words (for financial documents)
                formatted = format_number_as_words(raw)
                fmt_type = "text"

            # Write the formatted value to the Word bookmark
            try:
                # Word bookmarks work by: find bookmark → get its range → set text → recreate bookmark
                rng = doc.Bookmarks(bookmark).Range  # Get the bookmark's location
                rng.Text = formatted                 # Replace text at that location
                doc.Bookmarks.Add(bookmark, rng)     # Recreate bookmark (Word deletes it when you change text)
                
                # User feedback
                print(f"✅ Wrote {fmt_type} '{formatted}' into '{bookmark}'")
                
            except Exception as e:
                # Handle individual bookmark errors without crashing entire script
                print(f"❌ Error writing to bookmark '{bookmark}': {e}")

        # Save all changes to the Word document
        doc.Save()
        
    finally:
        # CRITICAL: Always close Word properly
        doc.Close(False)  # Close document without saving again
        word_app.Quit()   # Close Word application


# SCRIPT ENTRY POINT
# ==================
# This is Python's way of saying "if this file is run directly (not imported), run main()"
if __name__ == '__main__':
    main()
