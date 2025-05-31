# still in the repo root
type nul > README.md           # creates an empty file (Windows PowerShell)
# or: echo > README.md
code README.md                 # opens it in VS Code

# G-Slide: Excel-to-Word Automation Tool

**Automate copying data from Excel spreadsheets into Word document bookmarks**

G-Slide eliminates manual copy-paste workflows by automatically extracting values from Excel cells and inserting them into predefined locations (bookmarks) in Word documents. Perfect for financial reports, regular document updates, and template population.

## 🚀 What It Does

- **Reads data** from specific Excel cells across multiple worksheets
- **Formats numbers** as currency, percentages, or converts to written words
- **Writes to Word bookmarks** automatically with proper formatting
- **Handles multiple mappings** in a single operation
- **Provides both GUI and command-line interfaces**

### Example Use Case
You have a monthly financial report template in Word with bookmarks like `total_revenue`, `profit_margin`, etc. G-Slide reads the latest numbers from your Excel budget file and automatically populates all the bookmarks in your Word report.

## 📋 Quick Start

### For Non-Technical Users (GUI)

1. **Run the application**: Double-click `gui_tk.exe` (if available) or run `python gui_tk.py`
2. **Select your Excel file** containing the source data
3. **Select your Word template** with bookmarks already set up
4. **Click Run** and watch your document get populated automatically!

### For Technical Users (Command Line)

```bash
# Using configuration file (recommended)
python run_value_into_word.py --config "config/g-slide mapping.xlsx" --excel "data.xlsx" --word "report.docx"

# Direct mapping specification
python run_value_into_word.py --mapping "Sheet1" "B5" "revenue_bookmark" "{:,.0f}" --excel "data.xlsx" --word "report.docx"
```

## 🛠️ Installation & Setup

### Prerequisites
- **Windows** (required for Microsoft Office COM automation)
- **Python 3.8+**
- **Microsoft Excel** installed and licensed
- **Microsoft Word** installed and licensed

### Install Dependencies
```bash
pip install pandas openpyxl num2words pywin32 pillow
```

### Project Structure
```
G-Slide/
├── run_value_into_word.py      # Main CLI script
├── gui_tk.py                   # User-friendly GUI application
├── src/gslide/
│   ├── excel_reader.py         # Excel automation functions
│   ├── word_writer.py          # Word automation functions
│   └── __init__.py            # Package marker
├── config/
│   └── g-slide mapping.xlsx    # Configuration file
├── tests/
│   ├── assets/
│   │   ├── Sample.xlsx         # Example Excel file
│   │   └── Template.docm       # Example Word template
│   ├── test_excel_reader.py    # Excel function tests
│   ├── test_integration.py     # End-to-end tests
│   └── test_range_image.py     # Image copying tests
└── README.md                   # This documentation
```

## ⚙️ Configuration

### Setting Up Excel-to-Word Mappings

Create an Excel configuration file (`config/g-slide mapping.xlsx`) with these columns:

| Sheet Name | Cell | Bookmark | Formatting |
|------------|------|----------|------------|
| Summary | B10 | total_revenue | `{:,.0f}` |
| Summary | C15 | profit_margin | `{:.1%}` |
| Summary | D20 | growth_words | *(empty)* |

**Column Explanations:**
- **Sheet Name**: Excel worksheet name (e.g., "Summary", "Data")
- **Cell**: Cell address to read from (e.g., "B10", "C15")
- **Bookmark**: Word bookmark name to write to
- **Formatting**: 
  - `{:,.0f}` = Numbers with commas (1500000 → "1,500,000")
  - `{:.1%}` = Percentage with 1 decimal (0.15 → "15.0%")
  - *(empty)* = Convert to words (1500000 → "(One Million Five Hundred Thousand Dollars)")

### Setting Up Word Bookmarks

1. **Open your Word template**
2. **Place cursor** where you want data inserted
3. **Insert → Bookmark** (or Ctrl+Shift+F5)
4. **Name the bookmark** (e.g., "total_revenue")
5. **Click Add**

The bookmark name must match exactly what's in your configuration file.

## 📖 Usage Examples

### GUI Application
```bash
python gui_tk.py
```
- Browse for Excel and Word files
- Click Run
- Progress bar shows operation status
- Success/error messages appear in popup dialogs

### Command Line Examples

**Using configuration file:**
```bash
python run_value_into_word.py \
  --config "config/g-slide mapping.xlsx" \
  --excel "monthly_data.xlsx" \
  --word "report_template.docx"
```

**Direct mapping (single bookmark):**
```bash
python run_value_into_word.py \
  --mapping "Summary" "B10" "revenue" "{:,.0f}" \
  --excel "data.xlsx" \
  --word "report.docx"
```

**Multiple direct mappings:**
```bash
python run_value_into_word.py \
  --mapping "Summary" "B10" "revenue" "{:,.0f}" \
  --mapping "Summary" "C15" "profit" "{:.1%}" \
  --excel "data.xlsx" \
  --word "report.docx"
```

## 🔧 Advanced Features

### Number Formatting Options

**Comma-separated numbers:**
```python
Formatting: "{:,.0f}"
1500000 → "1,500,000"
```

**Currency with decimals:**
```python
Formatting: "${:,.2f}"
1500000 → "$1,500,000.00"
```

**Percentages:**
```python
Formatting: "{:.1%}"
0.15 → "15.0%"
```

**Words conversion (empty formatting):**
```python
Formatting: (leave empty)
1500000 → "(One Million Five Hundred Thousand Dollars)"
```

### Batch Processing
The tool is optimized for efficiency:
- Opens Excel **once** and reads all required cells
- Opens Word **once** and writes all bookmarks
- Much faster than opening/closing files repeatedly

### Error Handling
- **Missing Excel cells**: Logs warning, continues with other mappings
- **Missing Word bookmarks**: Logs warning, continues with other mappings  
- **File access errors**: Clear error messages with suggestions
- **COM errors**: Graceful handling of Office application issues

## 🧪 Testing

Run the test suite to verify everything works:

```bash
# Test Excel reading functions
python -m pytest tests/test_excel_reader.py -v

# Test end-to-end workflow  
python -m pytest tests/test_integration.py -v

# Test image copying (advanced feature)
python -m pytest tests/test_range_image.py -v

# Run all tests
python -m pytest tests/ -v
```

## 🐛 Troubleshooting

### Common Issues

**"Excel file not found"**
- Check file path is correct
- Use absolute paths if relative paths don't work
- Ensure file isn't open in Excel

**"Bookmark not found"**
- Verify bookmark names match exactly (case-sensitive)
- Check bookmarks exist in Word template
- Bookmark names can't contain spaces or special characters

**"COM error" or "Excel/Word won't start"**
- Ensure Office applications are properly installed
- Close any existing Excel/Word processes
- Run as administrator if permission issues occur
- Restart computer if COM registration is corrupted

**GUI freezes or becomes unresponsive**
- Large files can take time to process
- Look for console output to see progress
- Force close and try command-line version for debugging

**"Permission denied" errors**
- Close Excel/Word applications before running
- Check file isn't marked as read-only
- Run from location where you have write permissions

### Performance Tips

- **Use .xlsx format** instead of older .xls files
- **Close other Office applications** while running
- **Use smaller test files** to verify setup before processing large files
- **Run on local drives** rather than network drives when possible

## 🏗️ For Developers

### Architecture Overview

```
CLI/GUI Layer
    ↓
Main Script (run_value_into_word.py)
    ↓
Excel Reader ← → Word Writer
    ↓              ↓
Windows COM Automation
    ↓              ↓
Excel.exe     Word.exe
```

### Key Components

**`run_value_into_word.py`**
- Main orchestration logic
- Command-line argument parsing
- Configuration file loading
- Error handling and user feedback

**`excel_reader.py`**
- COM automation for Excel
- Cell value extraction
- Image/range copying (advanced feature)
- Resource cleanup and error handling

**`word_writer.py`**
- COM automation for Word
- Bookmark manipulation
- Document saving and cleanup
- Flexible session management

**`gui_tk.py`**
- Tkinter-based user interface
- Threading for responsive UI
- File dialogs and progress indicators
- GUI-specific error handling

### Extending the Tool

**Adding new formatting options:**
```python
# In run_value_into_word.py, modify format_number()
def format_number(value, fmt_spec):
    if fmt_spec == "currency":
        return f"${value:,.2f}"
    elif fmt_spec == "percentage":
        return f"{value:.1%}"
    # Add your custom formats here
```

**Adding new data sources:**
- Create new reader module following `excel_reader.py` pattern
- Implement `get_value(source, location)` interface
- Import and use in main script

### COM Automation Notes

- **Threading**: Each thread needs `pythoncom.CoInitialize()`
- **Cleanup**: Always close applications to prevent memory leaks
- **Error handling**: COM operations can fail unpredictably
- **Timeouts**: Some operations need delays for reliability

## 📝 License & Credits

This project demonstrates professional automation techniques for Microsoft Office applications using Python. It showcases:

- Windows COM automation
- GUI development with Tkinter
- Command-line interface design
- Error handling and resource management
- Test-driven development practices

**Created for automating document workflows and eliminating manual copy-paste operations.**

---

## 📞 Support

For issues or questions:

1. **Check the troubleshooting section** above
2. **Run tests** to verify installation
3. **Check console output** for detailed error messages
4. **Try command-line version** if GUI has issues

**Remember**: This tool requires Windows and licensed Microsoft Office applications to function properly.