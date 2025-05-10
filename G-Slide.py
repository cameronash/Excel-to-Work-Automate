"""
📌 Version: 3.0
📆 Date: 4/3/2025
✍️ What This Script Does:
- **Works with already open Excel and Word files**: Does not open or close files.
- **Uses hotkey approach for Copy as Picture**: But ensures Excel is properly focused.
- **Automates data transfer**: Transfers numbers, tables, and charts from Excel to Word.
"""

import win32com.client
import time
import pyautogui
import win32gui  # For finding and activating windows
from num2words import num2words  # Ensure installed: pip install num2words

# Define file paths (still needed for identification)
EXCEL_FILE = r"H:\Valuation\1. Doc Archive Folders\Tauranga\T\TAURIKURA DRIVE\420\Lots 580 and 581\2025 01 MV (3035046)\Reports and Spreadsheets\Automation Example\Taurikura Drive 420 and Kiriwehi Street 81 MV (2025)_TEST.xlsm"
WORD_FILE = r"H:\Valuation\1. Doc Archive Folders\Tauranga\T\TAURIKURA DRIVE\420\Lots 580 and 581\2025 01 MV (3035046)\Reports and Spreadsheets\Automation Example\Taurikura Drive 420 and Kiriwehi Street 81 MV (2025)_TEST.docx"

print("🚀 Starting script...")

# Function to ensure Excel window is in focus
def ensure_excel_focus():
    try:
        # Get Excel main window handle
        excel_hwnd = excel.hwnd
        
        # Bring window to foreground
        win32gui.SetForegroundWindow(excel_hwnd)
        
        # Give time for window to come to foreground
    
        time.sleep(0.5)
        print("📊 Excel window activated and brought to foreground")
    except Exception as e:
        print(f"⚠️ Could not focus Excel window directly, trying fallback method: {str(e)}")
        # Fallback - try activating through COM
        try:
            excel.Visible = True  # Make sure it's visible
            wb.Activate()
            excel.ActiveWindow.Activate()
            time.sleep(0.5)
            print("📊 Excel window activated using fallback method")
        except Exception as e2:
            print(f"⚠️ Both focus methods failed: {str(e2)}")
            # Continue anyway

# Get references to already open applications
excel = win32com.client.GetActiveObject("Excel.Application")
word = win32com.client.GetActiveObject("Word.Application")

print("✅ Connected to Excel and Word applications.")

# Find the workbook and document by filename
wb = None
for workbook in excel.Workbooks:
    if workbook.FullName.lower() == EXCEL_FILE.lower():
        wb = workbook
        print(f"✅ Found open Excel workbook: {workbook.Name}")
        break

if wb is None:
    print(f"❌ Could not find open Excel file: {EXCEL_FILE}")
    print("Please make sure the Excel file is open before running this script.")
    exit(1)

doc = None
for document in word.Documents:
    if document.FullName.lower() == WORD_FILE.lower():
        doc = document
        print(f"✅ Found open Word document: {document.Name}")
        break

if doc is None:
    print(f"❌ Could not find open Word file: {WORD_FILE}")
    print("Please make sure the Word file is open before running this script.")
    exit(1)

# Function to insert number from Excel into Word
def insert_number_from_excel(sheet_name, cell_ref, bookmark_name, bold=False, to_words=False, clear_chars=11):
    ws = wb.Sheets(sheet_name)
    value = ws.Range(cell_ref).Value

    if value is None:
        print(f"⚠️ Cell {sheet_name}!{cell_ref} is empty. Skipping insert.")
        return

    # Format number with thousands separator or convert to words
    if not to_words:
        formatted_value = f"{int(value):,}"
    else:
        # Convert to words, title case, but make "And" lowercase
        words = num2words(int(value), lang='en').title()
        words = words.replace(" And ", " and ")  # Keep "and" lowercase
        formatted_value = f"({words} Dollars)"

    # Check if bookmark exists
    if bookmark_name in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks(bookmark_name)
        bookmark_range = bookmark.Range

        # **Clear existing content inside and to the right of the bookmark**
        bookmark_range.MoveEnd(Unit=3, Count=clear_chars)  # Extend selection
        bookmark_range.Text = ""  # Delete existing content

        # Insert new number
        bookmark_range.Text = formatted_value

        # Apply bold formatting if required
        if bold:
            bookmark_range.Font.Bold = True

        print(f"✅ Inserted '{formatted_value}' at '{bookmark_name}' (Old value removed & formatting applied).")

    else:
        print(f"⚠️ Bookmark '{bookmark_name}' not found. Skipping insert.")

# Insert **Adopted Market Value** (Formatted number)
insert_number_from_excel("VAL SUM", "F22", "AdoptedMarketValuePlaceholder", bold=True, clear_chars=11)

# Insert **Market Value in Words** (Converted to words, in bold)
insert_number_from_excel("VAL SUM", "F22", "MarketValueWordsPlaceholder", bold=True, to_words=True, clear_chars=60)

# Function to copy & paste all tables as images
def copy_all_tables():
    tables_to_copy = [
        {"sheet": "VAL SUM", "range": "B2:F28", "bookmark": "ExcelTablePlaceholder"},
        {"sheet": "DCF", "range": "B3:N41", "bookmark": "DCFTablePlaceholder"},
        {"sheet": "RENT", "range": "D3:J531", "bookmark": "ContractRentalTablePlaceholder"},
        {"sheet": "RENT", "range": "D3:N531", "bookmark": "MarketRentalTablePlaceholder"},
        {"sheet": "Sales", "range": "B1:M74", "bookmark": "SalesTablePlaceholder"},
        {"sheet": "CAPVAL", "range": "B3:F33", "bookmark": "IncomeCapTablePlaceholder"},
    ]

    for table in tables_to_copy:
        try:
            print(f"\n📄 Processing table from {table['sheet']}!{table['range']} to {table['bookmark']}...")
            
            # Activate the sheet first using COM objects
            ws = wb.Sheets(table["sheet"])
            wb.Activate()  # Activate the workbook
            ws.Activate()  # Activate the sheet
            
            # Select the range
            table_range = ws.Range(table["range"])
            table_range.Select()
            
            # Try to get focus (both methods)
            try:
                ensure_excel_focus()
            except:
                print("⚠️ Focus error - continuing anyway")
                
            # Make sure Excel is fully the active window before sending keys
            time.sleep(1)
            
            # Apply hotkeys to copy as picture
            print("📸 Using 'Copy as Picture...' for maximum clarity...")
            pyautogui.hotkey("alt", "h")  # Open Home tab
            time.sleep(0.7)
            pyautogui.press("c")  # Open Copy dropdown
            time.sleep(0.7)
            pyautogui.press("p")  # Select "Copy as Picture..."
            time.sleep(1.2)  # Wait longer for window to open

            # Ensure we're on the Copy Picture dialog and press Enter
            pyautogui.press("enter")
            time.sleep(1.2)  # Wait longer for copy to complete

            # Switch to Word and find the bookmark
            if table["bookmark"] in [b.Name for b in doc.Bookmarks]:
                # Activate Word explicitly
                word.Visible = True
                word.Activate()
                doc.Activate()
                time.sleep(0.5)  # Wait for Word to activate
                
                bookmark = doc.Bookmarks(table["bookmark"])
                bookmark_range = bookmark.Range
                
                # STEP 1: IMPORTANT - Delete existing images near this bookmark
                deleted_any = False
                
                # Delete inline shapes (e.g., pictures, charts)
                try:
                    # Create a list first to avoid modification during iteration
                    shapes_to_delete = []
                    for i in range(1, doc.InlineShapes.Count + 1):
                        shape = doc.InlineShapes(i)
                        try:
                            # Check if this shape is near our bookmark
                            if abs(shape.Range.Start - bookmark_range.Start) < 20:
                                shapes_to_delete.append(i)
                        except:
                            continue
                            
                    # Delete the shapes (in reverse order to avoid index changes)
                    for idx in reversed(shapes_to_delete):
                        print(f"🗑 Removing old InlineShape image at '{table['bookmark']}'")
                        doc.InlineShapes(idx).Delete()
                        deleted_any = True
                except Exception as shape_err:
                    print(f"⚠️ Error checking inline shapes: {str(shape_err)}. Continuing...")

                # Delete floating shapes
                try:
                    # Create a list first to avoid modification during iteration
                    shapes_to_delete = []
                    for i in range(1, doc.Shapes.Count + 1):
                        shape = doc.Shapes(i)
                        try:
                            # Check if this shape is near our bookmark
                            if abs(shape.Anchor.Start - bookmark_range.Start) < 20:
                                shapes_to_delete.append(i)
                        except:
                            continue
                            
                    # Delete the shapes (in reverse order to avoid index changes)
                    for idx in reversed(shapes_to_delete):
                        print(f"🗑 Removing old Shape (floating) image at '{table['bookmark']}'")
                        doc.Shapes(idx).Delete()
                        deleted_any = True
                except Exception as shape_err:
                    print(f"⚠️ Error checking floating shapes: {str(shape_err)}. Continuing...")

                if deleted_any:
                    print(f"✅ Successfully removed previous image(s) at bookmark '{table['bookmark']}'")
                
                # STEP 2: Now paste the new image
                # Clear any text at the bookmark
                bookmark_range.Text = ""
                
                # Paste the new image
                print(f"📌 Pasting image at bookmark '{table['bookmark']}'...")
                bookmark_range.Paste()
                
                # Ensure the bookmark still exists (recreate if needed)
                if table["bookmark"] not in [b.Name for b in doc.Bookmarks]:
                    doc.Bookmarks.Add(Name=table["bookmark"], Range=bookmark_range)
                    
                print(f"✅ Table from {table['sheet']}!{table['range']} pasted at '{table['bookmark']}'")
            else:
                print(f"⚠️ Bookmark '{table['bookmark']}' not found in Word document.")
                
        except Exception as e:
            print(f"❌ Error processing {table['sheet']}!{table['range']}: {str(e)}")
            # Continue with next table
            continue

# Run all tables in one go
copy_all_tables()

# Function to copy & paste all charts as images (Paste Special)
def copy_all_charts():
    charts_to_copy = [
        {"sheet": "WALT", "chart": "Chart 2", "bookmark": "WALEPlaceholder"}
    ]

    for chart in charts_to_copy:
        try:
            print(f"\n📊 Processing chart {chart['chart']} from {chart['sheet']} to {chart['bookmark']}...")
            
            # Activate the sheet first
            ws = wb.Sheets(chart["sheet"])
            wb.Activate()  # Activate workbook
            ws.Activate()  # Activate sheet
            
            # Get the chart object
            chart_obj = ws.ChartObjects(chart["chart"])
            chart_obj.Activate()  # Activate chart
            chart_obj.Select()
            
            # Try to get focus
            try:
                ensure_excel_focus()
            except:
                print("⚠️ Focus error for chart - continuing anyway")
                
            # Make sure Excel is fully the active window
            time.sleep(1)
            
            # Try different approach - Copy picture with Chart.Export 
            print("📸 Using alternative method to copy chart...")
            
            # Use keyboard shortcut method since it's most reliable for charts
            try:
                # First, make sure we're selecting just the chart
                excel.Selection.ShapeRange.Select()
                time.sleep(0.3)
                
                # Try first with Alt+H, C, P (Home tab > Copy > Copy as Picture)
                ensure_excel_focus()
                pyautogui.hotkey("alt", "h")
                time.sleep(0.7)
                pyautogui.press("c")
                time.sleep(0.7)
                pyautogui.press("p")
                time.sleep(1.2)
                pyautogui.press("enter")  # Accept default 'As shown on screen'
                time.sleep(1.2)
            except:
                # Fallback to Ctrl+C if the menu approach fails
                print("⚠️ Menu approach failed, trying Ctrl+C")
                pyautogui.hotkey("ctrl", "c")
                time.sleep(1.2)

            # Switch to Word and find the bookmark
            if chart["bookmark"] in [b.Name for b in doc.Bookmarks]:
                # Activate Word explicitly
                word.Visible = True
                word.Activate()
                doc.Activate()
                time.sleep(0.5)  # Wait for Word to activate
                
                bookmark = doc.Bookmarks(chart["bookmark"])
                bookmark_range = bookmark.Range
                
                # STEP 1: IMPORTANT - Delete existing images near this bookmark
                deleted_any = False
                
                # Delete inline shapes (e.g., pictures, charts)
                try:
                    # Create a list first to avoid modification during iteration
                    shapes_to_delete = []
                    for i in range(1, doc.InlineShapes.Count + 1):
                        shape = doc.InlineShapes(i)
                        try:
                            # Check if this shape is near our bookmark
                            if abs(shape.Range.Start - bookmark_range.Start) < 20:
                                shapes_to_delete.append(i)
                        except:
                            continue
                            
                    # Delete the shapes (in reverse order to avoid index changes)
                    for idx in reversed(shapes_to_delete):
                        print(f"🗑 Removing old InlineShape image at '{chart['bookmark']}'")
                        doc.InlineShapes(idx).Delete()
                        deleted_any = True
                except Exception as shape_err:
                    print(f"⚠️ Error checking inline shapes: {str(shape_err)}. Continuing...")

                # Delete floating shapes
                try:
                    # Create a list first to avoid modification during iteration
                    shapes_to_delete = []
                    for i in range(1, doc.Shapes.Count + 1):
                        shape = doc.Shapes(i)
                        try:
                            # Check if this shape is near our bookmark
                            if abs(shape.Anchor.Start - bookmark_range.Start) < 20:
                                shapes_to_delete.append(i)
                        except:
                            continue
                            
                    # Delete the shapes (in reverse order to avoid index changes)
                    for idx in reversed(shapes_to_delete):
                        print(f"🗑 Removing old Shape (floating) image at '{chart['bookmark']}'")
                        doc.Shapes(idx).Delete()
                        deleted_any = True
                except Exception as shape_err:
                    print(f"⚠️ Error checking floating shapes: {str(shape_err)}. Continuing...")

                if deleted_any:
                    print(f"✅ Successfully removed previous image(s) at bookmark '{chart['bookmark']}'")
                
                # STEP 2: Now paste the chart as an image
                # Clear any text at the bookmark
                bookmark_range.Text = ""
                
                # Activate the range where we want to paste
                bookmark_range.Select()
                
                # Try Paste Special as Device Independent Bitmap (best for charts without handles)
                print(f"📌 Pasting chart at bookmark '{chart['bookmark']}' using alternative method...")
                
                try:
                    # Paste as Picture (Device Independent Bitmap)
                    bookmark_range.PasteSpecial(DataType=8)  # 8 = wdPasteDeviceIndependentBitmap
                    
                    # Try to deselect by moving the cursor away and clicking
                    doc.ActiveWindow.Selection.Collapse()
                    
                    # Get the last inserted inline shape and set a standard width
                    if doc.InlineShapes.Count > 0:
                        last_shape = doc.InlineShapes(doc.InlineShapes.Count)
                        last_shape.LockAspectRatio = True
                        last_shape.Width = 450  # Standard width in points
                except Exception as e:
                    print(f"⚠️ Primary paste method failed ({str(e)}), trying alternative...")
                    # Alternative paste method if the special paste fails
                    word.Selection.Paste()
                
                # Move selection away from the pasted object
                word.Selection.MoveRight(1, 1)  # Move right by one character
                
                # Ensure the bookmark still exists
                if chart["bookmark"] not in [b.Name for b in doc.Bookmarks]:
                    doc.Bookmarks.Add(Name=chart["bookmark"], Range=bookmark_range)
                    
                print(f"✅ Chart {chart['chart']} from {chart['sheet']} pasted at '{chart['bookmark']}'")
            else:
                print(f"⚠️ Bookmark '{chart['bookmark']}' not found in Word document.")
                
        except Exception as e:
            print(f"❌ Error processing chart {chart['chart']} on sheet {chart['sheet']}: {str(e)}")
            # Continue with next chart
            continue

# Run all charts in one go
copy_all_charts()

# Function to update a number placeholder_Round 1
def update_number_placeholder(sheet_name, cell_ref, bookmark_name, suffix_text):
    sheet = wb.Sheets(sheet_name)
    value = sheet.Range(cell_ref).Value  # Get value from Excel
    
    if value is None:
        print(f"⚠️ Value is empty in {sheet_name}!{cell_ref}. Skipping update.")
        return
    
    # Format value as ##.# (e.g., 25.0 or 4.32)
    formatted_value = f"{value:.1f}"  # Keep up to 2 decimal places if needed
    
    # Check if bookmark exists in Word
    if bookmark_name in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks(bookmark_name)
        bookmark_range = bookmark.Range

        # Extend the selection to delete text after the placeholder
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range
        bookmark_range.Text = ""  # Delete existing text

        # Insert the new text
        bookmark_range.Text = f"{formatted_value}{suffix_text}"

        print(f"✅ Updated {bookmark_name} to '{formatted_value}{suffix_text}'.")
    else:
        print(f"⚠️ {bookmark_name} not found in Word. Skipping update.")

# Update OPEX Budget Placeholder
update_number_placeholder(
    sheet_name="START",
    cell_ref="C6",
    bookmark_name="OPEXBudgetPlaceholder",
    suffix_text=" per m². These charges are applicable to any gross leases and during vacancy periods (if any)."
)

# Function to update WALE
def update_number_placeholder2(sheet_name, cell_ref, bookmark_name, suffix_text):
    sheet = wb.Sheets(sheet_name)
    value = sheet.Range(cell_ref).Value  # Get value from Excel
    
    if value is None:
        print(f"⚠️ Value is empty in {sheet_name}!{cell_ref}. Skipping update.")
        return
    
    # Format value as ##.# (e.g., 25.0 or 4.32)
    formatted_value = f"{value:.2f}"  # Keep up to 2 decimal places if needed
    
    # Check if bookmark exists in Word
    if bookmark_name in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks(bookmark_name)
        bookmark_range = bookmark.Range

        # Extend the selection to delete text after the placeholder
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range
        bookmark_range.Text = ""  # Delete existing text

        # Insert the new text
        bookmark_range.Text = f"{formatted_value}{suffix_text}"

        print(f"✅ Updated {bookmark_name} to '{formatted_value}{suffix_text}'.")
    else:
        print(f"⚠️ {bookmark_name} not found in Word. Skipping update.")

# Update WALE Number Placeholder
update_number_placeholder2(
    sheet_name="TEN SCHEDULE",
    cell_ref="X32",
    bookmark_name="WALENumberPlaceholder",
    suffix_text=" years, providing long-term lease security."
)

# Function to update a number placeholder_Round 2
def update_number_placeholder3(sheet_name, cell_ref, bookmark_name, suffix_text):
    sheet = wb.Sheets(sheet_name)
    value = sheet.Range(cell_ref).Value  # Get value from Excel
    
    if value is None:
        print(f"⚠️ Value is empty in {sheet_name}!{cell_ref}. Skipping update.")
        return
    
    # Format value with apostrophes for thousands/millions and round to 0 decimal places
    formatted_value = f"{int(round(value)):,}"  # Uses comma as thousands separator
    
    # Check if bookmark exists in Word
    if bookmark_name in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks(bookmark_name)
        bookmark_range = bookmark.Range

        # Extend the selection to delete text after the placeholder
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range
        bookmark_range.Text = ""  # Delete existing text

        # Insert the new text
        bookmark_range.Text = f"{formatted_value}{suffix_text}"

        print(f"✅ Updated {bookmark_name} to '{formatted_value}{suffix_text}'.")
    else:
        print(f"⚠️ {bookmark_name} not found in Word. Skipping update.")

# Update Contract Rental Number Placeholder
update_number_placeholder3(
    sheet_name="RENT",
    cell_ref="J531",
    bookmark_name="ContractRentalNumberPlaceholder",
    suffix_text=" per annum plus GST and operating expenses."
)

# Function to update Market Rental Number Placeholder with multiple values
def update_market_rental_placeholder():
    sheet = wb.Sheets("RENT")
    
    # Extract values from Excel
    rent_value = sheet.Range("L531").Value  # Main rent value
    increase_value = sheet.Range("M531").Value  # Increase amount (can be negative)
    increase_percent = sheet.Range("N531").Value * 100  # Convert decimal to percentage

    # Check if values are empty
    if None in (rent_value, increase_value, increase_percent):
        print(f"⚠️ One or more values are missing in RENT!L531, M531, or N531. Skipping update.")
        return
    
    # Check if market rent equals contract rent (no change)
    if abs(increase_value) < 0.01 or abs(increase_percent) < 0.01:  # Account for tiny rounding differences
        final_text = "We consider the contract rent is within acceptable market tolerances."
    else:
        # Determine whether to say "increase" or "decrease"
        change_word = "increase" if increase_value >= 0 else "decrease"
        
        # Format values: absolute value for increase, add commas for thousands/millions
        formatted_rent = f"${int(round(rent_value)):,}"  # e.g., "$256,340"
        formatted_increase = f"${abs(int(round(increase_value))):,}"  # Always positive e.g., "$26,340"
        formatted_percent = f"{increase_percent:.1f}%"  # Keeps actual sign from Excel e.g., "-11.5%"

        # Construct the final sentence
        final_text = f"We have assessed the market rental to be {formatted_rent} per annum plus GST and operating expenses. This reflects a {change_word} over the contract rent of {formatted_increase} per annum or {formatted_percent}."

    # Check if bookmark exists in Word
    if "MarketRentalNumberPlaceholder" in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks("MarketRentalNumberPlaceholder")
        bookmark_range = bookmark.Range

        # Extend the selection to delete text after the placeholder
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range
        bookmark_range.Text = ""  # Delete existing text

        # Insert the new text
        bookmark_range.Text = final_text

        print(f"✅ Updated MarketRentalNumberPlaceholder to: '{final_text}'")
    else:
        print("⚠️ MarketRentalNumberPlaceholder not found in Word. Skipping update.")

# Run the function to update the placeholder
update_market_rental_placeholder()

# Function to update Market Cap Rate Range Placeholder
def update_cap_rate_range(sheet_name, low_rate_cell, high_rate_cell, bookmark_name):
    sheet = wb.Sheets(sheet_name)
    low_rate = sheet.Range(low_rate_cell).Value  # Get low cap rate from Excel
    high_rate = sheet.Range(high_rate_cell).Value  # Get high cap rate from Excel

    if low_rate is None or high_rate is None:
        print(f"⚠️ One or both values are empty in {sheet_name}!{low_rate_cell}/{high_rate_cell}. Skipping update.")
        return

    # Convert decimal to percentage format
    formatted_low_rate = f"{low_rate * 100:.2f}%"
    formatted_high_rate = f"{high_rate * 100:.2f}%"

    # Construct the final text
    cap_rate_text = f"{formatted_low_rate} - {formatted_high_rate}"

    # Check if bookmark exists in Word
    if bookmark_name in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks(bookmark_name)
        bookmark_range = bookmark.Range

        # Extend selection and clear existing text
        bookmark_range.End = bookmark_range.End + 15
        bookmark_range.Text = cap_rate_text  # Replace with new text

        print(f"✅ Updated {bookmark_name} to '{cap_rate_text}'.")
    else:
        print(f"⚠️ {bookmark_name} not found in Word. Skipping update.")

# Update Market Cap Rate Range Placeholder
update_cap_rate_range(
    sheet_name="VAL SUM",
    low_rate_cell="L5",
    high_rate_cell="M5",
    bookmark_name="MarketCapRateRangePlaceholder"
)

# Update Discount Rate Range Placeholder
update_cap_rate_range(
    sheet_name="VAL SUM",
    low_rate_cell="L19",
    high_rate_cell="M19",
    bookmark_name="DiscountRateRangePlaceholder"
)

# Update Terminal Rate Range Placeholder
update_cap_rate_range(
    sheet_name="VAL SUM",
    low_rate_cell="L20",
    high_rate_cell="M20",
    bookmark_name="TerminalRangePlaceholder"
)

# All other placeholder functions remain the same as your original code
# Function to update Income Cap Range and Midpoint Placeholder with multiple values
def update_incomecap_placeholder():
    sheet = wb.Sheets("VAL SUM")
    
    # Extract values from Excel
    lowcap_percent = sheet.Range("L5").Value * 100  # Convert decimal to percentage
    highcap_percent = sheet.Range("M5").Value * 100  # Convert decimal to percentage
    midpoint_percentage = sheet.Range("F5").Value * 100

    # Check if values are empty
    if None in (lowcap_percent, highcap_percent, midpoint_percentage):
        print(f"⚠️ One or more values are missing in VAL SUM!L5, M5. Skipping update.")
        return
    
    # Format values: absolute value for increase, add commas for thousands/millions
    formatted_lowcap_percent = f"{lowcap_percent:.2f}%"
    formatted_highcap_percent = f"{highcap_percent:.2f}%"
    formatted_midpoint_percent = f"{midpoint_percentage:.2f}%"


    # Construct the final sentence
    final_text = f"{formatted_lowcap_percent} - {formatted_highcap_percent}, accordingly we have adopted a midpoint of {formatted_midpoint_percent}."

    # Check if bookmark exists in Word
    if "IncomeCapRangeMidpointPlaceholder" in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks("IncomeCapRangeMidpointPlaceholder")
        bookmark_range = bookmark.Range

        # Extend the selection to delete text after the placeholder
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range
        bookmark_range.Text = ""  # Delete existing text

        # Insert the new text
        bookmark_range.Text = final_text

        print(f"✅ Updated IncomeCapRangeMidpointPlaceholder to: '{final_text}'")
    else:
        print("⚠️ IncomeCapRangeMidpointPlaceholder not found in Word. Skipping update.")

# Run the function to update the placeholder
update_incomecap_placeholder()

# Function to update Capital Adjustments bullet points from visible cells
def update_capital_adjustments():
    sheet = wb.Sheets("CAPVAL")
    
    # Get only visible cells in B24:B29
    try:
        visible_cells = sheet.Range("B24:B29").SpecialCells(12)  # xlCellTypeVisible (12)
        adjustments = [cell.Value for cell in visible_cells if cell.Value is not None]
    except:
        print("⚠️ No visible cells found in CAPVAL!B24:B29. Skipping update.")
        return

    # Check if we got any visible data
    if not adjustments:
        print("⚠️ No valid data found in visible cells. Skipping update.")
        return

    # Construct text with new lines but **no additional bullet points**
    bullet_text = "\n".join(adjustments)

    # Check if the bookmark exists in Word
    if "CapitalAdjustmentsPlaceholder" in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks("CapitalAdjustmentsPlaceholder")
        bookmark_range = bookmark.Range

        # Extend the range to clear existing text
        bookmark_range.MoveEnd(Unit=3, Count=50)  # Extend range to delete old text
        bookmark_range.Text = ""  # Delete existing bullet points

        # Insert new text (Word will maintain bullet point formatting)
        bookmark_range.Text = bullet_text

        # 🔹 Fix Line Spacing: Set it to **Single Line (No Extra Space)**
        paragraph_format = bookmark_range.ParagraphFormat
        paragraph_format.SpaceAfter = 0  # Remove extra spacing after each bullet point
        paragraph_format.SpaceBefore = 2  # Remove extra spacing before each bullet point
        paragraph_format.LineSpacingRule = 0  # Single line spacing

        print(f"✅ Updated CapitalAdjustmentsPlaceholder with visible adjustments (fixed line spacing).")
    else:
        print("⚠️ CapitalAdjustmentsPlaceholder not found in Word. Skipping update.")

# Run the function to update capital adjustments
update_capital_adjustments()

# Function to update DCF Parameters bullet points with numbers from Excel while keeping the wording
def update_dcf_parameters():
    sheet = wb.Sheets("DCF")
    
    # Extract values from Excel with required formatting
    bullet_1_value = f"{sheet.Range('D37').Value * 100:.1f}"  # 1dp (x100)
    bullet_2_value = f"{sheet.Range('D38').Value:.1f}"  # 1dp
    bullet_3_value = f"{sheet.Range('E40').Value:.1f}"  # 1dp
    year_5_condition = sheet.Range("E39").Value == 0  # If E39 is 0, remove "Year 5"
    bullet_4_value = f"{sheet.Range('K39').Value * 100:.2f}"  # 2dp (x100)
    bullet_5_value = f"{sheet.Range('M37').Value * 100:.2f}"  # 2dp (x100)

    # Define bullet point sentences
    bullet_1 = f"A minimum capital expenditure allowance equivalent to {bullet_1_value}% of gross income per annum."
    bullet_2 = f"A make good allowance of ${bullet_2_value} per m² (in today's dollars) over the floor area."
    
    # Bullet 3: Conditionally remove "Year 5"
    if year_5_condition:
        bullet_3 = f"General refurbishment allowance of ${bullet_3_value} per m² (in today's dollars) over the total floor area of the building in Year 11."
    else:
        bullet_3 = f"General refurbishment allowance of ${bullet_3_value} per m² (in today's dollars) over the total floor area of the building in Year 5 and Year 11."
    
    bullet_4 = f"A discount rate of {bullet_4_value}% based on our analysis of recent sales and having consideration towards the characteristics and risk profile of the property."
    bullet_5 = f"A terminal yield of {bullet_5_value}%."

    # Combine bullets into a single formatted string
    final_bullet_points = (
        f"{bullet_1}\n"
        f"{bullet_2}\n"
        f"{bullet_3}\n"
        f"{bullet_4}\n"
        f"{bullet_5}"
    )

    # Check if bookmark exists in Word
    if "DCFParametersPlaceholder" in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks("DCFParametersPlaceholder")
        bookmark_range = bookmark.Range

        # Extend the selection to remove existing bullet points
        bookmark_range.MoveEnd(Unit=3, Count=100)  # Extend range
        bookmark_range.Text = ""  # Clear existing content

        # Insert new bullet points
        bookmark_range.Text = final_bullet_points

        # 🔹 Fix Line Spacing: Set it to **Single Line (No Extra Space)**
        paragraph_format = bookmark_range.ParagraphFormat
        paragraph_format.SpaceAfter = 2  # Remove extra spacing after each bullet point
        paragraph_format.SpaceBefore = 0  # Remove extra spacing before each bullet point
        paragraph_format.LineSpacingRule = 0  # Single line spacing

        print(f"✅ Updated DCFParametersPlaceholder with new bullet points.")
    else:
        print("⚠️ DCFParametersPlaceholder not found in Word. Skipping update.")

# Run the function to update the DCF Parameters section
update_dcf_parameters()

# Function to update Sales Yield Ranges placeholder with new numbers from Excel
def update_sales_yield_ranges():
    sheet = wb.Sheets("Sales")
    
    # Extract values from Excel and multiply by 100 for percentage formatting
    initial_low = f"{sheet.Range('E72').Value * 100:.2f}"  # 2dp
    initial_high = f"{sheet.Range('E74').Value * 100:.2f}"  # 2dp
    equivalent_low = f"{sheet.Range('F72').Value * 100:.2f}"  # 2dp
    equivalent_high = f"{sheet.Range('F74').Value * 100:.2f}"  # 2dp
    irr_low = f"{sheet.Range('H72').Value * 100:.2f}"  # 2dp
    irr_high = f"{sheet.Range('H74').Value * 100:.2f}"  # 2dp

    # Construct the updated text block with new values
    updated_text = (
        f"The above sales reflect an initial yield range of {initial_low}% - {initial_high}%, "
        f"and equivalent yield range of {equivalent_low}% - {equivalent_high}% and an IRR range of {irr_low}% - {irr_high}%. "
        "The upper end of these ranges reflects less desirable investment property, whereas the lower end of these ranges "
        "reflects sought after property underpinned by strong lease covenant, future rental growth expectations or lower value quantum."
    )

    # Check if the bookmark exists in Word
    if "SalesYieldRangesPlaceholder" in [b.Name for b in doc.Bookmarks]:
        bookmark = doc.Bookmarks("SalesYieldRangesPlaceholder")
        bookmark_range = bookmark.Range

        # Extend the selection to clear existing text
        bookmark_range.MoveEnd(Unit=3, Count=500)  # Extend range
        bookmark_range.Text = ""  # Delete existing content

        # Insert the new updated text
        bookmark_range.Text = updated_text

        print(f"✅ Updated SalesYieldRangesPlaceholder with new yield ranges.")
    else:
        print("⚠️ SalesYieldRangesPlaceholder not found in Word. Skipping update.")

# Run the function to update the Sales Yield Ranges
update_sales_yield_ranges()

print("✅ Script finished successfully!")