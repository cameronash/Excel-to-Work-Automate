import win32com.client as win32
import pythoncom
import time
import os

def read_table_config(config_path):
    """
    Read the table mapping configuration from Excel file
    Expects a "Tables" tab with columns: sheet_name, cell_range, bookmark_name
    """
    pythoncom.CoInitialize()
    excel = None
    
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(config_path)
        
        # Check if Tables tab exists
        sheet_names = [ws.Name for ws in wb.Worksheets]
        if "Tables" not in sheet_names:
            print("Note: No 'Tables' tab found in config file")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return []
        
        ws = wb.Worksheets("Tables")
        
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
        
        table_configs = []
        for row in range(2, last_row + 1):
            sheet_name = ws.Cells(row, 1).Value      # Column A
            cell_range = ws.Cells(row, 2).Value      # Column B
            bookmark_name = ws.Cells(row, 3).Value   # Column C
            
            if sheet_name and cell_range and bookmark_name:
                table_configs.append({
                    'sheet_name': sheet_name,
                    'cell_range': cell_range,
                    'bookmark_name': bookmark_name
                })
        
        wb.Close(SaveChanges=False)
        excel.Quit()
        
        print(f"Found {len(table_configs)} table(s) to process")
        return table_configs
        
    except Exception as e:
        print(f"Error reading table config: {e}")
        if excel:
            excel.Quit()
        return []
    finally:
        pythoncom.CoUninitialize()

def process_all_tables_fast(config_path=None, excel_path=None, word_path=None, delay_seconds=0.3):
    """
    Fast batch processing for tables - opens Excel/Word once for all tables
    """
    # Default file paths
    if config_path is None:
        config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    if excel_path is None:
        excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    if word_path is None:
        word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("=== Excel Tables to Word - Fast Batch Mode ===\n")
    
    # Check files exist
    if not os.path.exists(config_path):
        print(f"❌ Config file not found: {config_path}")
        return
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return
    if not os.path.exists(word_path):
        print(f"❌ Word file not found: {word_path}")
        return
    
    # Read configurations
    print(f"📖 Reading table config from: {os.path.basename(config_path)}")
    table_configs = read_table_config(config_path)
    
    if not table_configs:
        print("No tables to process! Make sure you have a 'Tables' tab in your config file.")
        return
    
    print(f"\n⚡ Fast processing {len(table_configs)} table(s)...\n")
    
    # Initialize COM
    pythoncom.CoInitialize()
    excel = None
    word = None
    
    start_time = time.time()
    
    try:
        # Open Excel and Word ONCE
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        word.ScreenUpdating = False
        
        # Open documents ONCE
        wb = excel.Workbooks.Open(excel_path)
        doc = word.Documents.Open(word_path)
        
        success_count = 0
        
        # Process all tables without closing/reopening
        for i, config in enumerate(table_configs, 1):
            sheet_name = config['sheet_name']
            cell_range = config['cell_range']
            bookmark_name = config['bookmark_name']
            
            print(f"{i}. {cell_range} → {bookmark_name}... ", end='', flush=True)
            
            try:
                # Get worksheet
                ws = wb.Worksheets(sheet_name)
                
                # Copy the range
                range_to_copy = ws.Range(cell_range)
                range_to_copy.Copy()  # Use Copy() for EMF format
                
                # Delay for clipboard
                time.sleep(delay_seconds)
                
                # Check if bookmark exists
                if bookmark_name not in [bm.Name for bm in doc.Bookmarks]:
                    print("❌ Bookmark not found")
                    continue
                
                # Get bookmark
                bookmark = doc.Bookmarks(bookmark_name)
                bookmark_range = bookmark.Range
                bookmark_start = bookmark_range.Start
                
                # Clear existing content
                # Clear tables first
                tables_to_delete = []
                for table in bookmark_range.Tables:
                    tables_to_delete.append(table)
                for table in tables_to_delete:
                    table.Delete()
                
                # Clear shapes
                shapes_to_delete = []
                for shape in bookmark_range.InlineShapes:
                    shapes_to_delete.append(shape)
                for shape in shapes_to_delete:
                    shape.Delete()
                
                bookmark_range.Text = ""
                
                # Paste as EMF
                bookmark_range.Select()
                word.Selection.PasteSpecial(
                    Link=False,
                    DataType=9,  # EMF format
                    Placement=0,  # wdInLine
                    DisplayAsIcon=False
                )
                
                # Recreate bookmark properly
                word.Selection.MoveRight(Unit=1, Count=1)  # Move after pasted content
                word.Selection.TypeText(" ")  # Add space
                new_range = doc.Range(Start=bookmark_start, End=word.Selection.Range.End)
                doc.Bookmarks.Add(Name=bookmark_name, Range=new_range)
                
                print("✓")
                success_count += 1
                
            except Exception as e:
                print(f"❌ {str(e)}")
        
        # Save ONCE at the end
        print("\n💾 Saving document...")
        doc.Save()
        
        # Calculate time taken
        elapsed_time = time.time() - start_time
        
        print(f"\n{'='*50}")
        print(f"✅ Complete: {success_count}/{len(table_configs)} tables processed")
        print(f"⏱️  Time taken: {elapsed_time:.1f} seconds")
        print(f"📊 That's {elapsed_time/len(table_configs):.1f} seconds per table!")
        
        # Close documents
        doc.Close()
        wb.Close(SaveChanges=False)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Cleanup
        time.sleep(0.5)
        try:
            if word:
                word.ScreenUpdating = True
                word.Quit()
        except:
            pass
        try:
            if excel:
                excel.ScreenUpdating = True
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

def process_all_content_fast(config_path=None, excel_path=None, word_path=None, delay_seconds=0.3):
    """
    Process BOTH tables and charts in one run!
    Reads from both "Tables" and "Graphs" tabs in the config file
    """
    # Default file paths
    if config_path is None:
        config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    if excel_path is None:
        excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    if word_path is None:
        word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("=== Excel Tables & Charts to Word - Combined Fast Mode ===\n")
    
    # Check files exist
    if not os.path.exists(config_path):
        print(f"❌ Config file not found: {config_path}")
        return
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return
    if not os.path.exists(word_path):
        print(f"❌ Word file not found: {word_path}")
        return
    
    # Read both configs
    print(f"📖 Reading config from: {os.path.basename(config_path)}")
    
    # Import the chart config reader from your other file
    from excel_charts_fast import read_chart_config
    
    table_configs = read_table_config(config_path)
    chart_configs = read_chart_config(config_path)
    
    total_items = len(table_configs) + len(chart_configs)
    if total_items == 0:
        print("No tables or charts to process!")
        return
    
    print(f"\n⚡ Processing {len(table_configs)} table(s) and {len(chart_configs)} chart(s)...\n")
    
    # Initialize COM
    pythoncom.CoInitialize()
    excel = None
    word = None
    
    start_time = time.time()
    
    try:
        # Open Excel and Word ONCE
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        word.ScreenUpdating = False
        
        # Open documents ONCE
        wb = excel.Workbooks.Open(excel_path)
        doc = word.Documents.Open(word_path)
        
        success_count = 0
        item_number = 0
        
        # Process all tables first
        if table_configs:
            print("📊 Processing Tables:\n")
            for config in table_configs:
                item_number += 1
                sheet_name = config['sheet_name']
                cell_range = config['cell_range']
                bookmark_name = config['bookmark_name']
                
                print(f"{item_number}. Table: {cell_range} → {bookmark_name}... ", end='', flush=True)
                
                try:
                    ws = wb.Worksheets(sheet_name)
                    range_to_copy = ws.Range(cell_range)
                    range_to_copy.Copy()
                    
                    time.sleep(delay_seconds)
                    
                    if bookmark_name in [bm.Name for bm in doc.Bookmarks]:
                        bookmark = doc.Bookmarks(bookmark_name)
                        bookmark_range = bookmark.Range
                        bookmark_start = bookmark_range.Start
                        
                        # Clear existing content
                        for table in bookmark_range.Tables:
                            table.Delete()
                        for shape in bookmark_range.InlineShapes:
                            shape.Delete()
                        bookmark_range.Text = ""
                        
                        # Paste
                        bookmark_range.Select()
                        word.Selection.PasteSpecial(DataType=9)
                        
                        # Recreate bookmark
                        word.Selection.MoveRight(Unit=1, Count=1)
                        word.Selection.TypeText(" ")
                        new_range = doc.Range(Start=bookmark_start, End=word.Selection.Range.End)
                        doc.Bookmarks.Add(Name=bookmark_name, Range=new_range)
                        
                        print("✓")
                        success_count += 1
                    else:
                        print("❌ Bookmark not found")
                        
                except Exception as e:
                    print(f"❌ {str(e)}")
        
        # Process all charts
        if chart_configs:
            print("\n📈 Processing Charts:\n")
            for config in chart_configs:
                item_number += 1
                sheet_name = config['sheet_name']
                chart_name = config['chart_name']
                bookmark_name = config['bookmark_name']
                
                print(f"{item_number}. Chart: {chart_name} → {bookmark_name}... ", end='', flush=True)
                
                try:
                    ws = wb.Worksheets(sheet_name)
                    
                    chart_found = False
                    for chart_obj in ws.ChartObjects():
                        if chart_obj.Name == chart_name:
                            chart_obj.Copy()
                            chart_found = True
                            break
                    
                    if not chart_found:
                        print("❌ Chart not found")
                        continue
                    
                    time.sleep(delay_seconds)
                    
                    if bookmark_name in [bm.Name for bm in doc.Bookmarks]:
                        bookmark = doc.Bookmarks(bookmark_name)
                        bookmark_range = bookmark.Range
                        bookmark_start = bookmark_range.Start
                        
                        # Clear existing content
                        for shape in bookmark_range.InlineShapes:
                            shape.Delete()
                        bookmark_range.Text = ""
                        
                        # Paste
                        bookmark_range.Select()
                        word.Selection.PasteSpecial(DataType=9)
                        
                        # Recreate bookmark
                        word.Selection.MoveRight(Unit=1, Count=1)
                        word.Selection.TypeText(" ")
                        new_range = doc.Range(Start=bookmark_start, End=word.Selection.Range.End)
                        doc.Bookmarks.Add(Name=bookmark_name, Range=new_range)
                        
                        print("✓")
                        success_count += 1
                    else:
                        print("❌ Bookmark not found")
                        
                except Exception as e:
                    print(f"❌ {str(e)}")
        
        # Save ONCE at the end
        print("\n💾 Saving document...")
        doc.Save()
        
        # Calculate time taken
        elapsed_time = time.time() - start_time
        
        print(f"\n{'='*50}")
        print(f"✅ Complete: {success_count}/{total_items} items processed")
        print(f"⏱️  Time taken: {elapsed_time:.1f} seconds")
        print(f"📊 That's {elapsed_time/total_items:.1f} seconds per item!")
        
        # Close documents
        doc.Close()
        wb.Close(SaveChanges=False)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Cleanup
        time.sleep(0.5)
        try:
            if word:
                word.ScreenUpdating = True
                word.Quit()
        except:
            pass
        try:
            if excel:
                excel.ScreenUpdating = True
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Process just tables
    process_all_tables_fast()
    
    # Or process BOTH tables and charts in one run!
    # process_all_content_fast()
    
    # Or with a longer delay if needed
    # process_all_tables_fast(delay_seconds=0.8)