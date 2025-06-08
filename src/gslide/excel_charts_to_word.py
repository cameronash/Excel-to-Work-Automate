import win32com.client as win32
import pythoncom
import time
import os

def read_chart_config(config_path):
    """
    Read the chart mapping configuration from Excel file using win32com
    """
    pythoncom.CoInitialize()
    excel = None
    
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(config_path)
        ws = wb.Worksheets("Graphs")
        
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
        
        chart_configs = []
        for row in range(2, last_row + 1):
            sheet_name = ws.Cells(row, 1).Value
            chart_name = ws.Cells(row, 2).Value
            bookmark_name = ws.Cells(row, 3).Value
            
            if sheet_name and chart_name and bookmark_name:
                chart_configs.append({
                    'sheet_name': sheet_name,
                    'chart_name': chart_name,
                    'bookmark_name': bookmark_name
                })
        
        wb.Close(SaveChanges=False)
        excel.Quit()
        
        print(f"Found {len(chart_configs)} chart(s) to process")
        return chart_configs
        
    except Exception as e:
        print(f"Error reading config file: {e}")
        if excel:
            excel.Quit()
        return []
    finally:
        pythoncom.CoUninitialize()

def process_all_charts_fast(config_path=None, excel_path=None, word_path=None):
    """
    FAST batch processing - opens Excel/Word once for all charts
    Should take ~30 seconds instead of 10 minutes!
    """
    # Default file paths
    if config_path is None:
        config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    if excel_path is None:
        excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    if word_path is None:
        word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("=== Excel Charts to Word - FAST Batch Mode ===\n")
    
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
    print(f"📖 Reading config from: {os.path.basename(config_path)}")
    chart_configs = read_chart_config(config_path)
    
    if not chart_configs:
        print("No charts to process!")
        return
    
    print(f"\n⚡ Fast processing {len(chart_configs)} charts...\n")
    
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
        
        # Process all charts without closing/reopening
        for i, config in enumerate(chart_configs, 1):
            sheet_name = config['sheet_name']
            chart_name = config['chart_name']
            bookmark_name = config['bookmark_name']
            
            print(f"{i}. {chart_name} → {bookmark_name}... ", end='', flush=True)
            
            try:
                # Get worksheet
                ws = wb.Worksheets(sheet_name)
                
                # Find and copy chart
                chart_found = False
                for chart_obj in ws.ChartObjects():
                    if chart_obj.Name == chart_name:
                        chart_obj.Copy()
                        chart_found = True
                        break
                
                if not chart_found:
                    print("❌ Chart not found")
                    continue
                
                # Minimal delay for clipboard
                time.sleep(0.3)
                
                # Check if bookmark exists
                if bookmark_name not in [bm.Name for bm in doc.Bookmarks]:
                    print("❌ Bookmark not found")
                    continue
                
                # Get bookmark
                bookmark = doc.Bookmarks(bookmark_name)
                bookmark_range = bookmark.Range
                bookmark_start = bookmark_range.Start
                
                # Clear existing content
                shapes_to_delete = []
                for shape in bookmark_range.InlineShapes:
                    shapes_to_delete.append(shape)
                
                for shape in shapes_to_delete:
                    shape.Delete()
                
                bookmark_range.Text = ""
                
                # Paste
                bookmark_range.Select()
                word.Selection.PasteSpecial(
                    Link=False,
                    DataType=9,  # EMF
                    Placement=0,
                    DisplayAsIcon=False
                )
                
                # Recreate bookmark
                word.Selection.MoveRight(Unit=1, Count=1)
                word.Selection.TypeText(" ")
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
        print(f"✅ Complete: {success_count}/{len(chart_configs)} charts processed")
        print(f"⏱️  Time taken: {elapsed_time:.1f} seconds")
        print(f"🚀 That's {elapsed_time/len(chart_configs):.1f} seconds per chart!")
        
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

def process_single_chart_fast(row_number, config_path=None, excel_path=None, word_path=None):
    """
    Process a single chart (for testing/debugging)
    """
    if config_path is None:
        config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    if excel_path is None:
        excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    if word_path is None:
        word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    # Read config and get specific chart
    chart_configs = read_chart_config(config_path)
    
    if row_number < 2 or row_number > len(chart_configs) + 1:
        print(f"Invalid row number. Use between 2 and {len(chart_configs) + 1}")
        return
    
    # Use the fast batch processor for just one chart
    single_config = [chart_configs[row_number - 2]]
    
    # Temporarily replace the config reading
    original_read = read_chart_config
    globals()['read_chart_config'] = lambda x: single_config
    
    process_all_charts_fast(config_path, excel_path, word_path)
    
    # Restore original
    globals()['read_chart_config'] = original_read

if __name__ == "__main__":
    # FAST batch processing (recommended)
    process_all_charts_fast()
    
    # Or process a single chart for testing
    # process_single_chart_fast(row_number=2)