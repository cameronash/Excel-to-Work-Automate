"""
Excel to Word Handler for G-SLIDE
Handles both charts and tables from Excel to Word bookmarks
"""
import win32com.client as win32
import pythoncom
import time
import os
import logging
from typing import List, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class ExcelToWordHandler:
    """Handles copying Excel content (charts and tables) to Word bookmarks"""
    
    def __init__(self, config_path: str, excel_path: str, word_path: str, delay_seconds: float = 0.3):
        """
        Initialize the handler with file paths
        
        Args:
            config_path: Path to g-slide mapping.xlsx config file
            excel_path: Path to source Excel file
            word_path: Path to destination Word document
            delay_seconds: Delay between copy/paste operations (default 0.3)
        """
        self.config_path = config_path
        self.excel_path = excel_path
        self.word_path = word_path
        self.delay_seconds = delay_seconds
        
    def validate_files(self) -> bool:
        """Check if all required files exist"""
        files_valid = True
        
        if not os.path.exists(self.config_path):
            logger.error(f"Config file not found: {self.config_path}")
            files_valid = False
            
        if not os.path.exists(self.excel_path):
            logger.error(f"Excel file not found: {self.excel_path}")
            files_valid = False
            
        if not os.path.exists(self.word_path):
            logger.error(f"Word file not found: {self.word_path}")
            files_valid = False
            
        return files_valid
    
    def read_config(self, tab_name: str, columns: List[str]) -> List[Dict[str, str]]:
        """
        Read configuration from Excel config file
        
        Args:
            tab_name: Name of the worksheet tab to read
            columns: List of column names to read (e.g., ['sheet_name', 'chart_name', 'bookmark_name'])
            
        Returns:
            List of dictionaries with configuration data
        """
        pythoncom.CoInitialize()
        excel = None
        configs = []
        
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(self.config_path)
            
            # Check if tab exists
            sheet_names = [ws.Name for ws in wb.Worksheets]
            if tab_name not in sheet_names:
                logger.info(f"No '{tab_name}' tab found in config file")
                wb.Close(SaveChanges=False)
                excel.Quit()
                return []
            
            ws = wb.Worksheets(tab_name)
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
            
            # Read data starting from row 2 (skip header)
            for row in range(2, last_row + 1):
                config_item = {}
                all_values_present = True
                
                for col_idx, col_name in enumerate(columns, 1):
                    value = ws.Cells(row, col_idx).Value
                    if value:
                        config_item[col_name] = value
                    else:
                        all_values_present = False
                        break
                
                if all_values_present:
                    configs.append(config_item)
            
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            logger.info(f"Found {len(configs)} {tab_name.lower()} to process")
            return configs
            
        except Exception as e:
            logger.error(f"Error reading {tab_name} config: {e}")
            if excel:
                excel.Quit()
            return []
        finally:
            pythoncom.CoUninitialize()
    
    def read_chart_config(self) -> List[Dict[str, str]]:
        """Read chart configuration from Graphs tab"""
        return self.read_config("Graphs", ["sheet_name", "chart_name", "bookmark_name"])
    
    def read_table_config(self) -> List[Dict[str, str]]:
        """Read table configuration from Tables tab"""
        return self.read_config("Tables", ["sheet_name", "cell_range", "bookmark_name"])
    
    def process_all_content(self, content_type: str = "both") -> Tuple[int, int]:
        """
        Process Excel content to Word bookmarks
        
        Args:
            content_type: "charts", "tables", or "both" (default)
            
        Returns:
            Tuple of (success_count, total_count)
        """
        if not self.validate_files():
            return 0, 0
        
        # Read configurations based on content type
        table_configs = []
        chart_configs = []
        
        if content_type in ["tables", "both"]:
            table_configs = self.read_table_config()
            
        if content_type in ["charts", "both"]:
            chart_configs = self.read_chart_config()
        
        total_items = len(table_configs) + len(chart_configs)
        if total_items == 0:
            logger.info(f"No {content_type} to process")
            return 0, 0
        
        logger.info(f"Processing {len(table_configs)} table(s) and {len(chart_configs)} chart(s)")
        
        # Initialize COM and process
        pythoncom.CoInitialize()
        excel = None
        word = None
        success_count = 0
        
        try:
            # Open Excel and Word once for all operations
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            word.ScreenUpdating = False
            
            # Open documents once
            wb = excel.Workbooks.Open(self.excel_path)
            doc = word.Documents.Open(self.word_path)
            
            # Process tables
            for config in table_configs:
                if self._process_table(wb, doc, word, config):
                    success_count += 1
            
            # Process charts
            for config in chart_configs:
                if self._process_chart(wb, doc, word, config):
                    success_count += 1
            
            # Save document once at the end
            logger.info("Saving Word document...")
            doc.Save()
            
            # Close documents
            doc.Close()
            wb.Close(SaveChanges=False)
            
            logger.info(f"Completed: {success_count}/{total_items} items processed successfully")
            return success_count, total_items
            
        except Exception as e:
            logger.error(f"Critical error during processing: {e}")
            return success_count, total_items
            
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
    
    def _process_table(self, wb, doc, word, config: Dict[str, str]) -> bool:
        """Process a single table"""
        sheet_name = config['sheet_name']
        cell_range = config['cell_range']
        bookmark_name = config['bookmark_name']
        
        logger.info(f"Processing table: {cell_range} → {bookmark_name}")
        
        try:
            # Get worksheet and copy range
            ws = wb.Worksheets(sheet_name)
            range_to_copy = ws.Range(cell_range)
            range_to_copy.Copy()
            
            # Delay for clipboard
            time.sleep(self.delay_seconds)
            
            # Check if bookmark exists
            if bookmark_name not in [bm.Name for bm in doc.Bookmarks]:
                logger.warning(f"Bookmark '{bookmark_name}' not found")
                return False
            
            # Process bookmark
            bookmark = doc.Bookmarks(bookmark_name)
            bookmark_range = bookmark.Range
            bookmark_start = bookmark_range.Start
            
            # Clear existing content
            self._clear_bookmark_content(bookmark_range)
            
            # Paste as EMF
            bookmark_range.Select()
            word.Selection.PasteSpecial(
                Link=False,
                DataType=9,  # EMF format
                Placement=0,  # wdInLine
                DisplayAsIcon=False
            )
            
            # Recreate bookmark - use exact logic from original working code
            word.Selection.MoveRight(Unit=1, Count=1)  # Move after pasted content
            word.Selection.TypeText(" ")  # Add space
            new_range = doc.Range(Start=bookmark_start, End=word.Selection.Range.End)
            doc.Bookmarks.Add(Name=bookmark_name, Range=new_range)
            
            logger.info(f"✓ Successfully processed table: {cell_range}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to process table {cell_range}: {e}")
            return False
    
    def _process_chart(self, wb, doc, word, config: Dict[str, str]) -> bool:
        """Process a single chart"""
        sheet_name = config['sheet_name']
        chart_name = config['chart_name']
        bookmark_name = config['bookmark_name']
        
        logger.info(f"Processing chart: {chart_name} → {bookmark_name}")
        
        try:
            # Get worksheet and find chart
            ws = wb.Worksheets(sheet_name)
            
            chart_found = False
            for chart_obj in ws.ChartObjects():
                if chart_obj.Name == chart_name:
                    chart_obj.Copy()
                    chart_found = True
                    break
            
            if not chart_found:
                logger.warning(f"Chart '{chart_name}' not found on sheet '{sheet_name}'")
                return False
            
            # Delay for clipboard
            time.sleep(self.delay_seconds)
            
            # Check if bookmark exists
            if bookmark_name not in [bm.Name for bm in doc.Bookmarks]:
                logger.warning(f"Bookmark '{bookmark_name}' not found")
                return False
            
            # Process bookmark
            bookmark = doc.Bookmarks(bookmark_name)
            bookmark_range = bookmark.Range
            bookmark_start = bookmark_range.Start
            
            # Clear existing content
            self._clear_bookmark_content(bookmark_range)
            
            # Paste as EMF
            bookmark_range.Select()
            word.Selection.PasteSpecial(
                Link=False,
                DataType=9,  # EMF format
                Placement=0,  # wdInLine
                DisplayAsIcon=False
            )
            
            # Recreate bookmark - use exact logic from original working code
            word.Selection.MoveRight(Unit=1, Count=1)  # Move after pasted content
            word.Selection.TypeText(" ")  # Add space
            new_range = doc.Range(Start=bookmark_start, End=word.Selection.Range.End)
            doc.Bookmarks.Add(Name=bookmark_name, Range=new_range)
            
            logger.info(f"✓ Successfully processed chart: {chart_name}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to process chart {chart_name}: {e}")
            return False
    
    def _clear_bookmark_content(self, bookmark_range):
        """Clear all content from a bookmark range"""
        # Clear tables
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
        
        # Clear text
        bookmark_range.Text = ""


# Convenience functions for backward compatibility
def process_excel_to_word(config_path: str, excel_path: str, word_path: str, 
                         content_type: str = "both", delay_seconds: float = 0.3) -> Tuple[int, int]:
    """
    Process Excel content to Word bookmarks
    
    Args:
        config_path: Path to g-slide mapping.xlsx
        excel_path: Path to source Excel file
        word_path: Path to destination Word file
        content_type: "charts", "tables", or "both"
        delay_seconds: Delay between operations
        
    Returns:
        Tuple of (success_count, total_count)
    """
    handler = ExcelToWordHandler(config_path, excel_path, word_path, delay_seconds)
    return handler.process_all_content(content_type)


if __name__ == "__main__":
    """Run directly for testing"""
    # Set up logging to console
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # Default file paths
    config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("=" * 60)
    print("Excel to Word Handler - Test Run")
    print("=" * 60)
    print(f"\nConfig: {os.path.basename(config_path)}")
    print(f"Excel:  {os.path.basename(excel_path)}")
    print(f"Word:   {os.path.basename(word_path)}")
    
    # Ask what to process
    print("\nWhat would you like to process?")
    print("1. Charts only")
    print("2. Tables only")
    print("3. Both charts and tables")
    
    choice = input("\nEnter choice (1-3): ").strip()
    
    content_type = {
        "1": "charts",
        "2": "tables", 
        "3": "both"
    }.get(choice, "both")
    
    print(f"\nProcessing {content_type}...")
    
    # Process
    success, total = process_excel_to_word(
        config_path=config_path,
        excel_path=excel_path,
        word_path=word_path,
        content_type=content_type,
        delay_seconds=0.3
    )
    
    print(f"\n{'='*60}")
    print(f"Complete: {success}/{total} items processed successfully")
    print(f"{'='*60}")