"""
Test script for excel_to_word_handler module
Run this from the G-SLIDE root directory
"""
import sys
import os
import logging
from pathlib import Path

# Add src to path so we can import from gslide
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from gslide.excel_to_word_handler import ExcelToWordHandler, process_excel_to_word

# Set up logging to see what's happening
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def test_excel_to_word():
    """Test the excel to word handler"""
    
    # File paths - UPDATE THESE AS NEEDED
    config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("=" * 60)
    print("Testing Excel to Word Handler")
    print("=" * 60)
    print(f"\nConfig: {os.path.basename(config_path)}")
    print(f"Excel:  {os.path.basename(excel_path)}")
    print(f"Word:   {os.path.basename(word_path)}")
    print()
    
    # Test 1: Charts only
    print("\n" + "="*60)
    print("TEST 1: Processing CHARTS only")
    print("="*60)
    success, total = process_excel_to_word(
        config_path=config_path,
        excel_path=excel_path,
        word_path=word_path,
        content_type="charts",
        delay_seconds=0.3
    )
    print(f"\nResult: {success}/{total} charts processed successfully")
    
    # Wait a bit between tests
    input("\nPress Enter to continue to TEST 2...")
    
    # Test 2: Tables only
    print("\n" + "="*60)
    print("TEST 2: Processing TABLES only")
    print("="*60)
    success, total = process_excel_to_word(
        config_path=config_path,
        excel_path=excel_path,
        word_path=word_path,
        content_type="tables",
        delay_seconds=0.3
    )
    print(f"\nResult: {success}/{total} tables processed successfully")
    
    # Wait a bit between tests
    input("\nPress Enter to continue to TEST 3...")
    
    # Test 3: Both charts and tables
    print("\n" + "="*60)
    print("TEST 3: Processing BOTH charts and tables")
    print("="*60)
    success, total = process_excel_to_word(
        config_path=config_path,
        excel_path=excel_path,
        word_path=word_path,
        content_type="both",
        delay_seconds=0.3
    )
    print(f"\nResult: {success}/{total} items processed successfully")
    
    print("\n" + "="*60)
    print("Testing complete!")
    print("="*60)

def test_with_class():
    """Test using the class directly (more control)"""
    
    config_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\config\g-slide mapping.xlsx"
    excel_path = r"H:\Valuation\Commercial\Valuer Resources\Forbury\G-Slide\tests\assets\Sample2.xlsx"
    word_path = r"C:\Users\CSpeirs-Hutton\Desktop\Template2.docm"
    
    print("\n" + "="*60)
    print("Testing with ExcelToWordHandler class")
    print("="*60)
    
    # Create handler instance
    handler = ExcelToWordHandler(config_path, excel_path, word_path, delay_seconds=0.3)
    
    # Check files exist
    if not handler.validate_files():
        print("ERROR: One or more files not found!")
        return
    
    # Read configs to see what we're working with
    chart_configs = handler.read_chart_config()
    table_configs = handler.read_table_config()
    
    print(f"\nConfig summary:")
    print(f"- Charts found: {len(chart_configs)}")
    print(f"- Tables found: {len(table_configs)}")
    
    # Process everything
    success, total = handler.process_all_content("both")
    print(f"\nFinal result: {success}/{total} items processed")

if __name__ == "__main__":
    # Run the main test
    test_excel_to_word()
    
    # Uncomment to test the class-based approach
    # test_with_class()