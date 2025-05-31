"""
EXCEL AUTOMATION MODULE
=======================
This module handles all Excel interactions using Windows COM (Component Object Model).
It provides two main capabilities:
1. Reading individual cell values from Excel files
2. Copying Excel ranges as images (EMF format) for embedding in Word/PowerPoint

WINDOWS COM AUTOMATION EXPLAINED:
COM is Microsoft's technology for programs to talk to each other on Windows.
When we "automate" Excel, we're actually remote-controlling a real Excel process
running in the background. This is powerful but can be fragile - Excel might
crash, clipboard operations might fail, etc.

WHY EMF FORMAT?
EMF (Enhanced Metafile) is a vector graphics format that Word and PowerPoint
can embed without quality loss. Unlike PNG/JPEG (bitmap images), EMF files
scale perfectly when zoomed or printed.

PROGRAMMING CONCEPTS DEMONSTRATED:
- COM automation (controlling other Windows applications)
- Windows clipboard manipulation
- Retry patterns for unreliable operations
- Context managers and resource cleanup
- Defensive programming against flaky external systems
- Module-level initialization
- Logging for debugging complex operations
"""

from __future__ import annotations  # Allows using newer type hint syntax in older Python

import os           # File system operations
import time         # Sleep/timing operations  
import tempfile     # Creating temporary files
import logging      # Professional logging instead of print()
import traceback    # Detailed error reporting
from pathlib import Path  # Modern path handling

# Windows-specific libraries for COM automation
import pythoncom           # Core COM support
import win32com.client as win32  # Excel/Office automation
import win32clipboard      # Windows clipboard access
import win32con           # Windows constants

from PIL import ImageGrab  # Pillow library for grabbing clipboard images

# Define what functions other modules can import from this one
__all__ = ["get_value", "copy_range_as_emf"]

# Set up logging for this module (professional debugging)
logger = logging.getLogger(__name__)

# EXCEL CONSTANTS
# These are magic numbers that Excel's COM interface expects
XL_PICTURE = -4147  # Tells Excel to copy as picture (not text)
XL_SCREEN = 1       # Copy as it appears on screen (not as printed)

# CRITICAL: Initialize COM at module level
# COM must be initialized before we can talk to Excel
# This is Windows-specific plumbing that makes automation possible
try:
    pythoncom.CoInitialize()  # "Hey Windows, we want to use COM"
except:
    # If COM is already initialized in this thread, that's fine
    # This can happen in some environments (like Jupyter notebooks)
    pass


# ══════════════════════════════════════════════════════════════
# INTERNAL HELPER FUNCTIONS (private to this module)
# ══════════════════════════════════════════════════════════════

def _open_excel(path: str):
    """
    EXCEL LAUNCHER
    ==============
    Opens Excel invisibly and loads a specific workbook.
    Returns both the Excel application handle and the workbook handle.
    
    This is COM automation in action - we're literally starting Excel
    and remote-controlling it through Windows APIs.
    
    Args:
        path (str): Full path to Excel file
        
    Returns:
        tuple: (excel_application, workbook) - both are COM objects
        
    PROGRAMMING CONCEPTS:
    - COM Dispatch: win32.Dispatch creates a connection to Excel
    - Exception handling: File existence checking
    - Resource management: Caller must clean up what we return
    """
    try:
        # Create connection to Excel application
        # This actually starts Excel.exe if it's not running
        xl = win32.Dispatch("Excel.Application")
        
        try:
            # Make Excel invisible (run in background)
            # Some corporate security policies block this - that's OK
            xl.Visible = False
            
            # Turn off Excel's dialog boxes (like "File already open" warnings)
            xl.DisplayAlerts = False
            
        except AttributeError:
            # If we can't control visibility/alerts, continue anyway
            logger.debug("Could not set Excel visibility or alerts")
            
        # Verify file exists before trying to open it
        if not os.path.exists(path):
            raise FileNotFoundError(f"Excel file not found: {path}")
            
        # Open the workbook in read-only mode (safer, faster)
        wb = xl.Workbooks.Open(path, ReadOnly=True)
        
        return xl, wb  # Return both handles - caller needs both
        
    except Exception as e:
        logger.error(f"Failed to open Excel: {e}")
        raise  # Re-raise the exception so caller knows it failed


def _safe_close(xl, wb):
    """
    EXCEL CLEANUP GUARDIAN
    ======================
    Safely closes Excel workbook and application, even if they're already closed
    or in an error state. This prevents Excel processes from hanging around
    in memory after our script finishes.
    
    CRITICAL IMPORTANCE:
    If we don't properly close Excel, you'll end up with invisible Excel.exe
    processes eating memory. In server environments, this can crash the system!
    
    Args:
        xl: Excel application handle (can be None)
        wb: Workbook handle (can be None)
        
    PROGRAMMING CONCEPTS:
    - Defensive programming: Handle None values and exceptions gracefully
    - Resource cleanup: Always clean up external resources
    - Exception suppression: We don't want cleanup to throw errors
    """
    try:
        # Close workbook first (if it exists)
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)  # Close without saving
            except:
                # If closing fails, continue anyway - maybe it's already closed
                pass
                
        # Close Excel application (if it exists)        
        if xl is not None:
            try:
                xl.Quit()  # Tell Excel to shut down
            except:
                # If quitting fails, continue anyway
                pass
                
    except:
        # If anything else goes wrong, just ignore it
        # We're trying to clean up, not create new problems
        pass


def _clear_clipboard():
    """
    CLIPBOARD CLEANER
    ================
    Clears the Windows clipboard to ensure clean state for copying operations.
    
    WHY THIS IS NEEDED:
    The Windows clipboard can sometimes contain data from previous operations
    that interferes with our image copying. Clearing it first ensures we get
    clean results.
    
    Returns:
        bool: True if clipboard was successfully cleared, False otherwise
        
    PROGRAMMING CONCEPTS:
    - Retry pattern: Try multiple times if operation fails
    - Windows API usage: Direct clipboard manipulation
    - Error recovery: Keep trying even if individual attempts fail
    """
    # Try multiple times because clipboard operations can be flaky
    for attempt in range(3):
        try:
            # Windows clipboard API sequence: Open → Clear → Close
            win32clipboard.OpenClipboard()      # Get exclusive access
            win32clipboard.EmptyClipboard()     # Clear all data
            win32clipboard.CloseClipboard()     # Release access
            return True  # Success!
            
        except Exception as e:
            logger.debug(f"Clipboard clear attempt {attempt+1} failed: {e}")
            time.sleep(0.1)  # Brief pause before retry
            
    return False  # All attempts failed


# ══════════════════════════════════════════════════════════════
# PUBLIC FUNCTIONS (what other modules use)
# ══════════════════════════════════════════════════════════════

def get_value(path: str, sheet: str, cell: str):
    """
    SIMPLE CELL READER
    ==================
    Reads a single value from an Excel cell. This is the simpler of our two
    main functions - just open Excel, read one cell, close Excel.
    
    Args:
        path (str): Full path to Excel workbook
        sheet (str): Worksheet name (e.g., "Sheet1", "Summary")
        cell (str): Cell address (e.g., "A1", "B5", "Z100")
        
    Returns:
        Value from the cell (could be number, text, date, etc.)
        
    Raises:
        ValueError: If cell can't be read (bad sheet name, cell address, etc.)
        
    USAGE EXAMPLE:
        revenue = get_value("budget.xlsx", "Summary", "B10")
        
    PROGRAMMING CONCEPTS:
    - Resource management: Open → Use → Close pattern
    - Exception handling: Convert technical errors to user-friendly ones
    - finally block: Ensures cleanup happens even if errors occur
    """
    xl, wb = None, None  # Initialize to None for safety
    
    try:
        # Open Excel and get handles
        xl, wb = _open_excel(path)
        
        # Navigate: Workbook → Worksheet → Cell → Value
        ws = wb.Worksheets(sheet)   # Get the specific worksheet
        return ws.Range(cell).Value  # Get the cell's value
        
    except Exception as e:
        # Convert technical COM errors to user-friendly messages
        raise ValueError(f"Error getting value from {sheet}!{cell}: {e}")
        
    finally:
        # CRITICAL: Always clean up Excel, even if errors occurred
        _safe_close(xl, wb)


def copy_range_as_emf(
    path: str,
    sheet: str, 
    cell_range: str,
    timeout: float = 15.0,
    retry_count: int = 3,
) -> str:
    """
    EXCEL RANGE TO IMAGE CONVERTER
    ==============================
    This is the complex function - it copies a range of Excel cells as an image
    and saves it as an EMF (Enhanced Metafile) that can be embedded in Word.
    
    THE PROCESS:
    1. Open Excel and navigate to the range
    2. Copy the range as a picture to Windows clipboard
    3. Wait for the clipboard to contain the image (this can take time!)
    4. Grab the image from clipboard using PIL
    5. Save as EMF file
    6. Return path to the EMF file
    
    WHY SO COMPLEX?
    - Excel's clipboard operations are asynchronous (take time)
    - Windows clipboard can be flaky
    - COM automation can fail in many ways
    - We need retry logic to handle all these issues
    
    Args:
        path (str): Full path to Excel workbook
        sheet (str): Worksheet name
        cell_range (str): Excel range like "A1:C10" or "B5:D20"
        timeout (float): Max seconds to wait for clipboard image
        retry_count (int): How many times to retry if it fails
        
    Returns:
        str: Path to temporary EMF file (caller should delete when done)
        
    Raises:
        RuntimeError: If all retry attempts fail
        
    PROGRAMMING CONCEPTS:
    - Retry patterns with exponential backoff
    - Asynchronous operation handling (polling)
    - Windows message pumping for COM
    - Temporary file management
    - Complex error recovery
    """
    
    # Create a secure temporary file path
    # mktemp() creates a unique filename but doesn't create the file yet
    tmp_path = Path(tempfile.mktemp(suffix=".emf"))
    
    # Track success across retry attempts
    success = False
    xl, wb = None, None  # Excel handles
    
    # RETRY LOOP: Try multiple times because this operation is flaky
    for attempt in range(retry_count):
        try:
            logger.debug(f"Attempt {attempt+1}/{retry_count} to copy range as EMF")
            
            # STEP 1: Open Excel and navigate to range
            xl, wb = _open_excel(path)
            
            # Navigate to the specific worksheet
            sht = wb.Worksheets(sheet)
            sht.Activate()  # Make it the active sheet (required for some operations)
            
            # Select the range we want to copy
            rng = sht.Range(cell_range)
            rng.Select()  # Select it (like clicking and dragging in Excel)
            
            # STEP 2: Clear clipboard for clean operation
            _clear_clipboard()
            
            # STEP 3: Copy range as picture (the magic happens here!)
            logger.debug(f"Copying range {cell_range} as picture")
            rng.CopyPicture(
                Appearance=XL_SCREEN,   # Copy as it appears on screen
                Format=XL_PICTURE       # Copy as image (not text)
            )
            
            # STEP 4: Wait for Excel to process the copy operation
            # Excel's clipboard operations are asynchronous - they take time!
            time.sleep(0.5)  # Give Excel a moment
            
            # WINDOWS MESSAGE PUMPING:
            # This is advanced Windows programming - we need to let Windows
            # process its internal messages so COM operations can complete
            for _ in range(10):
                pythoncom.PumpWaitingMessages()  # Process Windows messages
                time.sleep(0.05)  # Brief pause
            
            # STEP 5: POLLING LOOP - Wait for clipboard to contain image
            start = time.time()
            img = None
            
            # Keep checking until we get an image or timeout
            while (time.time() - start) < timeout:
                try:
                    # Keep processing Windows messages (important for COM!)
                    pythoncom.PumpWaitingMessages()
                    
                    # Check if clipboard has image data
                    has_image = False
                    try:
                        win32clipboard.OpenClipboard()
                        # Check for DIB format (Device Independent Bitmap)
                        has_image = win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB)
                        win32clipboard.CloseClipboard()
                    except:
                        # Clipboard access can fail - just try again
                        pass
                    
                    # If clipboard has image, try to grab it
                    if has_image:
                        img = ImageGrab.grabclipboard()  # PIL function to grab clipboard image
                        if img is not None:
                            break  # Success! Got the image
                            
                except Exception as e:
                    logger.debug(f"Error checking clipboard: {e}")
                
                time.sleep(0.1)  # Brief pause before checking again
            
            # STEP 6: Clean up Excel ASAP (don't keep it open longer than necessary)
            _safe_close(xl, wb)
            xl, wb = None, None
            
            # STEP 7: Check if we got an image
            if img is None:
                logger.warning(f"Attempt {attempt+1}: Clipboard did not receive image")
                continue  # Try next attempt
            
            # STEP 8: Save image as EMF file
            try:
                logger.debug(f"Saving image to {tmp_path}")
                img.save(str(tmp_path), "EMF")  # Save as Enhanced Metafile
                success = True
                return str(tmp_path)  # Return path to caller
                
            except Exception as e:
                logger.error(f"Failed to save EMF: {e}")
                
                # FALLBACK: If EMF save fails, try PNG
                try:
                    png_path = str(tmp_path).replace(".emf", ".png")
                    img.save(png_path, "PNG")
                    logger.warning(f"Saved as PNG instead: {png_path}")
                    return png_path
                except:
                    # If both save methods fail, continue to next attempt
                    continue
                
        except Exception as e:
            logger.error(f"Attempt {attempt+1} failed: {str(e)}")
            logger.debug(traceback.format_exc())  # Full error details for debugging
            
            # Clean up Excel before next attempt
            _safe_close(xl, wb)
            xl, wb = None, None
    
    # If we get here, all retry attempts failed
    # Clean up the temporary file path
    try:
        if tmp_path.exists():
            os.unlink(tmp_path)  # Delete the empty temp file
    except:
        pass
            
    # Raise error to let caller know we failed completely
    raise RuntimeError(f"Failed to copy range {cell_range} as EMF after {retry_count} attempts")
