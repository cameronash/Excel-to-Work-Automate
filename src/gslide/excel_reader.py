from __future__ import annotations

"""Excel helpers: read a single value or export a range as a picture (EMF).

This module keeps all the COM‑interop gunk in one place so that higher‑level
code can stay tidy.  It works in two steps:

1.  Open Excel invisibly via win32com ─ returning the workbook handle.
2.  Either pull a cell value or copy an arbitrary range as a bitmap which is
    then written to a temporary **Enhanced Metafile** (.emf) that Word /
    PowerPoint can embed loss‑lessly.

The implementation is careful to clean Excel up even when things go wrong and
tries to survive the flaky clipboard behaviour some Windows builds exhibit.
"""

import os
import time
import tempfile
import logging
import traceback
from pathlib import Path

import pythoncom
import win32com.client as win32
import win32clipboard
import win32con
from PIL import ImageGrab  # Pillow

__all__ = ["get_value", "copy_range_as_emf"]

# Configure logging
logger = logging.getLogger(__name__)

# Excel constants
XL_PICTURE = -4147  # xlPicture
XL_SCREEN = 1       # xlScreen

# Initialize COM at module level for global COM access
# This helps ensure COM is initialized for all threads using this module
try:
    pythoncom.CoInitialize()
except:
    # It might already be initialized in this thread
    pass

# ──────────────────────────────
# internal helpers
# ──────────────────────────────

def _open_excel(path: str):
    """Return *(excel_app, workbook)* with Excel hidden (if allowed).

    Excel COM is initialized at the module level, so we don't need to
    initialize it again here.
    """
    try:
        xl = win32.Dispatch("Excel.Application")
        try:
            xl.Visible = False  # some builds may block this; that's OK
            xl.DisplayAlerts = False
        except AttributeError:
            logger.debug("Could not set Excel visibility or alerts")
            
        if not os.path.exists(path):
            raise FileNotFoundError(f"Excel file not found: {path}")
            
        wb = xl.Workbooks.Open(path, ReadOnly=True)
        return xl, wb
    except Exception as e:
        logger.error(f"Failed to open Excel: {e}")
        raise

def _safe_close(xl, wb):
    """Safely close Excel workbook and application."""
    try:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except:
                pass
    except:
        pass

def _clear_clipboard():
    """Safely clear the clipboard."""
    for _ in range(3):  # Try a few times if it fails
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            return True
        except Exception as e:
            logger.debug(f"Error clearing clipboard: {e}")
            time.sleep(0.1)
    return False

# ──────────────────────────────
# public helpers
# ──────────────────────────────

def get_value(path: str, sheet: str, cell: str):
    """Return the value of a single Excel cell."""
    xl, wb = None, None
    try:
        xl, wb = _open_excel(path)
        ws = wb.Worksheets(sheet)
        return ws.Range(cell).Value
    except Exception as e:
        raise ValueError(f"Error getting value from {sheet}!{cell}: {e}")
    finally:
        _safe_close(xl, wb)

def copy_range_as_emf(
    path: str,
    sheet: str,
    cell_range: str,
    timeout: float = 15.0,
    retry_count: int = 3,
) -> str:
    """Copy *cell_range* from *sheet* to the clipboard and persist as an EMF.

    Parameters
    ----------
    path : str
        Full path to the workbook.
    sheet : str
        Worksheet name.
    cell_range : str
        Excel A1 range (e.g. ``"B2:N10"``).
    timeout : float, default 15.0
        Seconds to wait for Excel to populate the clipboard.
    retry_count : int, default 3
        Number of times to retry clipboard operations if they fail.

    Returns
    -------
    str
        Absolute path to a temporary ``.emf`` file (caller may delete).

    Raises
    ------
    RuntimeError
        If the clipboard never receives an image of the copied range.
    """
    # Use a secure temporary file
    tmp_path = Path(tempfile.mktemp(suffix=".emf"))
    
    # Track if we succeeded
    success = False
    xl, wb = None, None
    
    for attempt in range(retry_count):
        try:
            logger.debug(f"Attempt {attempt+1} to copy range as EMF")
            
            # Open Excel
            xl, wb = _open_excel(path)
            
            # Access the worksheet and range
            sht = wb.Worksheets(sheet)
            sht.Activate()
            rng = sht.Range(cell_range)
            rng.Select()
            
            # Clear clipboard
            _clear_clipboard()
            
            # Copy as picture
            logger.debug(f"Copying range {cell_range} as picture")
            rng.CopyPicture(Appearance=XL_SCREEN, Format=XL_PICTURE)
            
            # Wait a bit to ensure Excel has time to process the copy operation
            time.sleep(0.5)
            
            # Process Windows messages to help Excel
            for _ in range(10):
                pythoncom.PumpWaitingMessages()
                time.sleep(0.05)
            
            # Poll clipboard until an image appears
            start = time.time()
            img = None
            
            while (time.time() - start) < timeout:
                try:
                    # Process Windows messages
                    pythoncom.PumpWaitingMessages()
                    
                    # Check if clipboard has image data
                    has_image = False
                    try:
                        win32clipboard.OpenClipboard()
                        has_image = win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB)
                        win32clipboard.CloseClipboard()
                    except:
                        pass
                    
                    if has_image:
                        img = ImageGrab.grabclipboard()
                        if img is not None:
                            break
                except Exception as e:
                    logger.debug(f"Error checking clipboard: {e}")
                
                time.sleep(0.1)
            
            # Close Excel as soon as we have the image or timed out
            _safe_close(xl, wb)
            xl, wb = None, None
            
            if img is None:
                logger.warning(f"Attempt {attempt+1}: Clipboard did not receive image")
                # Try next attempt
                continue
            
            # Save as EMF
            try:
                logger.debug(f"Saving image to {tmp_path}")
                img.save(str(tmp_path), "EMF")
                success = True
                return str(tmp_path)
            except Exception as e:
                logger.error(f"Failed to save EMF: {e}")
                try:
                    # If EMF fails, try saving as PNG as a fallback
                    png_path = str(tmp_path).replace(".emf", ".png")
                    img.save(png_path, "PNG")
                    logger.warning(f"Saved as PNG instead: {png_path}")
                    return png_path
                except:
                    # Continue to next attempt if both save methods fail
                    continue
                
        except Exception as e:
            logger.error(f"Attempt {attempt+1} failed: {str(e)}")
            logger.debug(traceback.format_exc())
            # Clean up Excel before next attempt
            _safe_close(xl, wb)
            xl, wb = None, None
    
    # Clean up the temporary file if we didn't succeed
    try:
        if tmp_path.exists():
            os.unlink(tmp_path)
    except:
        pass
            
    raise RuntimeError(f"Failed to copy range {cell_range} as EMF after {retry_count} attempts")