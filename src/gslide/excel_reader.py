from __future__ import annotations
"""excel_emf_helpers.py ─ helpers for reading Excel and exporting ranges as vector EMF.

Changes in **v2** (fixing the *CopyEnhMetaFile* error)
-----------------------------------------------------
* Switch from the missing **win32api.CopyEnhMetaFile** to the
  always‑present *GDI* path: ``win32gui.GetEnhMetaFileBits`` + manual file
  write.  This works on every current pywin32 build.
* Imports ``win32gui`` directly; no optional import dance needed.
* Still falls back to PNG when Excel supplies a bitmap instead of EMF.

The rest of the logic (fresh Excel instance, CopyPicture, clipboard guard)
stays the same as in v1.
"""

import logging
import os
import tempfile
import time
import traceback
from pathlib import Path
from types import TracebackType
from typing import Optional, Type

import pythoncom
import win32clipboard
import win32com.client as win32
import win32con
import win32gui  # GDI helpers – contains GetEnhMetaFileBits/DeleteEnhMetaFile
from PIL import ImageGrab  # pillow

__all__ = ["get_value", "copy_range_as_emf"]

logger = logging.getLogger(__name__)

# Excel constants
XL_PICTURE = -4147  # xlPicture
XL_SCREEN = 1       # xlScreen
CF_ENHMETAFILE = 14  # winuser.h


# ---------------------------------------------------------------------------
# Utility context‑managers
# ---------------------------------------------------------------------------

class _OpenClipboard:  # pylint: disable=too-few-public-methods
    """Ensure the clipboard always gets closed."""

    def __enter__(self):
        win32clipboard.OpenClipboard()
        return win32clipboard

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc: Optional[BaseException],
        tb: Optional[TracebackType],
    ) -> bool:
        win32clipboard.CloseClipboard()
        return False  # don’t swallow exceptions


def _clear_clipboard(retries: int = 3) -> None:
    for _ in range(retries):
        try:
            with _OpenClipboard():
                win32clipboard.EmptyClipboard()
            return
        except Exception:
            time.sleep(0.1)
    logger.debug("Could not clear clipboard after %d retries", retries)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _open_excel(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(path, ReadOnly=True)
    return xl, wb


def _safe_close(xl, wb):  # pylint: disable=invalid-name
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if xl is not None:
            xl.Quit()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_value(path: str, sheet: str, cell: str):
    pythoncom.CoInitialize()
    xl = wb = None
    try:
        xl, wb = _open_excel(path)
        return wb.Worksheets(sheet).Range(cell).Value
    finally:
        _safe_close(xl, wb)


def copy_range_as_emf(
    path: str,
    sheet: str,
    cell_range: str,
    *,
    timeout: float = 15.0,
    retries: int = 3,
) -> str:
    """Copy *cell_range* to the clipboard as EMF; return a temp‑file path.

    If Excel provides only a bitmap the function transparently saves PNG
    instead, so callers still get a usable image path either way.
    """

    pythoncom.CoInitialize()
    last_exc: Optional[BaseException] = None

    for attempt in range(1, retries + 1):
        logger.debug("Attempt %d/%d", attempt, retries)
        xl = wb = None
        tmp_emf = Path(tempfile.mktemp(suffix=".emf"))
        try:
            xl, wb = _open_excel(path)
            rng = wb.Worksheets(sheet).Range(cell_range)

            _clear_clipboard()
            rng.CopyPicture(Appearance=XL_SCREEN, Format=XL_PICTURE)

            # Wait for Excel/COM
            end_time = time.time() + timeout
            emf_handle = None
            bitmap = None
            while time.time() < end_time:
                pythoncom.PumpWaitingMessages()
                with _OpenClipboard() as cb:
                    if cb.IsClipboardFormatAvailable(CF_ENHMETAFILE):
                        emf_handle = cb.GetClipboardData(CF_ENHMETAFILE)
                        break
                bitmap = ImageGrab.grabclipboard()
                if bitmap is not None:
                    break
                time.sleep(0.1)

            # Nothing arrived – try again
            if emf_handle is None and bitmap is None:
                raise RuntimeError("clipboard never received EMF/bitmap")

            # --------------- Vector path ------------------
            if emf_handle is not None:
                bits = win32gui.GetEnhMetaFileBits(emf_handle)
                with open(tmp_emf, "wb") as f:
                    f.write(bits)
                win32gui.DeleteEnhMetaFile(emf_handle)
                logger.info("Saved EMF → %s", tmp_emf)
                return str(tmp_emf)

            # --------------- Bitmap fallback --------------
            png_path = tmp_emf.with_suffix(".png")
            bitmap.save(png_path, "PNG")  # type: ignore[arg-type]
            logger.info("Saved PNG → %s", png_path)
            return str(png_path)

        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            logger.debug("%s\n%s", exc, traceback.format_exc())
        finally:
            _safe_close(xl, wb)

    raise RuntimeError(
        f"Failed to copy range {cell_range} as EMF after {retries} attempts"
    ) from last_exc
