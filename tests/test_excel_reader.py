import time
import tempfile
import os
import win32com.client as win32
import win32clipboard
from PIL import ImageGrab

# ... keep get_value and _open_excel here ...

def copy_range_as_emf(path: str, sheet: str, cell_range: str, timeout=1.0) -> str:
    """
    Copy an Excel range to the clipboard as a picture and save it as an EMF.

    Parameters
    ----------
    timeout : float – seconds to keep polling the clipboard before failing.

    Returns
    -------
    str – absolute path to the temporary .emf file (caller may delete).
    """
    xl, wb = _open_excel(path)
    try:
        sht = wb.Worksheets(sheet)

        # xlPicture = -4147, xlScreen = 1
        sht.Range(cell_range).CopyPicture(Appearance=1, Format=-4147)

        # Poll clipboard until the image arrives (Excel can be slow)
        start = time.time()
        img = None
        while (time.time() - start) < timeout:
            img = ImageGrab.grabclipboard()
            if img is not None:
                break
            time.sleep(0.1)

        if img is None:
            raise RuntimeError(
                f"Clipboard did not contain an image after copying {cell_range}"
            )

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".emf")
        img.save(tmp.name, "EMF")
        return os.path.abspath(tmp.name)

    finally:
        wb.Close(False)
        xl.Quit()
