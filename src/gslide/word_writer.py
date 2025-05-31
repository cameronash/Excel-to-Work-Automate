from __future__ import annotations
import logging
import win32com.client as win32
import pythoncom

# Configure a simple logger
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def _open_word(doc_path: str, readonly: bool = True):
    """
    Internal: initialize COM, launch Word, and open the document.
    Returns (word_app, doc).
    """
    pythoncom.CoInitialize()
    word_app = win32.Dispatch("Word.Application")
    try:
        word_app.Visible = False
    except AttributeError:
        pass
    doc = word_app.Documents.Open(
        doc_path,
        ReadOnly=readonly,
        AddToRecentFiles=False,
    )
    return word_app, doc


def write_to_bookmark(
    doc_path: str,
    bookmark: str,
    text,
    *,
    word_app: win32.Dispatch | None = None,
    doc=None,
    readonly_template: bool = True,
) -> None:
    """
    Overwrite or insert *text* at a Word *bookmark*.
    Warn on missing bookmarks or COM errors but continue the run.
    """
    own_session = word_app is None
    if own_session:
        word_app, doc = _open_word(doc_path, readonly_template)

    try:
        if not doc.Bookmarks.Exists(bookmark):
            logging.warning(f"Bookmark '{bookmark}' not found → skipping")
            return

        # Insert the text
        rng = doc.Bookmarks(bookmark).Range
        rng.Text = str(text)

        # Re-create the bookmark (inserting text removes the original)
        doc.Bookmarks.Add(bookmark, rng)
        logging.info(f"Wrote {text!r} into bookmark '{bookmark}'")

    except pythoncom.com_error as e:
        logging.warning(f"COM error on bookmark '{bookmark}': {e}")
    except Exception as e:
        logging.warning(f"Unexpected error on bookmark '{bookmark}': {e}")

    finally:
        if own_session:
            # Save and clean up if we opened Word ourselves
            try:
                if readonly_template:
                    doc.SaveAs2(doc_path)
                else:
                    doc.Save()
            except Exception:
                pass
            try:
                doc.Close(False)
            except Exception:
                pass
            try:
                word_app.Quit()
            except Exception:
                pass


# Backwards-compatible alias
write_value_to_bookmark = write_to_bookmark
