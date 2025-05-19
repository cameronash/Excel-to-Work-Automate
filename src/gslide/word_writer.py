# src/gslide/word_writer.py
from __future__ import annotations
import win32com.client as win32
import pythoncom

def _open_word(doc_path: str, readonly: bool = True):
    """
    Internal: launch Word and open the specified document once.
    Returns (word_app, document).
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
    Overwrite (or insert) *text* at a Word *bookmark*, warning on missing or error,
    but continuing without stopping the entire run.

    If no word_app/doc pair is supplied, this will open+close Word per bookmark;
    otherwise it reuses the provided session.
    """
    own = word_app is None
    if own:
        word_app, doc = _open_word(doc_path, readonly_template)

    try:
        if not doc.Bookmarks.Exists(bookmark):
            print(f"⚠️  Bookmark '{bookmark}' not found → skipping")
            return

        rng = doc.Bookmarks(bookmark).Range
        rng.Text = str(text)
        doc.Bookmarks.Add(bookmark, rng)
        print(f"✅  Wrote {text!r} into bookmark '{bookmark}'")

    except pythoncom.com_error as e:
        print(f"⚠️  Error on bookmark '{bookmark}': {e}")

    finally:
        if own:
            # save + tear down
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

# backwards‐compatible alias
write_value_to_bookmark = write_to_bookmark
