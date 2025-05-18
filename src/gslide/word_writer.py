from __future__ import annotations
import win32com.client as win32


# ──────────────────────────────────────────────────────────────────────────────
#  TEXT → BOOKMARK
# ──────────────────────────────────────────────────────────────────────────────
def write_to_bookmark(
    doc_path: str,
    bookmark: str,
    text,
    *,
    readonly_template: bool = True,
) -> None:
    """
    Overwrite (or insert) *text* at a Word *bookmark* and re-create the bookmark
    so it survives future runs.

    Parameters
    ----------
    doc_path : str
        Path to the .docx / .docm you want to modify.
    bookmark : str
        Bookmark name inside the document.
    text : Any
        Gets converted to str() before writing.
    readonly_template : bool, default True
        Opens the file read-only (handy when it's a template on a share) and
        then saves over the same path, so the original stays clean but you still
        get an updated file on disk.  Set False if you need in-place editing.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # readonly=True means Word won't let us Save(); so we SaveAs to the same
    # path when we're done.
    doc = word.Documents.Open(
        doc_path,
        ReadOnly=readonly_template,
        AddToRecentFiles=False,
    )
    try:
        if not doc.Bookmarks.Exists(bookmark):
            raise ValueError(f"Bookmark '{bookmark}' not found in {doc_path}")

        rng = doc.Bookmarks(bookmark).Range
        rng.Text = str(text)
        doc.Bookmarks.Add(bookmark, rng)           # restore the bookmark

        if readonly_template:
            doc.SaveAs2(doc_path)                  # overwrites same file
        else:
            doc.Save()
    finally:
        doc.Close(False)
        word.Quit()


# Provide the new name you wanted for clarity, but keep the old one alive
write_value_to_bookmark = write_to_bookmark



# ──────────────────────────────────────────────────────────────────────────────
#  IMAGE → BOOKMARK   (unchanged – we’ll use this later)
# ──────────────────────────────────────────────────────────────────────────────
def paste_image_at_bookmark(
    doc_path: str,
    bookmark: str,
    img_path: str,
    width_pts: int | None = None,
) -> None:
    """
    Insert an image at *bookmark* and re-create the bookmark below it.

    Parameters
    ----------
    doc_path   : str – target Word document (.docx / .docm)
    bookmark   : str – bookmark name
    img_path   : str – path to the image file (PNG, EMF, etc.)
    width_pts  : int | None – desired width in Word points (72 pt = 1”).  None keeps native size.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(doc_path)
    try:
        if not doc.Bookmarks.Exists(bookmark):
            raise ValueError(f"Bookmark '{bookmark}' not found in {doc_path}")

        rng   = doc.Bookmarks(bookmark).Range
        shape = rng.InlineShapes.AddPicture(img_path, LinkToFile=False, SaveWithDocument=True)

        if width_pts is not None:
            shape.LockAspectRatio = True
            shape.Width = width_pts

        # re-add bookmark spanning the new picture
        doc.Bookmarks.Add(bookmark, shape.Range)
        doc.Save()
    finally:
        doc.Close(False)
        word.Quit()
