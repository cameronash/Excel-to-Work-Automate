import pytest; pytest.skip("Image workflow postponed", allow_module_level=True)


from pathlib import Path
import shutil
import win32com.client as win32

from gslide.excel_reader import copy_range_as_emf
from gslide.word_writer  import paste_image_at_bookmark

ASSETS = Path(__file__).parent / "assets"

def test_range_as_image(tmp_path):
    excel    = ASSETS / "Sample.xlsx"
    template = ASSETS / "Template.docm"
    work_doc = tmp_path / "out.docm"

    shutil.copy(template, work_doc)

    img_path = copy_range_as_emf(str(excel), "PRP ValSum", "B2:N10")
    paste_image_at_bookmark(str(work_doc), "RangePic1", img_path, width_pts=350)

    # reopen Word to verify an image exists at that bookmark
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(work_doc))
    try:
        assert doc.Bookmarks.Exists("RangePic1")
        # the image is now the first InlineShape
        assert doc.InlineShapes.Count >= 1
    finally:
        doc.Close(False)
        word.Quit()
