from pathlib import Path
import shutil
import win32com.client as win32

from gslide.excel_reader import get_value
from gslide.word_writer import write_to_bookmark

ASSETS = Path(__file__).parent / "assets"

def test_value_into_word(tmp_path):
    excel    = ASSETS / "Sample.xlsx"
    template = ASSETS / "Template.docm"
    work_doc = tmp_path / "out.docm"

    # copy template so the original stays untouched
    shutil.copy(template, work_doc)

    value = get_value(str(excel), "PRP ValSum", "F22")
    write_to_bookmark(str(work_doc), "MarketRent1", value)

    # reopen to verify
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(work_doc))
    try:
        assert str(value) in doc.Content.Text
    finally:
        doc.Close(False)
        word.Quit()
