from pathlib import Path
from gslide.excel_reader import get_value

ASSETS = Path(__file__).parent / "assets"

def test_get_value():
    excel = ASSETS / "Sample.xlsx"
    assert get_value(str(excel), "PRP ValSum", "F22") == 123456