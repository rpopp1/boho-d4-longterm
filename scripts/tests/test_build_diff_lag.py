import openpyxl, pathlib
from build_diff_lag import write_diff_lag_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_diff_lag_header_and_fill():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["DiffLag"]
    assert ws["A1"].value == "Row"
    assert ws["B1"].value == "Date"
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B11"].value == '=IF(Data!B11="","",Data!A11)'
