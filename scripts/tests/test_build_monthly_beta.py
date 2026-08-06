import openpyxl, pathlib
from build_monthly_beta import write_monthly_beta_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_monthly_beta_header_and_fill():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Monthly Beta"]
    assert ws["A1"].value == "Date"
    assert ws["H1"].value == "Beta_Locked_Monthly"
    assert ws["A2"].value == "=Data!A2"
    assert ws["F2"].value == "=EOMONTH(A2,0)"
