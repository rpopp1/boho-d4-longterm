import openpyxl, pathlib
from build_matrices_johansen import write_matrices_sheet, write_johansen_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_matrices_window_length_uses_config_name():
    wb = openpyxl.Workbook()
    write_matrices_sheet(wb, SOURCE_PATH)
    ws = wb["Matrices"]
    assert ws["B3"].value == "=JohansenWindowLen"
    assert ws["B2"].value == "=Johansen_Monthly!$K$3"

def test_matrices_b4_formula_rewritten_via_merged_map():
    wb = openpyxl.Workbook()
    write_matrices_sheet(wb, SOURCE_PATH)
    ws = wb["Matrices"]
    b4 = ws["B4"].value
    assert "$B$3" not in b4
    assert "JohansenWindowLen" in b4
    assert b4 == "=MAX(3,$B$2-JohansenWindowLen+1)"

def test_johansen_sheet_replicated():
    wb = openpyxl.Workbook()
    write_johansen_sheet(wb, SOURCE_PATH)
    ws = wb["Johansen"]
    assert ws["B3"].value == "=Matrices!C50"
    assert ws["B8"].value == '=IF(OR(B3="",B6=""),"",B3+B6)'
