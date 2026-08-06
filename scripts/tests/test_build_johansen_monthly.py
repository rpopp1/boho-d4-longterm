import openpyxl, pathlib
from openpyxl.worksheet.formula import ArrayFormula
from build_johansen_monthly import write_johansen_monthly_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_active_end_row_is_dynamic_not_hardcoded():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    formula = ws["K3"].value
    assert formula != "=F69"
    assert "LOOKUP(2,1/(" in formula
    assert "F:F" in formula  # resolves against the log table's own F column
    assert ws["K2"].value == "Active End Row"  # label sits at K2, not K1, per the live file

def test_headers_present():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "EndRow"
    assert ws["H2"].value == "BOHO_Weight"
    assert ws["I2"].value == "D4_Weight"

def test_spill_anchors_are_array_formulas_written_once():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    for col in ("A", "F", "G"):
        cell = ws[f"{col}2"].value
        assert isinstance(cell, ArrayFormula), f"{col}2 should be a spill anchor"
        assert "FILTER(" in cell.text
        # ref must span more than one cell (not a single-cell legacy CSE array)
        # or Excel treats it as a single-cell array instead of a dynamic-array spill
        assert cell.ref == f"{col}2:{col}11", f"{col}2's ref should span multiple rows, got {cell.ref}"
        # Must not be duplicated into later rows.
        for row in range(3, 12):
            assert ws[f"{col}{row}"].value is None, f"{col}{row} should be empty"

def test_c_and_d_are_fixed_single_row_lookups_not_shifted():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["C2"].value == "=Johansen!D18"
    assert ws["D2"].value == "=Johansen!D19"
    # Must NOT be shifted to D19/D20 etc. at later rows — and must not be
    # duplicated at all, since this is a single current-value lookup.
    for row in range(3, 12):
        assert ws[f"C{row}"].value is None
        assert ws[f"D{row}"].value is None

def test_h_i_history_rows_left_empty_no_fabricated_data():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    for row in range(3, 12):
        assert ws[f"H{row}"].value is None
        assert ws[f"I{row}"].value is None
