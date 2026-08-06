import openpyxl, pathlib
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
    assert ws["F2"].value == "EndRow"
    assert ws["G2"].value == "Date"
    assert ws["H2"].value == "BOHO_Weight"
    assert ws["I2"].value == "D4_Weight"

def test_column_a_is_plain_diff_lag_mirror_not_a_spill_anchor():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    from openpyxl.worksheet.formula import ArrayFormula
    assert not isinstance(ws["A2"].value, ArrayFormula)
    assert ws["A2"].value == "=DiffLag!B2"
    assert ws["A5"].value == "=DiffLag!B5"
    assert ws["A11"].value == "=DiffLag!B11"

def test_f_g_are_manually_maintained_headers_only_no_fabricated_data():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    for row in range(3, 12):
        assert ws[f"F{row}"].value is None
        assert ws[f"G{row}"].value is None

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

def test_column_b_still_correctly_filled_per_row():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["B2"].value == '=IF(MONTH(A2)<>MONTH(A3), ROW(A2)+ROW(DiffLag!B3)-ROW(A2)-1, "")'
    assert ws["B3"].value == '=IF(MONTH(A3)<>MONTH(A4), ROW(A3)+ROW(DiffLag!B4)-ROW(A3)-1, "")'
