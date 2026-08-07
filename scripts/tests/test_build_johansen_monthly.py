import openpyxl, pathlib
from build_johansen_monthly import write_johansen_monthly_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from multi_month import mjr_offsets, shift_col
    johansen_col = shift_col("D", mjr_offsets("January")["Johansen"])  # D(4)+10=14='N'
    assert ws["A2"].value == "=DiffLag!B2"
    assert ws["C2"].value == f"=Matrices_Johansen_Results!{johansen_col}18"
    assert ws["D2"].value == f"=Matrices_Johansen_Results!{johansen_col}19"
    assert ws["F2"].value == "EndRow"
    assert ws["K2"].value == "Active End Row"
    assert ws["K3"].value == '=LOOKUP(2,1/(F3:F11<>""),F3:F11)'


def test_march_relocated_and_referencing_marchs_diff_lag_block():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from multi_month import column_offset, mjr_offsets, JOHANSEN_MONTHLY_BLOCK_WIDTH, DIFFLAG_BLOCK_WIDTH, GAP, shift_col
    self_offset = column_offset("March", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)  # 1*(11+2)=13
    diff_lag_offset = column_offset("March", DIFFLAG_BLOCK_WIDTH, GAP)      # 1*(8+2)=10
    a_col = shift_col("A", self_offset)
    diff_lag_col = shift_col("B", diff_lag_offset)
    assert ws[f"{a_col}2"].value == f"=DiffLag!{diff_lag_col}2"

    johansen_col = shift_col("D", mjr_offsets("March")["Johansen"])
    c_col = shift_col("C", self_offset)
    assert ws[f"{c_col}2"].value == f"=Matrices_Johansen_Results!{johansen_col}18"

    k_col = shift_col("K", self_offset)
    assert ws[f"{k_col}2"].value == "Active End Row"


def test_january_month_label_in_gap_after_block():
    from multi_month import column_offset, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    label_col = shift_col("A", column_offset("January", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP) + JOHANSEN_MONTHLY_BLOCK_WIDTH)
    assert wb["Johansen_Monthly"][f"{label_col}1"].value == "January"


def test_headers_present():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "EndRow"
    assert ws["F2"].value == "EndRow"
    assert ws["G2"].value == "Date"
    assert ws["H2"].value == "BOHO_Weight"
    assert ws["I2"].value == "D4_Weight"


def test_column_a_is_plain_diff_lag_mirror_not_a_spill_anchor():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from openpyxl.worksheet.formula import ArrayFormula
    assert not isinstance(ws["A2"].value, ArrayFormula)
    assert ws["A2"].value == "=DiffLag!B2"
    assert ws["A5"].value == "=DiffLag!B5"
    assert ws["A11"].value == "=DiffLag!B11"


def test_f_g_are_manually_maintained_headers_only_no_fabricated_data():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    for row in range(3, 12):
        assert ws[f"F{row}"].value is None
        assert ws[f"G{row}"].value is None


def test_c_and_d_are_fixed_single_row_lookups_not_shifted():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from multi_month import mjr_offsets, shift_col
    johansen_col = shift_col("D", mjr_offsets("January")["Johansen"])
    assert ws["C2"].value == f"=Matrices_Johansen_Results!{johansen_col}18"
    assert ws["D2"].value == f"=Matrices_Johansen_Results!{johansen_col}19"
    # Must NOT be shifted to D19/D20 etc. at later rows — and must not be
    # duplicated at all, since this is a single current-value lookup.
    for row in range(3, 12):
        assert ws[f"C{row}"].value is None
        assert ws[f"D{row}"].value is None


def test_h_i_history_rows_left_empty_no_fabricated_data():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    for row in range(3, 12):
        assert ws[f"H{row}"].value is None
        assert ws[f"I{row}"].value is None


def test_column_b_still_correctly_filled_per_row():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["B2"].value == '=IF(OR(A2="",A3=""),"",IF(MONTH(A2)<>MONTH(A3), ROW(A2)+ROW(DiffLag!B3)-ROW(A2)-1, ""))'
    assert ws["B3"].value == '=IF(OR(A3="",A4=""),"",IF(MONTH(A3)<>MONTH(A4), ROW(A3)+ROW(DiffLag!B4)-ROW(A3)-1, ""))'


def test_column_b_has_blank_guard_against_boundary_value_errors():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    # At boundary rows where DiffLag!B{row+1} returns "", column A becomes ""
    # and MONTH("") errors in Excel. The guard IF(OR(A{row}="",A{row+1}=""),...)
    # prevents this, matching the old design's blank-cell behavior.
    b_formula = ws["B5"].value
    assert 'IF(OR(A5="",A6="")' in b_formula, "Column B must have blank guard clause"
    assert "MONTH(" in b_formula, "Column B must still check month boundaries"
