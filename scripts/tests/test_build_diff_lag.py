import openpyxl, pathlib
from build_diff_lag import write_diff_lag_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_diff_lag_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["DiffLag"]
    assert ws["A1"].value == "Row"
    assert ws["B1"].value == "Date"
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B11"].value == '=IF(Data!B11="","",Data!A11)'


def test_diff_lag_march_relocated_and_referencing_marchs_data_block():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["DiffLag"]
    # DiffLag's own offset for March: 1 * (8 + 2) = 10 -> B(2)+10=12='L'
    # Data's offset for March: 1 * (20 + 2) = 22 -> A(1)+22=23='W', B(2)+22=24='X'
    assert ws["L1"].value == "Date"
    assert ws["L2"].value == '=IF(Data!X2="","",Data!W2)'


def test_diff_lag_accumulates_across_months():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    write_diff_lag_sheet(wb, SOURCE_PATH, "March", num_rows=5)
    ws = wb["DiffLag"]
    assert ws["B2"].value is not None   # January's block survives
    assert ws["L2"].value is not None   # March's block was added


def test_diff_lag_month_label_in_gap_after_block():
    from multi_month import column_offset, DIFFLAG_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    label_col = shift_col("A", column_offset("January", DIFFLAG_BLOCK_WIDTH, GAP) + DIFFLAG_BLOCK_WIDTH)
    assert wb["DiffLag"][f"{label_col}1"].value == "January"
