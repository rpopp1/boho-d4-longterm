import openpyxl, pathlib
from build_monthly_beta import write_monthly_beta_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_monthly_beta_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    ws = wb["Monthly Beta"]
    assert ws["A1"].value == "Date"
    assert ws["H1"].value == "Beta_Locked_Monthly"
    assert ws["A2"].value == "=Data!A2"


def test_monthly_beta_march_relocated():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "March", num_rows=5)
    ws = wb["Monthly Beta"]
    # Monthly Beta's own offset for March: 1 * (8 + 2) = 10 -> A(1)+10=11='K'
    # Data's offset for March: 22 -> A(1)+22=23='W'
    assert ws["K1"].value == "Date"
    assert ws["K2"].value == "=Data!W2"


def test_monthly_beta_month_label_in_gap_after_block():
    from multi_month import column_offset, MONTHLY_BETA_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    label_col = shift_col("A", column_offset("January", MONTHLY_BETA_BLOCK_WIDTH, GAP) + MONTHLY_BETA_BLOCK_WIDTH)
    assert wb["Monthly Beta"][f"{label_col}1"].value == "January"
