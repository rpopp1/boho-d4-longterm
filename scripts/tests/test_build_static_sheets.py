import openpyxl
from build_static_sheets import write_notes_sheet, write_data_sheet
from multi_month import MONTHS


def test_notes_sheet_describes_multi_month_workflow():
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    ws = wb["Notes"]
    text = "\n".join(str(ws[f"A{i}"].value) for i in range(1, ws.max_row + 1) if ws[f"A{i}"].value)
    assert "Config" in text
    assert "one column per contract month" in text.lower()


def test_data_sheet_january_lands_at_column_a():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    ws = wb["Data"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "BOHO"
    assert ws["E1"].value == "Settlement Pricing"
    assert ws["K1"].value == "Open Pricing"
    assert ws["T3"].value == "=T2/T3" or ws["T4"].value == "=T2/T3"  # ZL Gal formula, see Step 3


def test_data_sheet_march_lands_at_computed_offset():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    # March offset = 1 * (20 + 2) = 22 columns -> Date lands at column A+22 = "W"
    assert ws["W1"].value == "Date"
    assert ws["X1"].value == "BOHO"


def test_data_sheet_month_label_in_gap_after_each_block():
    # A plain text label in row 1 of the first gap column right after each
    # month's block, so it's visually obvious where one month's block ends
    # and the next begins (per the design spec's per-month-block layout).
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    jan_label_col = shift_col("A", column_offset("January", DATA_BLOCK_WIDTH, GAP) + DATA_BLOCK_WIDTH)
    mar_label_col = shift_col("A", column_offset("March", DATA_BLOCK_WIDTH, GAP) + DATA_BLOCK_WIDTH)
    assert ws[f"{jan_label_col}1"].value == "January"
    assert ws[f"{mar_label_col}1"].value == "March"


def test_data_sheet_accumulates_across_months():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    assert ws["A1"].value == "Date"   # January's block is still there
    assert ws["W1"].value == "Date"   # March's block was added, not a fresh sheet


def test_data_sheet_month_header_row():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "August")
    ws = wb["Data"]
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
    offset = column_offset("August", DATA_BLOCK_WIDTH, GAP)
    header_col = shift_col("A", offset)
    assert ws[f"{header_col}1"].value == "Date"
