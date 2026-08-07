import openpyxl
from openpyxl.utils import get_column_letter
from formula_tools import rewrite_constants, shift_formula_columns
from config_map import constant_map_for_month
from multi_month import column_offset, mjr_offsets, DIFFLAG_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
from build_matrices_johansen import COMBINED_SHEET_NAME

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, month: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" not in wb.sheetnames:
        ws = wb.create_sheet("Johansen_Monthly")
    else:
        ws = wb["Johansen_Monthly"]

    self_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
    diff_lag_offset = column_offset(month, DIFFLAG_BLOCK_WIDTH, GAP)
    johansen_offset = mjr_offsets(month)["Johansen"]
    month_constant_map = constant_map_for_month(month)

    def col(letter):
        return shift_col(letter, self_offset)

    for i, h in enumerate(HEADERS, start=1):
        ws[f"{get_column_letter(i + self_offset)}1"] = h

    # Column A: plain per-row mirror of this month's DiffLag Date column.
    for i in range(num_rows):
        row = 2 + i
        ws[f"{col('A')}{row}"] = f"=DiffLag!{shift_col('B', diff_lag_offset)}{row}"

    # Column B: month-boundary detector, referencing this month's own column A
    # and this month's DiffLag block.
    for i in range(num_rows):
        row = 2 + i
        diff_lag_next = shift_col("B", diff_lag_offset)
        a_col, b_col = col("A"), col("B")
        ws[f"{b_col}{row}"] = (
            f'=IF(OR({a_col}{row}="",{a_col}{row + 1}=""),"",IF(MONTH({a_col}{row})<>MONTH({a_col}{row + 1}), '
            f'ROW({a_col}{row})+ROW(DiffLag!{diff_lag_next}{row + 1})-ROW({a_col}{row})-1, ""))'
        )

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Johansen_Monthly"]
        c2 = shift_formula_columns(
            rewrite_constants(src["C2"].value, month_constant_map),
            {"Johansen": (johansen_offset, COMBINED_SHEET_NAME)},
        )
        d2 = shift_formula_columns(
            rewrite_constants(src["D2"].value, month_constant_map),
            {"Johansen": (johansen_offset, COMBINED_SHEET_NAME)},
        )
        ws[f"{col('C')}2"] = c2
        ws[f"{col('D')}2"] = d2
    finally:
        source_wb.close()

    ws[f"{col('F')}2"] = "EndRow"
    ws[f"{col('G')}2"] = "Date"
    ws[f"{col('H')}2"] = "BOHO_Weight"
    ws[f"{col('I')}2"] = "D4_Weight"

    ws[f"{col('K')}2"] = "Active End Row"
    last_log_row = num_rows + 1
    f_col = col("F")
    ws[f"{col('K')}3"] = f'=LOOKUP(2,1/({f_col}3:{f_col}{last_log_row}<>""),{f_col}3:{f_col}{last_log_row})'

    ws[f"{shift_col('A', self_offset + JOHANSEN_MONTHLY_BLOCK_WIDTH)}1"] = month
