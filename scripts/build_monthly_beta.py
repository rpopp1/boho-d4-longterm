from openpyxl.utils import get_column_letter
from sheet_cloner import fill_sheet_from_pattern
from config_map import constant_map_for_month
from multi_month import column_offset, MONTHLY_BETA_BLOCK_WIDTH, DATA_BLOCK_WIDTH, GAP

HEADERS = ["Date", "BOHO", "D4", "BOHO_Δ", "D4_Δ", "MonthEnd", "Beta_120d", "Beta_Locked_Monthly"]


def write_monthly_beta_sheet(wb, source_path: str, month: str, num_rows: int = 3000) -> None:
    if "Monthly Beta" not in wb.sheetnames:
        ws = wb.create_sheet("Monthly Beta")
    else:
        ws = wb["Monthly Beta"]

    self_offset = column_offset(month, MONTHLY_BETA_BLOCK_WIDTH, GAP)
    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)

    for i, h in enumerate(HEADERS, start=1):
        ws[f"{get_column_letter(i + self_offset)}1"] = h

    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Monthly Beta",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=constant_map_for_month(month),
        column_shift_map={None: (self_offset, None), "Data": (data_offset, "Data")},
    )

    ws[f"{get_column_letter(1 + self_offset + MONTHLY_BETA_BLOCK_WIDTH)}1"] = month
