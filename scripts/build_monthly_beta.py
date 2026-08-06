from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Date", "BOHO", "D4", "BOHO_Δ", "D4_Δ", "MonthEnd", "Beta_120d", "Beta_Locked_Monthly"]


def write_monthly_beta_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Monthly Beta" in wb.sheetnames:
        del wb["Monthly Beta"]
    ws = wb.create_sheet("Monthly Beta")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Monthly Beta",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
