from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Row", "Date", "BOHO", "D4", "dBOHO", "dD4", "BOHO(-1)", "D4(-1)"]


def write_diff_lag_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "DiffLag" in wb.sheetnames:
        del wb["DiffLag"]
    ws = wb.create_sheet("DiffLag")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
