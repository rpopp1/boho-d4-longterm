from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" in wb.sheetnames:
        del wb["Johansen_Monthly"]
    ws = wb.create_sheet("Johansen_Monthly")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Johansen_Monthly",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=4,
        constant_map=CONSTANT_MAP,
    )
    ws["K1"] = "Active End Row"
    # Self-updating last-row lookup, replacing Dec's hardcoded "=F69":
    # finds the last row in column F (the month-end log) that is non-blank.
    ws["K3"] = '=LOOKUP(2,1/(F:F<>""),F:F)'
