import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP
from formula_tools import rewrite_constants

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" in wb.sheetnames:
        del wb["Johansen_Monthly"]
    ws = wb.create_sheet("Johansen_Monthly")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)

    # Only column B (max_col=2) is a genuine per-row fill pattern; column A is a
    # spill anchor and is handled automatically by fill_sheet_from_pattern's
    # array-formula detection (Task 7 Addendum) within this same call.
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Johansen_Monthly",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=2,
        constant_map=CONSTANT_MAP,
    )

    # C2/D2: fixed single-row lookups of the current live Johansen weight.
    # Written once, unshifted — these lack $ anchors but are NOT row-relative
    # (Johansen!D18/D19 are fixed eigenvector-result cells), so they must bypass
    # the generic per-row shifter entirely rather than being shifted per row.
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Johansen_Monthly"]
        ws["C2"] = rewrite_constants(src["C2"].value, CONSTANT_MAP)
        ws["D2"] = rewrite_constants(src["D2"].value, CONSTANT_MAP)

        # F2/G2: spill anchors (month-end log compaction), same treatment as
        # column A — preserve as real ArrayFormula objects, written once.
        # ref must span more than one cell (see sheet_cloner.py's identical fix,
        # Task 7 Addendum) or Excel treats it as a single-cell legacy CSE array
        # formula instead of a real dynamic-array spill — this exact bug was
        # caught by Task 15's Dec-replay validation (#N/A cascading through
        # every downstream LOOKUP that depends on this log table).
        for col in ("F", "G"):
            cell = src[f"{col}2"]
            end_row = 2 + max(num_rows, 2) - 1
            ref = f"{col}2:{col}{end_row}"
            ws[f"{col}2"] = ArrayFormula(ref=ref, text=cell.value.text)
    finally:
        source_wb.close()

    # H2/I2: literal header labels for the log table (not formulas, matches
    # the live file exactly). Rows below are deliberately left empty — this
    # log is a manually-maintained, walk-forward history with no formula that
    # reproduces it (same category as the Data sheet's pasted price history).
    ws["H2"] = "BOHO_Weight"
    ws["I2"] = "D4_Weight"

    ws["K2"] = "Active End Row"
    # Self-updating last-row lookup, replacing Dec's hardcoded "=F69": finds
    # the last row in column F (the month-end log) that is non-blank. This
    # now resolves correctly because F is a real spill anchor, not empty.
    ws["K3"] = '=LOOKUP(2,1/(F:F<>""),F:F)'
