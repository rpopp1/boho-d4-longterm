import openpyxl
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants
from config_map import CONSTANT_MAP

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" in wb.sheetnames:
        del wb["Johansen_Monthly"]
    ws = wb.create_sheet("Johansen_Monthly")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)

    # Column A: plain per-row mirror of DiffLag's Date column. The live
    # workbook uses a FILTER() spill anchor here ("compact away blanks"), but
    # openpyxl cannot reliably reproduce Excel's dynamic-array spill semantics
    # (see Task 7 Addendum v2 / Task 15 finding), and DiffLag's Date column
    # has no interior blanks for real daily data, so a plain mirror is a safe,
    # robust substitution — column A is only consumed locally by column B's
    # own month-boundary comparison, within this same sheet.
    for i in range(num_rows):
        row = 2 + i
        ws[f"A{row}"] = f"=DiffLag!B{row}"

    # Column B: genuine per-row month-boundary detector — constructs a guarded
    # formula that handles boundary rows where DiffLag's formula returns "".
    # At the exact boundary row (last real data), DiffLag!B{row+1} returns ""
    # (a formula-computed empty string), and the old design's FILTER left such
    # cells genuinely blank. Our new column A mirrors DiffLag's formula output,
    # so A{row+1} becomes "" at the boundary. Without a guard, MONTH("") errors
    # in Excel. The guard IF(OR(A{row}="",A{row+1}=""),"",...)  matches the
    # old design's behavior: blank inputs → blank output, no error.
    for i in range(num_rows):
        row = 2 + i
        ws[f"B{row}"] = (
            f'=IF(OR(A{row}="",A{row+1}=""),"",IF(MONTH(A{row})<>MONTH(A{row+1}), '
            f'ROW(A{row})+ROW(DiffLag!B{row+1})-ROW(A{row})-1, ""))'
        )

    # C2/D2: fixed single-row lookups of the current live Johansen weight.
    # Written once, unshifted — these lack $ anchors but are NOT row-relative
    # (Johansen!D18/D19 are fixed eigenvector-result cells), so they must bypass
    # the per-row shifter entirely rather than being shifted per row.
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Johansen_Monthly"]
        ws["C2"] = rewrite_constants(src["C2"].value, CONSTANT_MAP)
        ws["D2"] = rewrite_constants(src["D2"].value, CONSTANT_MAP)
    finally:
        source_wb.close()

    # F2/G2 join H2/I2 as a uniformly manually-maintained walk-forward log —
    # header labels only, no formula, no fabricated history below the header.
    # (v1 tried to preserve F/G as FILTER spill anchors; Task 15's Dec-replay
    # validation found that openpyxl cannot make Excel treat a written
    # ArrayFormula as a true auto-resizing dynamic array rather than a
    # legacy, fixed-size CSE array that pads excess cells with #N/A — so this
    # redesign eliminates the spill formula entirely rather than attempting
    # unsupported OOXML metadata surgery.)
    ws["F2"] = "EndRow"
    ws["G2"] = "Date"
    ws["H2"] = "BOHO_Weight"
    ws["I2"] = "D4_Weight"

    ws["K2"] = "Active End Row"
    # Self-updating last-row lookup, replacing Dec's hardcoded "=F69": finds
    # the last row in column F (the month-end log) that is non-blank. F is
    # now manually maintained rather than auto-computed, which makes this
    # formula MORE robust, not less — no dependency on Excel's dynamic-array
    # file format at all.
    #
    # MAJOR 5 fix: scoped to F3:F<capacity> rather than the whole F:F column.
    # A whole-column LOOKUP(2,1/(F:F<>""),F:F) matches ANY non-blank cell in
    # F, including F2's own header text "EndRow" -- so on a freshly-built
    # book, before anyone has appended a real log row, K3 would return the
    # literal string "EndRow" instead of a genuine empty/error state. Row 3
    # is the log's first real data row (F2 is the header), and the sheet's
    # own row capacity (num_rows, matching every other engine sheet) is the
    # natural upper bound.
    last_log_row = num_rows + 1
    ws["K3"] = f'=LOOKUP(2,1/(F3:F{last_log_row}<>""),F3:F{last_log_row})'
