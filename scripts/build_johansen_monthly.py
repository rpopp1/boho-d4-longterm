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

    # Column B: genuine per-row month-boundary detector — extracted from the
    # live source and shifted, same as any ordinary per-day formula. Extracted
    # directly (not via fill_sheet_from_pattern) because that function would
    # also try to touch column A from source, where it's still the old FILTER
    # spill anchor we're deliberately replacing above.
    b_pattern = extract_row_formulas(source_path, "Johansen_Monthly", row=2, max_col=2)["B"]
    for i in range(num_rows):
        row = 2 + i
        shifted = shift_formula_rows(b_pattern, row - 2)
        ws[f"B{row}"] = rewrite_constants(shifted, CONSTANT_MAP)

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
    ws["K3"] = '=LOOKUP(2,1/(F:F<>""),F:F)'
