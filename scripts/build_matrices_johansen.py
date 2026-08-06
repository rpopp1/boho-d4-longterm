import openpyxl
from formula_tools import extract_row_formulas, rewrite_constants, unwrap_scalar_array_formula
from config_map import CONSTANT_MAP, MATRICES_ONLY_CONSTANT_MAP


def write_matrices_sheet(wb, source_path: str) -> None:
    if "Matrices" in wb.sheetnames:
        del wb["Matrices"]
    ws = wb.create_sheet("Matrices")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Matrices"]
    for row in src.iter_rows(min_row=1, max_row=53, max_col=8):
        for cell in row:
            if cell.value is None:
                continue
            # MINOR 6: this plain iter_rows-based copy has no dynamic-array/
            # spill handling (unlike sheet_cloner.fill_sheet_from_pattern's
            # explicit NotImplementedError guard) -- Matrices!H7 is a live
            # ArrayFormula (a legacy CSE array formula, not a true spill;
            # verified its ref is just its own single-cell coordinate) that
            # would otherwise be written into the new sheet as an
            # ArrayFormula object rather than a plain formula string.
            ws[cell.coordinate] = unwrap_scalar_array_formula(cell.value, cell.coordinate)
    # Structural template-specific overrides:
    # Task 4's reconciled config_map.py confirmed the real window-length cell is
    # B3 (label "Window Length (rows)" sits at A3), not C3 — C3 is empty in the
    # live workbook. Use the corrected address.
    ws["B3"] = "=JohansenWindowLen"          # was the literal 63
    ws["B2"] = "=Johansen_Monthly!$K$3"       # unchanged reference target, now resolved dynamically there
    # MATRICES_ONLY_CONSTANT_MAP (e.g. "$B$3") is deliberately excluded from the
    # generic CONSTANT_MAP passed to the per-day sheet cloner (Task 7) because its
    # keys are too generic/short to safely blind-substitute across other sheets'
    # formulas. It is only ever merged in here, scoped to this sheet's own formulas.
    matrices_map = {**CONSTANT_MAP, **MATRICES_ONLY_CONSTANT_MAP}
    for row in ws.iter_rows(min_row=1, max_row=53, max_col=8):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = rewrite_constants(cell.value, matrices_map)


def write_johansen_sheet(wb, source_path: str) -> None:
    if "Johansen" in wb.sheetnames:
        del wb["Johansen"]
    ws = wb.create_sheet("Johansen")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Johansen"]
    for row in src.iter_rows(min_row=1, max_row=26, max_col=6):
        for cell in row:
            if cell.value is not None:
                ws[cell.coordinate] = unwrap_scalar_array_formula(cell.value, cell.coordinate)


def write_results_sheet(wb, source_path: str) -> None:
    if "Results" in wb.sheetnames:
        del wb["Results"]
    ws = wb.create_sheet("Results")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Results"]
    for row in src.iter_rows(min_row=1, max_row=11, max_col=2):
        for cell in row:
            if cell.value is not None:
                ws[cell.coordinate] = unwrap_scalar_array_formula(cell.value, cell.coordinate)
