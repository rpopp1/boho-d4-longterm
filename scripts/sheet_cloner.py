import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants


def _get_array_formula_columns(source_path: str, source_sheet: str, pattern_row: int, max_col: int) -> set:
    """Column letters at pattern_row whose live-workbook cell is an ArrayFormula
    (a dynamic-array spill anchor), read directly (not via extract_row_formulas,
    which unwraps ArrayFormula to plain text and loses this distinction)."""
    wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        ws = wb[source_sheet]
        cols = set()
        for col in range(1, max_col + 1):
            cell = ws.cell(row=pattern_row, column=col)
            if isinstance(cell.value, ArrayFormula):
                cols.add(get_column_letter(col))
        return cols
    finally:
        wb.close()


def fill_sheet_from_pattern(
    target_ws,
    source_path: str,
    source_sheet: str,
    pattern_row: int,
    target_start_row: int,
    num_rows: int,
    max_col: int,
    constant_map: dict,
) -> None:
    pattern = extract_row_formulas(source_path, source_sheet, pattern_row, max_col)
    spill_cols = _get_array_formula_columns(source_path, source_sheet, pattern_row, max_col)
    # Re-read the raw ArrayFormula objects for spill columns so the written-out
    # cell keeps its real type (a plain string would not spill correctly in Excel).
    raw_array_formulas = {}
    if spill_cols:
        wb = openpyxl.load_workbook(source_path, data_only=False)
        try:
            ws = wb[source_sheet]
            for col_letter in spill_cols:
                cell = ws[f"{col_letter}{pattern_row}"]
                raw_array_formulas[col_letter] = cell.value
        finally:
            wb.close()

    for i in range(num_rows):
        target_row = target_start_row + i
        row_delta = target_row - pattern_row
        for col_letter, value in pattern.items():
            if col_letter in spill_cols:
                if i == 0:
                    original = raw_array_formulas[col_letter]
                    target_ws[f"{col_letter}{target_start_row}"] = ArrayFormula(
                        ref=f"{col_letter}{target_start_row}", text=original.text
                    )
                continue
            if isinstance(value, str) and value.startswith("="):
                shifted = shift_formula_rows(value, row_delta)
                rewritten = rewrite_constants(shifted, constant_map)
                target_ws[f"{col_letter}{target_row}"] = rewritten
            else:
                target_ws[f"{col_letter}{target_row}"] = value
