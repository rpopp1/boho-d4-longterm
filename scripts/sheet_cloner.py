from openpyxl.utils import get_column_letter, column_index_from_string
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants


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
    for i in range(num_rows):
        target_row = target_start_row + i
        row_delta = target_row - pattern_row
        for col_letter, value in pattern.items():
            if isinstance(value, str) and value.startswith("="):
                shifted = shift_formula_rows(value, row_delta)
                rewritten = rewrite_constants(shifted, constant_map)
                target_ws[f"{col_letter}{target_row}"] = rewritten
            else:
                target_ws[f"{col_letter}{target_row}"] = value
