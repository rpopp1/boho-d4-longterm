import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter, column_index_from_string
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants, shift_formula_columns


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
    column_shift_map: dict | None = None,
) -> None:
    pattern = extract_row_formulas(source_path, source_sheet, pattern_row, max_col)
    spill_cols = _get_array_formula_columns(source_path, source_sheet, pattern_row, max_col)
    if spill_cols:
        raise NotImplementedError(
            f"Pattern row {pattern_row} of sheet '{source_sheet}' contains "
            f"dynamic-array spill anchor formula(s) in column(s) {sorted(spill_cols)}. "
            "openpyxl cannot reliably reproduce Excel's dynamic-array spill "
            "semantics (no supported way to write the cm=\"1\" cell-metadata "
            "Excel needs to recognize a true auto-resizing spill, vs. a "
            "legacy fixed-size CSE array formula that pads excess cells with "
            "#N/A) — this exact failure mode was found and traced by Task 15's "
            "Dec-replay validation. Redesign the sheet builder to avoid spill "
            "formulas entirely (e.g. a plain per-row mirror formula, or a "
            "manually-maintained column) rather than special-casing this here "
            "— see Task 11 v2's redesign of Johansen_Monthly for the pattern."
        )

    self_offset = (column_shift_map or {}).get(None, (0, None))[0]

    for i in range(num_rows):
        target_row = target_start_row + i
        row_delta = target_row - pattern_row
        for col_letter, value in pattern.items():
            target_col_letter = get_column_letter(column_index_from_string(col_letter) + self_offset)
            if isinstance(value, str) and value.startswith("="):
                shifted = shift_formula_rows(value, row_delta)
                rewritten = rewrite_constants(shifted, constant_map)
                if column_shift_map:
                    rewritten = shift_formula_columns(rewritten, column_shift_map)
                target_ws[f"{target_col_letter}{target_row}"] = rewritten
            else:
                target_ws[f"{target_col_letter}{target_row}"] = value
