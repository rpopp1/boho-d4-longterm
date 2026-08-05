import re

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter


def extract_row_formulas(path: str, sheet_name: str, row: int, max_col: int) -> dict[str, object]:
    """Extract formulas and values from a row in an Excel workbook.

    Returns a dictionary mapping column letters to their cell values.
    Formula cells return formula strings, array formulas are unwrapped to their text,
    and non-formula cells (e.g., literal numbers, strings) are returned as-is.
    Empty cells are omitted from the result.
    """
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        ws = wb[sheet_name]
        result = {}
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value
            if value is None:
                continue
            if isinstance(value, ArrayFormula):
                value = value.text
            result[get_column_letter(col)] = value
        return result
    finally:
        wb.close()


# Matches an optional sheet-qualifier, optional $ + column letters, optional $ + row digits.
# Full-column refs like G:G (no digits) simply won't match the row-digit group and pass through.
# The leading negative lookbehind prevents matching in the middle of a longer identifier/defined
# name (e.g. the "r1" tail of "ZThreshold_Tier1" must never be treated as a cell reference).
_CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?P<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<colabs>\$)?(?P<col>[A-Za-z]{1,3})"
    r"(?P<rowabs>\$)?(?P<row>\d+)"
    r"(?![A-Za-z0-9_(])"
)


def shift_formula_rows(formula: str, row_delta: int) -> str:
    def _replace(match: "re.Match") -> str:
        if match.group("rowabs"):
            return match.group(0)  # absolute row: never shift
        sheet = match.group("sheet") or ""
        colabs = match.group("colabs") or ""
        col = match.group("col")
        new_row = int(match.group("row")) + row_delta
        return f"{sheet}{colabs}{col}{new_row}"

    return _CELL_REF_RE.sub(_replace, formula)
