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
