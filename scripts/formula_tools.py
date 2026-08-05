import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter


def extract_row_formulas(path: str, sheet_name: str, row: int, max_col: int) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)
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
