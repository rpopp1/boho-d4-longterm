"""Applies the roll-day pricing fix to an existing Knox multi-month workbook.

1. Adds (or refreshes) a 'Roll Prices' sheet holding the pricing workbook's
   Roll Prices values (one row per roll day per month).
2. Rewrites each Position Model tab's price-change columns Z:AF with the
   roll-aware formulas from roll_fix.py -- after checking every existing cell
   still has its original plain day-over-day shape, so nothing hand-edited is
   overwritten silently.
3. Fully recalculates in an isolated Excel process, checks for error cells and
   saves.

Usage: python patch_knox_roll_fix.py <knox workbook.xlsx> <pricing workbook.xlsx>
Run it on a copy; it edits the Knox workbook in place.
"""

import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from excel_com import (
    isolated_workbook, error_cell_count, reset_calc_mode_to_automatic, retry, replace_sheet, sheet_names,
)
from multi_month import MONTHS, ROLL_PRICES_SHEET
from roll_fix import PRICE_CHANGE_COLS, price_change_formulas

EXCEL_EPOCH = datetime(1899, 12, 30)
FIRST_ROW = 3
LAST_ROW = 3001  # template-padded extent; live tabs hold formulas only down to their last data row


def read_roll_prices(pricing_path: Path) -> list:
    """The pricing workbook's Roll Prices sheet as cached values, row by row."""
    wb = openpyxl.load_workbook(pricing_path, read_only=True, data_only=True)
    try:
        return [list(r) for r in wb[ROLL_PRICES_SHEET].iter_rows(values_only=True)]
    finally:
        wb.close()


def _cell_value(v):
    """COM-safe value. Dates go in as Excel serial numbers: pywin32 treats a
    naive datetime as UTC and shifts it to local time (11/16 00:00 lands as
    11/16 05:00), which breaks the exact-date MATCH in the Position Model."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return (v - EXCEL_EPOCH).days + (v - EXCEL_EPOCH).seconds / 86400
    return v


def write_roll_prices_values(wb, rows: list) -> None:
    ws = replace_sheet(wb, ROLL_PRICES_SHEET)
    width = max(len(r) for r in rows)
    values = [[_cell_value(v) for v in r] + [""] * (width - len(r)) for r in rows]
    ws.Range(f"A1:{get_column_letter(width)}{len(values)}").Value = values
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            if isinstance(v, datetime):
                ws.Cells(r, c).NumberFormat = "yyyy-mm-dd"
    ws.Range(f"A1:{get_column_letter(width)}2").Font.Bold = True


def check_original_shape(ws, month: str) -> int:
    """Return the last row holding Z:AF formulas, after checking rows
    FIRST_ROW..that row all have the plain day-over-day formula the fix
    replaces and every row below it is blank. Raises otherwise."""
    block = ws.Range(f"Z{FIRST_ROW}:AF{LAST_ROW}").Formula
    cols = list(PRICE_CHANGE_COLS)
    last = max((FIRST_ROW + i for i, vals in enumerate(block) if any(vals)), default=FIRST_ROW - 1)
    bad = []
    for i, row_vals in enumerate(block):
        r = FIRST_ROW + i
        for col, actual in zip(cols, row_vals):
            expected = "=" + PRICE_CHANGE_COLS[col].format(r=r, p=r - 1) if r <= last else ""
            if actual != expected:
                bad.append(f"{col}{r}: {actual!r}")
    if bad:
        raise RuntimeError(f"Position Model - {month}: {len(bad)} unexpected Z:AF cells, e.g. {bad[:3]}")
    return last


def patch_position_model(ws, month: str) -> int:
    last = check_original_shape(ws, month)
    for col, formula in price_change_formulas(month, FIRST_ROW).items():
        # One formula string assigned to a whole column range: Excel shifts the
        # relative row references down the range, the Roll Prices ranges are absolute.
        ws.Range(f"{col}{FIRST_ROW}:{col}{last}").Formula2 = formula
    return last


def main(knox_path: Path, pricing_path: Path) -> None:
    rows = read_roll_prices(pricing_path)
    with isolated_workbook(knox_path, manual_calc=True) as (excel, wb):
        retry(lambda: write_roll_prices_values(wb, rows))
        for month in MONTHS:
            ws = wb.Worksheets(f"Position Model - {month}")
            last = retry(lambda ws=ws, m=month: patch_position_model(ws, m))
            print(f"Position Model - {month}: Z:AF rewritten for rows {FIRST_ROW}-{last}")
        excel.CalculateFullRebuild()
        errors = {name: error_cell_count(wb.Worksheets(name)) for name in sheet_names(wb)
                  if name.startswith("Position Model") or name.startswith("Dashboard")}
        print("error cells:", errors)
        wb.Save()
    reset_calc_mode_to_automatic(knox_path)
    print(f"Roll fix applied to {knox_path}")


if __name__ == "__main__":
    main(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve())
