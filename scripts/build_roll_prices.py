"""Adds a 'Roll Prices' sheet to the SBOxHO pricing workbook.

Each month's block feed (BOHO v D4 Pricing - All Months) switches SBO/HO to
next year's same-month contract, and D4 to the next vintage, on the first
business day on/after that month's Roll Month/Roll Day. Knox's Position Model
needs the OLD contract's prices on that day to price a position held through
the roll-day close; this sheet looks them up, one row per roll year:

    Year | Roll Date | Old Contract | SBO Settle | HO Settle | D4 | SBO Open | HO Open

Roll Date is the first trading date in the All Months feed on/after the rule
date, so a holiday rule date resolves to the same row the feed rolls on.
Future years stay blank until their roll date has data.

Usage: python build_roll_prices.py <pricing workbook.xlsx>
Edits the workbook in place through an isolated, manual-calculation Excel
process; only the new sheet is calculated (the ICE-fed sheets are untouched).
"""

import sys
from pathlib import Path

from openpyxl.utils import get_column_letter

from excel_com import isolated_workbook, error_cell_count, reset_calc_mode_to_automatic, retry, replace_sheet
from multi_month import (
    MONTHS, CONTRACT_CODES, ROLL_PRICES_SHEET, ROLL_PRICES_HEADERS, ROLL_PRICES_FIRST_ROW,
    ROLL_PRICES_YEARS, roll_prices_col,
)

FEED = "'BOHO v D4 Pricing - All Months'"
# Roll table on the feed sheet: one row per month (MONTHS order) from row 3.
ROLL_MONTH_COL, ROLL_DAY_COL, ROLL_TABLE_FIRST_ROW = "FX", "FY", 3
CODE_ARRAY = '{"F","H","K","N","Q","U","V","Z"}'


def old_year_offset(month: str) -> int:
    """January rolls in December, while still holding NEXT year's F contract
    (the feed's Active Contract Year is YEAR+1 for January, YEAR+0 otherwise)."""
    return 1 if month == "January" else 0


def roll_prices_formulas(month: str, row: int, year: int) -> dict:
    """{header: formula or value} for `month`'s block on `row` (roll year `year`)."""
    c = {h: f"{roll_prices_col(month, h)}{row}" for h in ROLL_PRICES_HEADERS}
    code = CONTRACT_CODES[month]
    table_row = ROLL_TABLE_FIRST_ROW + MONTHS.index(month)
    old_year = f"({c['Year']}+{old_year_offset(month)})"
    rd = c["Roll Date"]
    dates = f"{FEED}!$E$3:$E$5000"

    def guarded(expr):
        return f'=IF({rd}="","",IFERROR({expr},""))'

    def settle(sheet):
        idx = (f"MATCH(MOD({old_year},100),'{sheet}'!$B$3:$BM$3,0)"
               f'+MATCH("{code}",{CODE_ARRAY},0)')
        return guarded(f"VLOOKUP({rd},'{sheet}'!$A$4:$BM$1048576,{idx},FALSE)")

    def open_(sheet, prefix, last_col):
        idx = f"MATCH(\"{prefix} \"&{c['Old Contract']},'{sheet}'!$J$4:${last_col}$4,0)+1"
        return guarded(f"VLOOKUP({rd},'{sheet}'!$I$4:${last_col}$1048576,{idx},FALSE)")

    return {
        "Year": year,
        "Roll Date": (
            f"=LET(t,DATE({c['Year']},{FEED}!${ROLL_MONTH_COL}${table_row},{FEED}!${ROLL_DAY_COL}${table_row}),"
            f"d,t+IF(WEEKDAY(t,2)>5,8-WEEKDAY(t,2),0),"
            f'r,MINIFS({dates},{dates},">="&d),IF(r=0,"",r))'
        ),
        "Old Contract": f'="{code}"&TEXT(MOD({old_year},100),"00")',
        "SBO Settle": settle("Soybean Oil Full History"),
        "HO Settle": settle("Heating Oil Full History"),
        # Same fallback as the feed's own D4 column: the latest vintage up to
        # the old one that has a price that day (January's next-year vintage
        # often has no price yet in December).
        "D4": guarded(
            f"LET(rw,MATCH({rd},'D4 Price History'!$H:$H,0),"
            f"v,INDEX('D4 Price History'!$I:$I,rw):INDEX('D4 Price History'!$I:$P,rw,"
            f"MATCH(\"D4 \"&{old_year},'D4 Price History'!$I$1:$P$1,0)),"
            f"LOOKUP(2,1/ISNUMBER(v),v))"
        ),
        "SBO Open": open_("Soybean Oil Dec History", "ZL", "BZ"),
        "HO Open": open_("Heating Oil Dec History", "HO", "BU"),
    }


def write_roll_prices_sheet(wb) -> None:
    """Create (or replace) the Roll Prices sheet in a COM workbook."""
    ws = replace_sheet(wb, ROLL_PRICES_SHEET)
    for month in MONTHS:
        first = roll_prices_col(month, ROLL_PRICES_HEADERS[0])
        ws.Range(f"{first}1").Value = f"{month} ({CONTRACT_CODES[month]}): old contract on its roll day"
        for h in ROLL_PRICES_HEADERS:
            ws.Range(f"{roll_prices_col(month, h)}2").Value = h
        for i, year in enumerate(ROLL_PRICES_YEARS):
            row = ROLL_PRICES_FIRST_ROW + i
            for h, value in roll_prices_formulas(month, row, year).items():
                cell = ws.Range(f"{roll_prices_col(month, h)}{row}")
                if isinstance(value, str) and value.startswith("="):
                    cell.Formula2 = value
                else:
                    cell.Value = value
        last = ROLL_PRICES_FIRST_ROW + len(ROLL_PRICES_YEARS) - 1
        ws.Range(f"{roll_prices_col(month, 'Roll Date')}{ROLL_PRICES_FIRST_ROW}:"
                 f"{roll_prices_col(month, 'Roll Date')}{last}").NumberFormat = "yyyy-mm-dd"
    ws.Range(f"A1:{get_column_letter(ws.UsedRange.Columns.Count)}2").Font.Bold = True
    ws.Calculate()


def main(path: Path) -> None:
    with isolated_workbook(path, manual_calc=True) as (_excel, wb):
        retry(lambda: write_roll_prices_sheet(wb))
        errors = error_cell_count(wb.Worksheets(ROLL_PRICES_SHEET))
        if errors:
            raise RuntimeError(f"{ROLL_PRICES_SHEET}: {errors} error cells; not saved")
        wb.Save()
    reset_calc_mode_to_automatic(path)
    print(f"{ROLL_PRICES_SHEET} sheet written to {path}")


if __name__ == "__main__":
    main(Path(sys.argv[1]).resolve())
