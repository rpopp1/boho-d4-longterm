import datetime
import sys
import pathlib

import openpyxl
from openpyxl.chart import LineChart, Reference

HEADER_ROWS = 6
DATE_COL = 1
FIRST_CONTRACT_COL = 2
CODE_ROW = 2
YEAR_ROW = 3
SOURCE_SHEET = "Sheet1"
OUTPUT_SHEET = "Front Month"

MONTH_CODES = {"F": 1, "G": 2, "H": 3, "J": 4, "K": 5, "M": 6,
               "N": 7, "Q": 8, "U": 9, "V": 10, "X": 11, "Z": 12}


def parse_contract_columns(ws, code_row=CODE_ROW, year_row=YEAR_ROW,
                           first_contract_col=FIRST_CONTRACT_COL):
    """[(column, (delivery_year, delivery_month)), ...] sorted nearest delivery first.

    Reads the month code from `code_row` and the two-digit year from `year_row`
    (e.g. 'U' / 16 -> September 2016). Columns missing either are skipped.
    """
    contracts = []
    for col in range(first_contract_col, ws.max_column + 1):
        code = ws.cell(code_row, col).value
        year = ws.cell(year_row, col).value
        if code is None or year is None:
            continue
        month = MONTH_CODES.get(str(code).strip().upper())
        if month is None:
            continue
        contracts.append((col, (2000 + int(year), month)))
    contracts.sort(key=lambda item: item[1])
    return contracts


def front_month_settle(row, contracts, as_of):
    """Settle of the nearest contract that has not expired as of `as_of`.

    HO terminates on the last business day of the month preceding delivery, so
    the front month on any date is the nearest contract delivering in a *later*
    calendar month. Selecting on the calendar rather than on "first column with
    a non-zero value" matters because the source keeps publishing a stale settle
    for some contracts after they expire (HO U16 held 1.4102 through September
    2016, then printed a bad 3.7989); a value-driven roll never leaves them.

    If the front contract has no settle, roll forward to the next delivery --
    never back to one that has already expired.
    """
    delivery_floor = (as_of.year, as_of.month)
    for col, delivery in contracts:
        if delivery <= delivery_floor:
            continue
        value = row[col - 1]
        if value not in (None, 0):
            return value
    return None


def compute_front_month_series(ws, header_rows=HEADER_ROWS, date_col=DATE_COL,
                                 first_contract_col=FIRST_CONTRACT_COL,
                                 code_row=CODE_ROW, year_row=YEAR_ROW):
    contracts = parse_contract_columns(ws, code_row, year_row, first_contract_col)
    series = []
    for row in ws.iter_rows(min_row=header_rows + 1, values_only=True):
        date = row[date_col - 1]
        if not isinstance(date, datetime.date):
            continue
        settle = front_month_settle(row, contracts, date)
        if settle is not None:
            series.append((date, settle))
    return series


def write_front_month_sheet(wb, series, sheet_name=OUTPUT_SHEET):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = "Date"
    ws["B1"] = "Front-Month Settle"
    for i, (date, settle) in enumerate(series, start=2):
        ws.cell(row=i, column=1, value=date).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=settle)

    last_row = len(series) + 1
    if last_row > 1:
        chart = LineChart()
        chart.title = "NY Harbor ULSD (HO) Front-Month Settle, 2008-Today"
        chart.y_axis.title = "$/Gallon"
        chart.x_axis.title = "Date"
        chart.width = 30
        chart.height = 14
        data = Reference(ws, min_col=2, min_row=1, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.series[0].smooth = False
        ws.add_chart(chart, "D2")

    return ws


def build(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    source_ws = wb[SOURCE_SHEET]
    series = compute_front_month_series(source_ws)
    write_front_month_sheet(wb, series)
    wb.save(path)
    return len(series)


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "HO Historical Pricing_2008-YTD.xlsx"
    n = build(pathlib.Path(target))
    print(f"Wrote {n} front-month rows to '{OUTPUT_SHEET}' in {target}")
