import datetime
import openpyxl
from build_ho_front_month import (
    front_month_settle,
    parse_contract_columns,
    compute_front_month_series,
    write_front_month_sheet,
)

# Columns B/C/D carry Feb/Mar/Apr 2008 delivery, nearest first.
CONTRACTS = [(2, (2008, 2)), (3, (2008, 3)), (4, (2008, 4))]


def test_front_month_settle_picks_the_nearest_unexpired_delivery():
    row = (datetime.date(2008, 1, 30), 2.55, 2.49, 2.44)
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 1, 30)) == 2.55


def test_front_month_settle_rolls_once_the_delivery_month_starts():
    row = (datetime.date(2008, 2, 1), 2.53, 2.45, 2.41)
    # Feb delivery expired on the last business day of January, so the Feb
    # column is skipped even though it still carries a value.
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 2, 1)) == 2.45


def test_front_month_settle_ignores_a_stale_settle_on_an_expired_contract():
    # The Sep 2016 regression: HO U16 (Sep delivery) kept printing 1.4102 for
    # 18 business days after expiry, then a bad 3.7989. A value-driven roll
    # stayed on it all month; the calendar roll must be on V16 instead.
    contracts = [(2, (2016, 9)), (3, (2016, 10))]
    stale = (datetime.date(2016, 9, 15), 1.4102, 1.4162)
    bad_print = (datetime.date(2016, 9, 26), 3.7989, 1.4490)

    assert front_month_settle(contracts=contracts, row=stale,
                              as_of=datetime.date(2016, 9, 15)) == 1.4162
    assert front_month_settle(contracts=contracts, row=bad_print,
                              as_of=datetime.date(2016, 9, 26)) == 1.4490


def test_front_month_settle_rolls_forward_when_the_front_contract_is_empty():
    # Happens on the expiry day itself, when the source drops the settle.
    row = (datetime.date(2008, 1, 31), 0, 2.51, 2.47)
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 1, 31)) == 2.51

    row = (datetime.date(2008, 1, 31), None, 2.51, 2.47)
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 1, 31)) == 2.51


def test_front_month_settle_returns_none_when_nothing_is_live():
    row = (datetime.date(2008, 4, 1), 2.55, 2.49, 2.44)
    # Every delivery in CONTRACTS is April 2008 or earlier.
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 4, 1)) is None

    row = (datetime.date(2008, 1, 30), 0, None, 0)
    assert front_month_settle(row, CONTRACTS, datetime.date(2008, 1, 30)) is None


def _build_source_ws(wb, rows):
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([None, "HO", "HO", "HO"])
    ws.append([None, "G", "H", "J"])
    ws.append([None, 8, 8, 8])
    ws.append(["Time Series", "HO G08", "HO H08", "HO J08"])
    ws.append([None, "Settle", "Settle", "Settle"])
    ws.append([None] + ["HEATING OIL FUTURES"] * 3)
    for row in rows:
        ws.append(row)
    return ws


def test_parse_contract_columns_reads_code_and_year_in_delivery_order():
    wb = openpyxl.Workbook()
    ws = _build_source_ws(wb, [])
    assert parse_contract_columns(ws) == CONTRACTS


def test_compute_front_month_series_follows_the_calendar_roll():
    wb = openpyxl.Workbook()
    ws = _build_source_ws(wb, [
        # Feb delivery is front month through January...
        [datetime.date(2008, 1, 30), 2.55, 2.49, 2.44],
        [datetime.date(2008, 1, 31), 2.53, 2.51, 2.47],
        # ...then rolls to Mar on Feb 1, even though the Feb column still
        # carries a (now meaningless) settle.
        [datetime.date(2008, 2, 1), 2.52, 2.45, 2.41],
        # Nothing live once April delivery is reached: row is dropped.
        [datetime.date(2008, 4, 1), 2.52, 2.45, 2.41],
    ])

    assert compute_front_month_series(ws) == [
        (datetime.date(2008, 1, 30), 2.55),
        (datetime.date(2008, 1, 31), 2.53),
        (datetime.date(2008, 2, 1), 2.45),
    ]


def test_compute_front_month_series_skips_rows_without_a_date():
    wb = openpyxl.Workbook()
    ws = _build_source_ws(wb, [
        [datetime.date(2008, 1, 30), 2.55, 2.49, 2.44],
        [None, 2.54, 2.48, 2.43],
        ["total", 2.54, 2.48, 2.43],
    ])

    assert compute_front_month_series(ws) == [(datetime.date(2008, 1, 30), 2.55)]


def test_write_front_month_sheet_creates_sheet_with_headers_data_and_chart():
    wb = openpyxl.Workbook()
    series = [
        (datetime.date(2008, 1, 30), 2.55),
        (datetime.date(2008, 1, 31), 2.53),
        (datetime.date(2008, 2, 1), 2.45),
    ]

    write_front_month_sheet(wb, series)

    ws = wb["Front Month"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "Front-Month Settle"
    assert ws["A2"].value == datetime.date(2008, 1, 30)
    assert ws["B2"].value == 2.55
    assert ws["A4"].value == datetime.date(2008, 2, 1)
    assert ws["B4"].value == 2.45
    assert len(ws._charts) == 1


def test_write_front_month_sheet_replaces_rather_than_duplicates():
    wb = openpyxl.Workbook()
    write_front_month_sheet(wb, [(datetime.date(2008, 1, 30), 2.55)])
    write_front_month_sheet(wb, [(datetime.date(2008, 1, 31), 2.53)])

    assert wb.sheetnames.count("Front Month") == 1
    ws = wb["Front Month"]
    assert ws["A2"].value == datetime.date(2008, 1, 31)
    assert ws["B2"].value == 2.53
