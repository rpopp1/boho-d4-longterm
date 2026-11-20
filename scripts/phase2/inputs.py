"""Reads what the Phase 2 study needs from a COPY of the Knox workbook:
each month's first-block BOHO/D4 (the prices the signals use), its Settle
SBO/HO (the prices PnL uses), its Roll Prices rows and its stored monthly
Johansen weights. Never open Ryan's own workbooks with this; pass a copy."""

import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from multi_month import (  # noqa: E402
    DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, MONTHS, ROLL_PRICES_FIRST_ROW, column_offset, roll_prices_col,
)

# First-block / Settle-block positions inside a Data month block (1-based).
_DATA = {"date": 1, "boho": 2, "d4": 3, "sbo": 6, "ho": 7}
# Stored monthly weights inside a Johansen_Monthly month block.
_JM = {"end": 6, "date": 7, "w_boho": 8, "w_d4": 9}
JM_FIRST_ROW = 3


@dataclass
class MonthInputs:
    month: str
    rows: list
    dates: list
    boho: list
    d4: list
    sbo: list
    ho: list
    rolls: list = field(default_factory=list)   # (roll date, old SBO settle, old HO settle)
    stored: list = field(default_factory=list)  # (end row, month-end date, BOHO weight, D4 weight)


def _data_col(month: str, key: str) -> int:
    return _DATA[key] + column_offset(month, DATA_BLOCK_WIDTH)


def data_boho_col(month: str) -> str:
    return get_column_letter(_data_col(month, "boho"))


def johansen_cols(month: str) -> dict:
    off = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH)
    return {k: get_column_letter(v + off) for k, v in _JM.items()}


def _as_date(v):
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def load_inputs(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        data = list(wb["Data"].iter_rows(min_row=1, values_only=True))
        rp = list(wb["Roll Prices"].iter_rows(min_row=1, values_only=True))
        jm = list(wb["Johansen_Monthly"].iter_rows(min_row=1, values_only=True))
    finally:
        wb.close()
    out = {}
    for month in MONTHS:
        c = {k: _data_col(month, k) - 1 for k in _DATA}
        m = MonthInputs(month, [], [], [], [], [], [])
        for r in range(1, len(data)):
            row = data[r]
            d = _as_date(row[c["date"]]) if len(row) > c["date"] else None
            if d is None or d.year < 2015:
                continue
            m.rows.append(r + 1)
            m.dates.append(d)
            for k in ("boho", "d4", "sbo", "ho"):
                getattr(m, k).append(row[c[k]])
        rc = {h: column_index_from_string(roll_prices_col(month, h)) - 1
              for h in ("Roll Date", "SBO Settle", "HO Settle")}
        for row in rp[ROLL_PRICES_FIRST_ROW - 1:]:
            d = _as_date(row[rc["Roll Date"]]) if len(row) > rc["Roll Date"] else None
            if d is not None:
                m.rolls.append((d, row[rc["SBO Settle"]], row[rc["HO Settle"]]))
        jc = {k: column_index_from_string(v) - 1 for k, v in johansen_cols(month).items()}
        for row in jm[JM_FIRST_ROW - 1:]:
            if len(row) <= jc["w_d4"] or not isinstance(row[jc["end"]], (int, float)):
                break
            m.stored.append((int(row[jc["end"]]), _as_date(row[jc["date"]]),
                             row[jc["w_boho"]], row[jc["w_d4"]]))
        out[month] = m
    return out
