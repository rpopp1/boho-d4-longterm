"""Window-local roll adjustment: the roll gap is removed only inside the
calculation windows that contain that roll, so every other value stays
identical to the Sandbox.

- Z-score (Position Model I:L): for each row, the spread values in that row's
  own AVERAGE/STDEV.S window get the gap of every roll inside the window added
  to the rows before the roll. Each row keeps its own month-end weight.
- Hedge weights: a month-end's 63-row Johansen window that contains a roll
  is re-estimated with the gap added to the rows before the roll; the rest
  keep their stored values.
- Volatility (CV, daily BOHO change): only the roll day's change loses the gap.
"""

import re
import statistics
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

from phase2.inputs import MonthInputs, load_inputs
from phase2.johansen import FIRST_ROW, WINDOW, johansen_weight
from multi_month import MONTHS

_WIN_RE = re.compile(r"AVERAGE\(\$?I\$?(\d+):\$?I\$?(\d+)\)")


def parse_window(formula):
    """(first row, last row) of a Position Model J formula's AVERAGE range."""
    if not isinstance(formula, str):
        return None
    m = _WIN_RE.search(formula)
    return (int(m.group(1)), int(m.group(2))) if m else None


@dataclass
class SignalInputs:
    m: MonthInputs
    windows: dict  # row -> (first, last) of its Z-score window
    z: list        # the workbook's Z-score (L) per row, as saved


def load_signal_inputs(path: Path) -> dict:
    inputs = load_inputs(path)
    j, l_ = column_index_from_string("J"), column_index_from_string("L")
    wf = openpyxl.load_workbook(path, read_only=True)
    wv = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = {}
        for month in MONTHS:
            m = inputs[month]
            last = m.rows[-1]
            jf = [r[0] for r in wf[f"Position Model - {month}"].iter_rows(min_row=1, max_row=last, min_col=j, max_col=j, values_only=True)]
            lv = [r[0] for r in wv[f"Position Model - {month}"].iter_rows(min_row=1, max_row=last, min_col=l_, max_col=l_, values_only=True)]
            out[month] = SignalInputs(m, {r: parse_window(jf[r - 1]) for r in m.rows}, [lv[r - 1] for r in m.rows])
        return out
    finally:
        wf.close()
        wv.close()


def _weight_lookup(m: MonthInputs, weights: list) -> list:
    """Each row's (BOHO, D4) weight: the latest month-end on or before the
    day before its date (Position Model I's LOOKUP(A-1, ...)), or None."""
    ends = [(d, wb, wd) for _e, d, wb, wd in weights]
    out, k = [], -1
    for d in m.dates:
        while k + 1 < len(ends) and ends[k + 1][0] <= d - timedelta(days=1):
            k += 1
        out.append((ends[k][1], ends[k][2]) if k >= 0 else None)
    return out


def _roll_rows(m: MonthInputs, gaps: list) -> list:
    index = {d: r for d, r in zip(m.dates, m.rows)}
    return [(index[d], g) for d, g in gaps if d in index]


def z_series(m: MonthInputs, windows: dict, gaps: list, weights: list) -> list:
    """Python Z-score per row (None where the workbook's would be an error)."""
    w = _weight_lookup(m, weights)
    pos = {r: i for i, r in enumerate(m.rows)}
    rolls = _roll_rows(m, gaps)
    out = []
    for i, r in enumerate(m.rows):
        win = windows.get(r)
        if not win:
            out.append(None)
            continue
        a, b = win
        inside = [(rr, g) for rr, g in rolls if a < rr <= b]
        vals = []
        for row in range(a, b + 1):
            k = pos.get(row)
            if k is None or w[k] is None or not isinstance(m.boho[k], (int, float)):
                vals = None
                break
            adj = sum(g for rr, g in inside if row < rr)
            vals.append(w[k][0] * (m.boho[k] + adj) + w[k][1] * m.d4[k])
        if not vals or len(vals) < 2:
            out.append(None)
            continue
        sd = statistics.stdev(vals)
        out.append((vals[-1] - sum(vals) / len(vals)) / sd if sd else None)
    return out


def local_weights(m: MonthInputs, gaps: list) -> list:
    """Stored month-end weights, re-estimated only where the Johansen window
    contains a roll (gap added to the window's rows before the roll)."""
    rolls = _roll_rows(m, gaps)
    base_b = dict(zip(m.rows, m.boho))
    d4 = dict(zip(m.rows, m.d4))
    out = []
    for end, d, wb, wd in m.stored:
        start = max(FIRST_ROW, end - WINDOW + 1)
        inside = [(rr, g) for rr, g in rolls if start <= rr <= end]
        if not inside:
            out.append((end, d, wb, wd))
            continue
        adj = {row: p + sum(g for rr, g in inside if row < rr) for row, p in base_b.items() if start - 1 <= row <= end}
        out.append((end, d, *johansen_weight({**base_b, **adj}, d4, end)))
    return out


def roll_day_changes(m: MonthInputs, gaps: list) -> dict:
    """{row: daily BOHO change with the gap removed} for each roll day."""
    pos = {r: i for i, r in enumerate(m.rows)}
    out = {}
    for rr, g in _roll_rows(m, gaps):
        k = pos[rr]
        if k > 0:
            out[rr] = m.boho[k] - m.boho[k - 1] - g
    return out


def rows_in_roll_windows(m: MonthInputs, windows: dict, gaps: list) -> list:
    rolls = [rr for rr, _g in _roll_rows(m, gaps)]
    return [r for r in m.rows if windows.get(r) and any(windows[r][0] < rr <= windows[r][1] for rr in rolls)]
