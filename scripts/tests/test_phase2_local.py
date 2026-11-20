import os
import pathlib
import statistics
import sys
from datetime import date, timedelta

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from phase2.inputs import MonthInputs
from phase2.local import local_weights, parse_window, z_series

COPY = os.environ.get("PHASE2_COPY")
N = 60
DATES = [date(2021, 1, 4) + timedelta(days=i) for i in range(N)]
ROWS = list(range(2, 2 + N))
BOHO = [1.0 + 0.01 * ((i * 7) % 11) + (0.5 if i >= 40 else 0.0) for i in range(N)]
D4 = [1.0 + 0.005 * ((i * 3) % 7) for i in range(N)]
WIN = {r: (max(2, r - 19), r) for r in ROWS}
STORED = [(r, DATES[r - 2], 0.6, -0.8) for r in ROWS if DATES[r - 2].day == 1 or r == 2]


def _m():
    return MonthInputs("May", ROWS, DATES, BOHO, D4, [0.0] * N, [0.0] * N, rolls=[], stored=STORED)


def test_parse_window():
    assert parse_window("=AVERAGE($I$21:I23)") == (21, 23)
    assert parse_window("=AVERAGE(I22:I41)") == (22, 41)
    assert parse_window(None) is None


def _plain_z(boho, t_row):
    a, b = WIN[t_row]
    spread = [0.6 * boho[r - 2] - 0.8 * D4[r - 2] for r in range(a, b + 1)]
    return (spread[-1] - statistics.mean(spread)) / statistics.stdev(spread)


def test_no_gaps_gives_the_plain_z_score():
    z = z_series(_m(), WIN, gaps=[], weights=STORED)
    assert z[ROWS.index(30)] == pytest.approx(_plain_z(BOHO, 30), abs=1e-12)


def test_roll_outside_a_window_leaves_it_unchanged():
    gaps = [(DATES[40], 0.5)]  # roll at row 42
    z0 = z_series(_m(), WIN, gaps=[], weights=STORED)
    z1 = z_series(_m(), WIN, gaps=gaps, weights=STORED)
    assert z1[ROWS.index(41)] == z0[ROWS.index(41)]          # window ends before the roll
    assert z1[ROWS.index(61)] == z0[ROWS.index(61)]          # window 42..61 starts at the roll row


def test_roll_inside_a_window_removes_the_jump_there_only():
    gaps = [(DATES[40], 0.5)]
    z1 = z_series(_m(), WIN, gaps=gaps, weights=STORED)
    adj = [p + (0.5 if i < 40 else 0.0) for i, p in enumerate(BOHO)]
    assert z1[ROWS.index(50)] == pytest.approx(_plain_z(adj, 50), abs=1e-12)


def test_local_weights_keep_stored_values_without_a_roll_in_the_window():
    m = _m()
    w = local_weights(m, gaps=[])
    assert w == STORED


@pytest.mark.skipif(not COPY, reason="PHASE2_COPY not set")
def test_python_z_reproduces_the_workbook_z_score():
    from phase2.local import load_signal_inputs
    for month, s in load_signal_inputs(pathlib.Path(COPY)).items():
        z = z_series(s.m, s.windows, gaps=[], weights=s.m.stored)
        diffs = [abs(a - b) for a, b in zip(z, s.z) if isinstance(a, float) and isinstance(b, (int, float))]
        assert len(diffs) > 1400, month
        assert max(diffs) < 1e-9, (month, max(diffs))
