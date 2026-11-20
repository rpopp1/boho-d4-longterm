import os
import pathlib
import sys
from datetime import date

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from phase2.back_adjust import back_adjust, boho_from_legs, roll_gaps
from phase2.inputs import MonthInputs, data_boho_col, johansen_cols, load_inputs

D = [date(2021, 1, d) for d in (4, 5, 6, 7, 8)]
COPY = os.environ.get("PHASE2_COPY")


def _m(dates, boho, rolls):
    n = len(dates)
    return MonthInputs(month="March", rows=list(range(2, 2 + n)), dates=dates, boho=boho, d4=[1.0] * n,
                       sbo=[0.0] * n, ho=[0.0] * n, rolls=rolls, stored=[])


def test_back_adjust_shifts_history_and_keeps_latest():
    out = back_adjust(D, [1.0, 1.1, 2.0, 2.1, 2.2], [(D[2], 0.85)])
    assert out == pytest.approx([1.85, 1.95, 2.0, 2.1, 2.2])
    assert out[2] - out[1] == pytest.approx(0.05)


def test_two_rolls_accumulate():
    p = [1.0, 1.1, 1.2, 1.3, 1.4]
    out = back_adjust(D, p, [(D[1], 0.5), (D[3], -0.2)])
    assert out[0] == pytest.approx(p[0] + 0.3)
    assert out[2] == pytest.approx(p[2] - 0.2)
    assert out[4] == p[4]


def test_roll_on_first_row_or_before_data_is_ignored():
    p = [1.0, 1.1, 1.2, 1.3, 1.4]
    assert back_adjust(D, p, [(D[0], 0.7)]) == p
    assert back_adjust(D, p, [(date(2020, 12, 15), 0.7)]) == p


def test_roll_date_missing_from_data_raises():
    m = _m(D, [1.0] * 5, [(date(2021, 1, 9), 50.0, 2.0)])
    with pytest.raises(ValueError, match=r"March.*2021-01-09"):
        roll_gaps(m)


def test_roll_gap_is_new_minus_old_boho():
    m = _m(D, [1.0, 1.1, 2.0, 2.1, 2.2], [(D[2], 40.0, 1.5)])
    assert roll_gaps(m) == [(D[2], pytest.approx(2.0 - (40.0 * 7.5 / 100 - 1.5)))]


@pytest.mark.skipif(not COPY, reason="PHASE2_COPY not set")
def test_boho_from_legs_matches_data_on_the_copy():
    for month, m in load_inputs(pathlib.Path(COPY)).items():
        ok = [abs(boho_from_legs(s, h) - b) < 1e-9 for s, h, b in zip(m.sbo, m.ho, m.boho)]
        bad = [m.dates[i] for i, v in enumerate(ok) if not v][:5]
        assert sum(ok) >= 0.99 * len(ok), (month, bad)


@pytest.mark.skipif(not COPY, reason="PHASE2_COPY not set")
def test_load_inputs_layout():
    inputs = load_inputs(pathlib.Path(COPY))
    assert inputs["January"].stored[0] == (20, date(2021, 1, 29), 0.6258677961771613, -0.7799291645453088)
    assert data_boho_col("March") == "X"
    assert johansen_cols("March")["w_boho"] == "U"
    assert inputs["December"].rows[0] == 2 and len(inputs["December"].rolls) >= 5
