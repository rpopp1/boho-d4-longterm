import os
import pathlib
import sys

import numpy as np
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from phase2.inputs import load_inputs
from phase2.johansen import gate_report, johansen_weight

COPY = os.environ.get("PHASE2_COPY")


def _pair(n=80, seed=7):
    rng = np.random.default_rng(seed)
    d4 = 1.0 + np.cumsum(rng.normal(0, 0.02, n))
    boho = 0.8 * d4 + rng.normal(0, 0.01, n) + 0.3
    rows = range(2, 2 + n)
    return dict(zip(rows, boho)), dict(zip(rows, d4))


def test_weight_is_unit_length_and_sign_rule():
    b, d = _pair()
    wb, wd = johansen_weight(b, d, end_row=81)
    assert np.hypot(wb, wd) == pytest.approx(1.0, abs=1e-12)
    assert wd <= 0


def test_window_uses_the_63_rows_ending_at_end_row():
    # end_row 81 -> difference rows 19..81, which read prices 18..81.
    b, d = _pair()
    base = johansen_weight(b, d, end_row=81)
    b2 = dict(b); b2[17] += 5.0
    assert johansen_weight(b2, d, end_row=81) == base
    b3 = dict(b); b3[18] += 0.5
    assert johansen_weight(b3, d, end_row=81) != base


def test_short_history_window_starts_at_row_3():
    # end_row 20 -> rows 3..20 (T = 18), reading prices 2..20.
    b, d = _pair()
    full = johansen_weight(b, d, end_row=20)
    assert johansen_weight(b, d, end_row=20, window=18) == full
    assert johansen_weight(b, d, end_row=20, window=17) != full


@pytest.mark.skipif(not COPY, reason="PHASE2_COPY not set")
def test_gate_reproduces_stored_weights():
    reports = {m: gate_report(v) for m, v in load_inputs(pathlib.Path(COPY)).items()}
    for m, r in reports.items():
        print(m, r)
    assert all(r["max_abs_diff"] < 1e-9 for r in reports.values()), reports
