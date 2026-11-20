import os
import pathlib
import sys
from datetime import date, timedelta

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from phase2.compare import match_trades, read_month, summary_stats, trades

COPY = os.environ.get("PHASE2_COPY")
D = [date(2021, 1, 4) + timedelta(days=i) for i in range(60)]


def test_trades_split_on_bs_runs():
    bs = [0, 1, 1, 0, 1, 1, 1, 0]
    v = [0, -1, -1, 0, 1, 1, 1, 1]
    bu = [0, 10, 25, 0, 5, -3, -8, 0]
    en = [0, 0, 1, 1, 0, 0, 0, 0]
    t = trades(D[:8], bs, v, bu, en)
    assert [(x["start"], x["end"], x["side"], x["pnl"], x["reason"]) for x in t] == [
        (D[1], D[2], -1, 25, "profit-take"), (D[4], D[6], 1, -8, "stop")]


def test_trade_ending_with_the_signal_is_a_signal_exit():
    t = trades(D[:5], [0, 1, 1, 0, 0], [0, 1, 1, 0, 0], [0, 4, 9, 0, 0], [0, 0, 0, 0, 0])
    assert t[0]["reason"] == "signal"


def test_match_trades_counts_shifted_not_added_removed():
    base = [dict(start=D[10], side=1, pnl=100.0), dict(start=D[40], side=-1, pnl=50.0)]
    test = [dict(start=D[11], side=1, pnl=100.0), dict(start=D[40], side=-1, pnl=70.0)]
    r = match_trades(base, test, D)
    assert (len(r["same"]), len(r["changed_pnl"]), len(r["shifted"]), len(r["added"]), len(r["removed"])) == \
        (0, 1, 1, 0, 0)


def test_far_apart_or_opposite_side_trades_are_added_and_removed():
    base = [dict(start=D[10], side=1, pnl=1.0)]
    test = [dict(start=D[10], side=-1, pnl=1.0), dict(start=D[30], side=1, pnl=1.0)]
    r = match_trades(base, test, D)
    assert (len(r["added"]), len(r["removed"])) == (2, 1)


def test_summary_stats_population_sharpe():
    s = summary_stats([1.0, -1.0, 1.0, -1.0, 2.0])
    assert s["pnl"] == 2.0
    assert s["sharpe"] == pytest.approx(0.4 / 1.2 * 252 ** 0.5)  # population SD of the sample = 1.2


@pytest.mark.skipif(not COPY, reason="PHASE2_COPY not set")
def test_summary_matches_dashboard_on_baseline():
    # The copy's Dashboard metric matrix (since inception, December row and
    # Portfolio row) is the reference: same BT data, same formulas.
    import openpyxl
    from dashboard_metrics import matrix_address
    wb = openpyxl.load_workbook(COPY, read_only=True, data_only=True)
    calc = wb["Dashboard_Calc"]
    want = {}
    for key, series in (("pnl", "December"), ("sharpe", "December"), ("pnl", "Portfolio"), ("sharpe", "Portfolio")):
        want[(key, series)] = calc[matrix_address(key, series, "All")].value
    wb.close()
    months = {m: read_month(pathlib.Path(COPY), m) for m in
              ("January", "March", "May", "July", "August", "September", "October", "December")}
    dec = summary_stats(months["December"]["bt"])
    assert dec["pnl"] == pytest.approx(want[("pnl", "December")], abs=0.01)
    assert dec["sharpe"] == pytest.approx(want[("sharpe", "December")], abs=1e-9)
    port = summary_stats([sum(x) for x in zip(*(m["bt"] for m in months.values()))])
    assert port["pnl"] == pytest.approx(want[("pnl", "Portfolio")], abs=0.01)
    assert port["sharpe"] == pytest.approx(want[("sharpe", "Portfolio")], abs=1e-9)
