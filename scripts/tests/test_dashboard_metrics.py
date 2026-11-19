import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from dashboard_metrics import (
    METRICS,
    _year_month_sharpe_formula,
    month_inception_formula,
    month_year_formula,
    portfolio_inception_formula,
)

# Formula text the pre-redesign Performance Since Inception table wrote
# (December row and Portfolio row), snapshotted 2026-11-19 so the redesign
# provably reuses the same since-inception math.
EXPECTED = json.loads((pathlib.Path(__file__).parent / "dashboard_expected_inception.json").read_text())
KEYS = ["pnl", "sharpe", "sortino", "calmar", "maxdd", "winrate", "avgwin", "avgloss", "pf", "trades"]


def test_metrics_order_and_labels():
    assert [k for k, _, _ in METRICS] == KEYS
    assert [label for _, label, _ in METRICS] == [
        "Total PnL", "Sharpe", "Sortino", "Calmar", "Max Drawdown",
        "Win Rate", "Avg $ Win", "Avg $ Loss", "Profit Factor", "Trade Count"]


def test_month_inception_formulas_match_current_table():
    for key in KEYS:
        assert month_inception_formula(key, "December") == EXPECTED["dec"][key], key


def test_portfolio_inception_formulas_match_current_table():
    for key in ["pnl", "sharpe", "sortino", "calmar", "maxdd"]:
        assert portfolio_inception_formula(key) == EXPECTED["port"][key], key


def test_month_year_formula_dispatch():
    assert month_year_formula("sharpe", "March", "$A12") == _year_month_sharpe_formula("March", "$A12")
