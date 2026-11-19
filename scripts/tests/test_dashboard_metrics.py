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


# ---------------------------------------------------------------------------
# Task 2: the metric matrix and control cells.
# ---------------------------------------------------------------------------
import openpyxl  # noqa: E402
import pytest  # noqa: E402

from build_dashboard import build_calc_sheet  # noqa: E402
from dashboard_metrics import matrix_address, matrix_lookup, portfolio_year_formula  # noqa: E402


@pytest.fixture(scope="module")
def calc():
    wb = openpyxl.Workbook()
    return build_calc_sheet(wb)


def test_matrix_address_layout():
    # Sharpe is metric block 1 (rows 12-21), December is series 9, 2024 is column FF.
    assert matrix_address("sharpe", "December", 2024) == "FF21"
    assert matrix_address("pnl", "Portfolio", "All") == "FB2"
    assert matrix_address("trades", "January", 2032) == "FN102"


def test_matrix_month_cells_reuse_existing_builders(calc):
    year_cell = calc[matrix_address("sharpe", "December", 2024)].value
    assert year_cell == ('=IF(FF$1>Dashboard_Calc!$Y$2,"",'
                         + month_year_formula("sharpe", "December", "FF$1").removeprefix("=") + ")")
    assert calc[matrix_address("sharpe", "December", "All")].value == month_inception_formula("sharpe", "December")
    assert calc["FC1"].value == 2021 and calc["FN1"].value == 2032 and calc["FB1"].value == "All"


def test_portfolio_cells(calc):
    assert calc[matrix_address("sharpe", "Portfolio", 2025)].value == (
        '=IF(FG$1>Dashboard_Calc!$Y$2,"",' + portfolio_year_formula("sharpe", "FG$1").removeprefix("=") + ")")
    assert calc[matrix_address("pnl", "Portfolio", "All")].value == portfolio_inception_formula("pnl")


def test_portfolio_trade_metrics_are_na_and_trades_sum(calc):
    assert calc[matrix_address("winrate", "Portfolio", "All")].value == '="n/a"'
    assert calc[matrix_address("winrate", "Portfolio", 2022)].value == '="n/a"'
    # Trade Count block is rows 100-110: Portfolio row 101, months 102-109.
    assert calc[matrix_address("trades", "Portfolio", 2022)].value == '=IF(FD$1>Dashboard_Calc!$Y$2,"",SUM(FD102:FD109))'
    assert calc[matrix_address("trades", "Portfolio", "All")].value == "=SUM(FB102:FB109)"


def test_portfolio_year_formulas_use_portfolio_columns():
    assert portfolio_year_formula("pnl", "FC$1") == \
        "=SUMPRODUCT((Dashboard_Calc!$A$2:$A$3001=FC$1)*Dashboard_Calc!$J$2:$J$3001)"
    assert "Dashboard_Calc!$BC$2:$BC$3001" in portfolio_year_formula("maxdd", "FC$1")
    assert "Dashboard_Calc!$BC$2:$BC$3001" in portfolio_year_formula("calmar", "FC$1")


def test_portfolio_year_drawdown_columns(calc):
    assert calc["BB3"].value == "=IF(A3<>A2,O3,MAX(BB2,O3))"
    assert calc["BC3"].value == "=IF(A3<>A2,MIN(0,O3-BB3),MIN(BC2,O3-BB3))"


def test_year_cells_blank_past_latest_year(calc):
    for key in KEYS:
        for series in ["Portfolio", "May"]:
            v = calc[matrix_address(key, series, 2030)].value
            if v != '="n/a"':
                assert v.startswith('=IF(FL$1>Dashboard_Calc!$Y$2,"",'), (key, series, v)


def test_year_options_follow_latest_year(calc):
    assert calc["EZ2"].value == "All years"
    assert calc["EZ3"].value == '=IF(2021<=$Y$2,2021,"")'
    assert calc["EZ14"].value == '=IF(2032<=$Y$2,2032,"")'


def test_control_cells(calc):
    assert calc["Z1"].value == "=MATCH(Dashboard!$B$18," + \
        '{"Portfolio","January","March","May","July","August","September","October","December"},0)'
    assert calc["Z2"].value == ('=IF(AND(ISNUMBER(Dashboard!$D$18),Dashboard!$D$18>=2021,'
                                'Dashboard!$D$18<=$Y$2),Dashboard!$D$18-2021+2,1)')
    assert calc["Z3"].value == "=IF($Z$2=1,0,Dashboard!$D$18)"
    assert calc["Z4"].value == "=IF($Z$3=0,1,MATCH($Z$3,$A$2:$A$3001,0))"
    assert calc["Z5"].value == "=IF($Z$3=0,$Y$1,$Z$4+COUNTIF($A$2:$A$3001,$Z$3)-1)"
    assert calc["Z6"].value.startswith("=MATCH(Dashboard!$F$18,{\"Total PnL\",")


def test_matrix_lookup_text():
    assert matrix_lookup("$Z$6", "$Z$1", "$Z$2") == \
        "INDEX(Dashboard_Calc!$FB$1:$FN$110,(($Z$6)-1)*11+1+($Z$1),$Z$2)"


def test_year_options_name_stops_at_the_latest_year(calc):
    # A workbook-level name so the Dashboard's Year dropdown has no blank rows.
    dn = calc.parent.defined_names["YearOptions"]
    assert dn.attr_text == "Dashboard_Calc!$EZ$2:INDEX(Dashboard_Calc!$EZ$2:$EZ$14,Dashboard_Calc!$Y$2-2021+2)"
