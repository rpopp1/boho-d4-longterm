import pathlib
import sys

import openpyxl
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_dashboard import (
    CALC_SHEET,
    PADDED_LAST_ROW,
    TABLE_YEARS,
    build_calc_sheet,
    build_dashboard_sheet,
    chart_range_names,
    dynamic_range_name,
)
from multi_month import MONTHS


@pytest.fixture(scope="module")
def built():
    wb = openpyxl.Workbook()
    build_calc_sheet(wb)
    build_dashboard_sheet(wb)
    return wb


def test_calc_sheet_is_prefilled_to_the_padded_last_row(built):
    calc = built[CALC_SHEET]
    # Every row up to the padding limit carries the in-use flag and the
    # flag-gated year/daily formulas, so pasted rows need no rebuild.
    assert calc[f"P{PADDED_LAST_ROW}"].value.startswith("=IFERROR(IF(AND(")
    assert calc[f"A{PADDED_LAST_ROW}"].value.startswith(f"=IF($P{PADDED_LAST_ROW}=1,")
    assert calc[f"B{PADDED_LAST_ROW}"].value.startswith(f"=IF($P{PADDED_LAST_ROW}=1,")
    assert calc["Y1"].value == f"=SUM(P2:P{PADDED_LAST_ROW})"
    assert calc["Y2"].value == f"=MAX(A2:A{PADDED_LAST_ROW})"


def test_performance_stats_divide_by_rows_in_use_not_a_fixed_count(built):
    dash = built["Dashboard"]
    formulas = [c.value for row in dash.iter_rows() for c in row if isinstance(c.value, str)]
    # Since-inception figures (the year-filtered tables divide by that year's day count).
    sharpe = [f for f in formulas if "SQRT(252)" in f and f"{CALC_SHEET}!$A$2" not in f]
    assert sharpe and all(f"{CALC_SHEET}!$Y$1" in f for f in sharpe)
    assert not any("AVERAGE(Dashboard_Calc" in f or "STDEV.P(" in f for f in formulas)


def test_calendar_year_tables_reserve_rows_through_the_last_table_year(built):
    dash = built["Dashboard"]
    title = next(r for r in range(1, 200)
                 if dash.cell(r, 1).value == "Calendar-Year PnL by Contract Month (Risk-Adjusted Basis)")
    first = title + 2
    for j, year in enumerate(TABLE_YEARS):
        assert dash.cell(first + j, 1).value == f'=IF({year}<={CALC_SHEET}!$Y$2,{year},"")'
        # Each cell follows its row's year cell and stays blank until it fills.
        assert dash.cell(first + j, 2).value.startswith(f'=IF($A{first + j}="","",')
        assert f"=$A{first + j})" in dash.cell(first + j, 2).value


def test_by_year_chart_years_slide_with_the_latest_year(built):
    calc = built[CALC_SHEET]
    assert [calc.cell(1, c).value for c in range(18, 24)] == [f"=$Y$2-{k}" for k in range(5, -1, -1)]


def test_date_charts_plot_named_ranges_that_end_at_the_last_row_in_use(built):
    names = chart_range_names()
    assert set(names) == {"Chart_Dates", "Chart_Portfolio", "Chart_Selected"} | {f"Chart_{m}" for m in MONTHS}
    assert set(built[CALC_SHEET].defined_names) == set(names)
    assert dynamic_range_name("K") == (
        f"{CALC_SHEET}!$K$2:INDEX({CALC_SHEET}!$K$2:$K${PADDED_LAST_ROW},{CALC_SHEET}!$Y$1)")
    charts = built["Dashboard"]._charts
    plotted = {s.val.numRef.f for ch in charts for s in ch.series}
    assert {f"{CALC_SHEET}!Chart_{m}" for m in MONTHS} <= plotted
    assert f"{CALC_SHEET}!Chart_Selected" in plotted
