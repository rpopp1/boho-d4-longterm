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
    EXPECTED_NA_RANGES,
    errors_outside,
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


def test_chart_names_follow_the_selected_period(built):
    names = chart_range_names()
    sel = {"Chart_SelDates", "Chart_SelEquity", "Chart_SelDD"}
    assert sel <= set(names)
    assert set(built[CALC_SHEET].defined_names) == set(names)
    assert names["Chart_SelEquity"] == (
        f"INDEX({CALC_SHEET}!$BD$2:$BD${PADDED_LAST_ROW},{CALC_SHEET}!$Z$4):"
        f"INDEX({CALC_SHEET}!$BD$2:$BD${PADDED_LAST_ROW},{CALC_SHEET}!$Z$5)")
    assert names["Chart_SelDates"].startswith("INDEX('Position Model - January'!$A$2:$A$3001,")
    assert "$BF$2:$BF$" in names["Chart_SelDD"]


def test_selected_series_columns_follow_the_period(built):
    calc = built[CALC_SHEET]
    assert calc["BD5"].value == "=IF($Z$3=0,AB5,AC5)"
    assert calc["BE5"].value == "=IF(AND($Z$3<>0,A5<>A4),BD5,MAX(BE4,BD5))"
    assert calc["BF5"].value == "=MIN(0,BD5-BE5)"


def test_overlay_gaps_instead_of_zero(built):
    calc = built[CALC_SHEET]
    assert "NA()" in calc["R100"].value and '""' not in calc["R100"].value


def test_selected_year_overlay_column(built):
    calc = built[CALC_SHEET]
    assert calc["BG1"].value == '=IF($Z$3=0,"",$Z$3)'
    assert calc["BG50"].value.startswith("=IF($Z$3=0,NA(),")


def test_unexpected_errors_ignore_the_expected_na_ranges():
    assert errors_outside(["R5", "W262", "BG2", "EX3", "A7", "BH9"], EXPECTED_NA_RANGES) == ["A7", "BH9"]
