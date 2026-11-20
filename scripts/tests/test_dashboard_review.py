import pathlib
import sys

import openpyxl
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_dashboard import CALC_SHEET, build_calc_sheet, build_dashboard_sheet
from dashboard_metrics import matrix_lookup
from dashboard_review import REVIEW_ANCHORS


@pytest.fixture(scope="module")
def dash():
    wb = openpyxl.Workbook()
    build_calc_sheet(wb)
    build_dashboard_sheet(wb)
    return wb["Dashboard"]


def _dv_formula(ws, cell):
    return next(dv.formula1 for dv in ws.data_validations.dataValidation if cell in dv.sqref)


def test_controls_have_dropdowns_and_defaults(dash):
    assert (dash["B18"].value, dash["D18"].value, dash["F18"].value) == ("Portfolio", "All years", "Total PnL")
    assert _dv_formula(dash, "B18") == '"Portfolio,January,March,May,July,August,September,October,December"'
    assert _dv_formula(dash, "D18") == "YearOptions"
    assert _dv_formula(dash, "F18").startswith('"Total PnL,Sharpe,Sortino,')


def test_tiles_index_the_matrix(dash):
    labels = [dash.cell(20, c).value for c in range(1, 9)]
    assert labels == ["Total PnL", "Sharpe", "Sortino", "Max Drawdown", "Calmar", "Win Rate",
                      "Profit Factor", "Trade Count"]
    assert dash["B21"].value == "=" + matrix_lookup("2", f"{CALC_SHEET}!$Z$1", f"{CALC_SHEET}!$Z$2")


def test_year_table_rows_and_portfolio_column(dash):
    first = REVIEW_ANCHORS["year_table"] + 2
    assert dash.cell(first - 1, 10).value == "Portfolio"
    assert dash[f"A{first}"].value == f'=IF(2021<={CALC_SHEET}!$Y$2,2021,"")'
    assert "+(1)," in dash[f"J{first}"].value
    assert "+(2)," in dash[f"B{first}"].value


def test_heatmap_has_no_red_font_and_formula_midpoint(dash):
    first = REVIEW_ANCHORS["year_table"] + 2
    rules = [r for cf in dash.conditional_formatting if f"B{first}" in str(cf.sqref) for r in cf.rules]
    scale = next(r for r in rules if r.type == "colorScale")
    assert scale.colorScale.cfvo[1].type == "formula"
    assert scale.colorScale.cfvo[1].val == f"{CALC_SHEET}!$Z$7"
    fmts = [r.dxf.numFmt.formatCode for r in rules if r.dxf is not None and r.dxf.numFmt is not None]
    assert fmts and not any("[RED]" in f.upper() for f in fmts)
    for r in range(first, first + 12):
        for c in range(2, 11):
            assert "[RED]" not in dash.cell(r, c).number_format.upper()


def test_print_area_and_gridlines(dash):
    assert dash.print_area.startswith("'Dashboard'!$A$17")
    assert dash.sheet_view.showGridLines is False
    assert dash.freeze_panes == "A3"


def test_review_charts(dash):
    charts = dash._charts
    assert len(charts) == 4
    plotted = [s.val.numRef.f for ch in charts for s in ch.series]
    assert f"{CALC_SHEET}!Chart_SelEquity" in plotted
    assert f"{CALC_SHEET}!Chart_SelDD" in plotted
    assert any("$BG$" in f for f in plotted)
    bars = next(ch for ch in charts if any("$EX$" in s.val.numRef.f for s in ch.series))
    assert len(bars.series) == 2 and bars.overlap == 100
    assert any("$EY$" in s.val.numRef.f for s in bars.series)


def test_only_the_period_chart_names_remain():
    from build_dashboard import chart_range_names
    assert set(chart_range_names()) == {"Chart_SelDates", "Chart_SelEquity", "Chart_SelDD"}


def test_regime_highlight_fills_use_the_differential_background_color(dash):
    # Conditional-format solid fills only render from bgColor.
    rules = [r for cf in dash.conditional_formatting if "E6" in str(cf.sqref) for r in cf.rules]
    fills = [r.dxf.fill for r in rules if r.dxf is not None and r.dxf.fill is not None]
    assert fills and all(f.bgColor.rgb.endswith(("D55E00", "F4C7A8")) for f in fills)


def test_dollar_tiles_use_millions_format(dash):
    assert dash["A21"].number_format == '$#,##0.00,,"M";($#,##0.00,,"M")'
    assert dash["D21"].number_format == '$#,##0.00,,"M";($#,##0.00,,"M")'


def test_month_outline_skips_empty_year_rows(dash):
    first = REVIEW_ANCHORS["year_table"] + 2
    rules = [r for cf in dash.conditional_formatting for r in cf.rules if r.type == "expression"]
    assert any(f'AND($A{first}<>"",B${first - 1}=$B$18)' in r.formula[0] for r in rules)


def test_chart_axes_are_readable(dash):
    charts = dash._charts
    bars = next(ch for ch in charts if any("$EX$" in s.val.numRef.f for s in ch.series))
    assert bars.y_axis.number_format.formatCode == '[>=1000]#,##0;[<=-1000]-#,##0;0.00'
    dd = next(ch for ch in charts if any("Chart_SelDD" in s.val.numRef.f for s in ch.series))
    assert dd.x_axis.delete is True


def test_portfolio_column_has_its_own_color_scale(dash):
    first = REVIEW_ANCHORS["year_table"] + 2
    last = first + 11
    scales = {str(cf.sqref) for cf in dash.conditional_formatting for r in cf.rules if r.type == "colorScale"}
    assert f"B{first}:I{last}" in scales and f"J{first}:J{last}" in scales
    calc = dash.parent[CALC_SHEET]
    assert f"MEDIAN(Dashboard!$B${first}:$I${last})" in calc["Z7"].value
