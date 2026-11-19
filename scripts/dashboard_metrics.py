"""Formula builders for every Dashboard performance metric (since inception
and per calendar year, per contract month and for the Portfolio), plus the
Dashboard_Calc layout constants they depend on."""

from openpyxl.utils import get_column_letter

from config_map import CONFIG_MONTH_COLUMNS
from multi_month import MONTHS

# Full template-padded row extent every Position Model tab shares (real data
# currently ends around row 1392/1393 depending on month; everything past
# that is genuinely blank). Using this full extent with ISNUMBER-gated
# aggregation means every formula below auto-extends as more real trading
# days get appended in the future, with no hardcoded per-month boundary --
# deliberately avoiding this project's known "hardcoded Dec-row-boundary"
# defect class.
PADDED_LAST_ROW = 3001

# Auto-extension. Dashboard_Calc is pre-filled with formulas down to
# PADDED_LAST_ROW (about the end of 2032 at ~252 trading days a year), so
# rows pasted after a build are picked up with no rebuild. The two earlier
# designs both went stale: a hand-set last row (the 2026-08-14 build, found
# ~3 months stale on 2026-11-16) and a last row computed at build time
# (stopped at 11/16 while data reached 11/18). Column P flags the rows where
# all 8 months hold a real date; unused rows carry zeros, Y1 counts the rows
# in use and every mean/variance divides by it, so unused rows change
# nothing. Charts plot named ranges that end at row 1 + Y1.
CALC_SHEET = "Dashboard_Calc"
VALID_COL = "P"
ROWS_IN_USE = f"{CALC_SHEET}!$Y$1"
LATEST_YEAR = f"{CALC_SHEET}!$Y$2"

# Calendar-Year tables: one row per year from the start of the data through
# 2032 (where row 3001 runs out). A year's row stays blank until data for it
# exists. The PnL-by-year chart shows the latest CHART_YEAR_COUNT years.
FIRST_YEAR = 2021
LAST_TABLE_YEAR = 2032
TABLE_YEARS = list(range(FIRST_YEAR, LAST_TABLE_YEAR + 1))
CHART_YEAR_COUNT = 6
MAX_DAY_INDEX = 260  # comfortably above any year's ~252 trading days

# Dashboard_Calc column layout for the per-month, year-reset cumulative/peak/
# drawdown helper columns (see build_calc_sheet) -- used by the Year x Month
# Max Drawdown/Calmar table formulas below to know which column belongs to
# which month.
YEARCUM_FIRST_COL = 30  # AD
YEARPEAK_FIRST_COL = 38  # AL
YEARDD_FIRST_COL = 46  # AT

def pm(month: str) -> str:
    return f"'Position Model - {month}'"


def config_tier_cell(month: str, row: int) -> str:
    return f"Config!{CONFIG_MONTH_COLUMNS[month]}{row}"


def calc_range(col: str) -> str:
    """Absolute Dashboard_Calc range over every pre-filled row."""
    return f"{CALC_SHEET}!${col}$2:${col}${PADDED_LAST_ROW}"


def dynamic_range_name(col: str, sheet: str = CALC_SHEET) -> str:
    """Formula text for a range that starts at row 2 and stops at the last
    row in use (row 1 + Dashboard_Calc!Y1). INDEX, not OFFSET, so the name
    is non-volatile."""
    ref = f"'{sheet}'" if " " in sheet else sheet
    return f"{ref}!${col}$2:INDEX({ref}!${col}$2:${col}${PADDED_LAST_ROW},{ROWS_IN_USE})"


def _mean_and_variance(rng: str) -> tuple[str, str]:
    """Population mean/variance over the rows in use. Unused rows hold 0, so
    SUM is unaffected; the variance weights each squared deviation by the
    row's in-use flag so unused rows add nothing."""
    mean = f"SUM({rng})/{ROWS_IN_USE}"
    var = f"SUMPRODUCT({calc_range(VALID_COL)}*({rng}-({mean}))^2)/{ROWS_IN_USE}"
    return mean, var


# ---------------------------------------------------------------------------
# Year x Month per-metric formula builders. All continuous-return metrics
# (PnL/Sharpe/Sortino/Calmar/Max Drawdown) are sourced from Dashboard_Calc's
# already-ISNUMBER-cleaned per-month daily column (B:I) filtered by year via
# bare-comparison SUMPRODUCT (year_range=year)*range -- never IF/IFERROR
# wrapping a multi-cell range, per this environment's known implicit-
# intersection gotcha. The trade-boundary metrics (Win Rate/Avg Win/Avg
# Loss/Profit Factor/Trade Count) extend the corrected since-inception
# formulas (BS active-to-inactive transition) with one extra COUNTIFS/
# SUMIFS year criterion pulled from that month's own live CO column.
# ---------------------------------------------------------------------------
# `year` in these builders is a cell reference (the table row's own Year
# cell, e.g. "$A12"), so a table row follows whatever year it displays.
# Unused Dashboard_Calc rows have Year 0, so a year filter skips them.
def _dc_daily_range(month: str) -> str:
    return calc_range(get_column_letter(2 + MONTHS.index(month)))


def _dc_year_range() -> str:
    return calc_range("A")


def _dc_yeardd_range(month: str) -> str:
    return calc_range(get_column_letter(YEARDD_FIRST_COL + MONTHS.index(month)))


def _year_filtered_mean(month: str, year: int) -> str:
    yr, rng = _dc_year_range(), _dc_daily_range(month)
    return f"SUMPRODUCT(({yr}={year})*{rng})/SUMPRODUCT(({yr}={year})*1)"


def _year_month_pnl_formula(month: str, year: int) -> str:
    yr, rng = _dc_year_range(), _dc_daily_range(month)
    return f"=SUMPRODUCT(({yr}={year})*{rng})"


def _year_month_sharpe_formula(month: str, year: int) -> str:
    yr, rng = _dc_year_range(), _dc_daily_range(month)
    mean = _year_filtered_mean(month, year)
    var = f"SUMPRODUCT(({yr}={year})*({rng}-({mean}))^2)/SUMPRODUCT(({yr}={year})*1)"
    return f"=IFERROR(({mean})/SQRT({var})*SQRT(252),0)"


def _year_month_sortino_formula(month: str, year: int) -> str:
    yr, rng = _dc_year_range(), _dc_daily_range(month)
    mean = _year_filtered_mean(month, year)
    downside_var = f"SUMPRODUCT(({yr}={year})*({rng}<0)*{rng}^2)/SUMPRODUCT(({yr}={year})*1)"
    return f"=IFERROR(({mean})/SQRT({downside_var})*SQRT(252),0)"


def _year_month_maxdd_expr(month: str, year: int) -> str:
    """Bare (no leading '=') expression -- reused inside Calmar's formula."""
    yr, dd_rng = _dc_year_range(), _dc_yeardd_range(month)
    return f"IFERROR(INDEX({dd_rng},MATCH({year},{yr},0)+COUNTIF({yr},{year})-1),0)"


def _year_month_maxdd_formula(month: str, year: int) -> str:
    return f"={_year_month_maxdd_expr(month, year)}"


def _year_month_calmar_formula(month: str, year: int) -> str:
    mean = _year_filtered_mean(month, year)
    maxdd = _year_month_maxdd_expr(month, year)
    return f"=IFERROR(({mean})*252/ABS({maxdd}),0)"


def _year_month_trade_parts(month: str, year: int):
    bs_cur = f"{pm(month)}!$BS$2:$BS${PADDED_LAST_ROW - 1}"
    bs_next = f"{pm(month)}!$BS$3:$BS${PADDED_LAST_ROW}"
    trade_pnl = f"{pm(month)}!$BU$2:$BU${PADDED_LAST_ROW - 1}"
    yr_raw = f"{pm(month)}!$CO$2:$CO${PADDED_LAST_ROW - 1}"
    trade_count = f'COUNTIFS({bs_cur},1,{bs_next},"<>1",{yr_raw},{year})'
    wins = f'COUNTIFS({bs_cur},1,{bs_next},"<>1",{yr_raw},{year},{trade_pnl},">0")'
    gross_win = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{yr_raw},{year},{trade_pnl},">0")'
    gross_loss = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{yr_raw},{year},{trade_pnl},"<0")'
    return bs_cur, bs_next, trade_pnl, yr_raw, trade_count, wins, gross_win, gross_loss


def _year_month_winrate_formula(month: str, year: int) -> str:
    _, _, _, _, trade_count, wins, _, _ = _year_month_trade_parts(month, year)
    return f"=IFERROR(({wins})/({trade_count}),0)"


def _year_month_avgwin_formula(month: str, year: int) -> str:
    bs_cur, bs_next, trade_pnl, yr_raw, _, _, _, _ = _year_month_trade_parts(month, year)
    return f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{yr_raw},{year},{trade_pnl},">0"),0)'


def _year_month_avgloss_formula(month: str, year: int) -> str:
    bs_cur, bs_next, trade_pnl, yr_raw, _, _, _, _ = _year_month_trade_parts(month, year)
    return f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{yr_raw},{year},{trade_pnl},"<0"),0)'


def _year_month_profitfactor_formula(month: str, year: int) -> str:
    _, _, _, _, _, _, gross_win, gross_loss = _year_month_trade_parts(month, year)
    return f"=IFERROR(({gross_win})/ABS({gross_loss}),0)"


def _year_month_tradecount_formula(month: str, year: int) -> str:
    _, _, _, _, trade_count, _, _, _ = _year_month_trade_parts(month, year)
    return f"={trade_count}"



# ---------------------------------------------------------------------------
# Metric catalogue and since-inception builders.
# ---------------------------------------------------------------------------
CURRENCY_FMT = '$#,##0;[RED]($#,##0)'
RATIO_FMT = '0.00'
PCT_FMT = '0.0%'

METRICS = [
    ("pnl", "Total PnL", CURRENCY_FMT),
    ("sharpe", "Sharpe", RATIO_FMT),
    ("sortino", "Sortino", RATIO_FMT),
    ("calmar", "Calmar", RATIO_FMT),
    ("maxdd", "Max Drawdown", CURRENCY_FMT),
    ("winrate", "Win Rate", PCT_FMT),
    ("avgwin", "Avg $ Win", CURRENCY_FMT),
    ("avgloss", "Avg $ Loss", CURRENCY_FMT),
    ("pf", "Profit Factor", RATIO_FMT),
    ("trades", "Trade Count", "0"),
]
METRIC_KEYS = [k for k, _, _ in METRICS]
TRADE_METRICS = {"winrate", "avgwin", "avgloss", "pf", "trades"}
SERIES = ["Portfolio"] + MONTHS

_YEAR_BUILDERS = {
    "pnl": _year_month_pnl_formula,
    "sharpe": _year_month_sharpe_formula,
    "sortino": _year_month_sortino_formula,
    "calmar": _year_month_calmar_formula,
    "maxdd": _year_month_maxdd_formula,
    "winrate": _year_month_winrate_formula,
    "avgwin": _year_month_avgwin_formula,
    "avgloss": _year_month_avgloss_formula,
    "pf": _year_month_profitfactor_formula,
    "trades": _year_month_tradecount_formula,
}


def month_year_formula(metric: str, month: str, year_ref: str) -> str:
    """`month`'s `metric` for the calendar year held in cell `year_ref`."""
    return _YEAR_BUILDERS[metric](month, year_ref)


def month_inception_formula(metric: str, month: str) -> str:
    """`month`'s since-inception `metric` (risk-adjusted basis)."""
    bv = f"{pm(month)}!$BV$2:$BV${PADDED_LAST_ROW}"
    by = f"{pm(month)}!$BY$2:$BY${PADDED_LAST_ROW}"
    clean_rng = calc_range(get_column_letter(2 + MONTHS.index(month)))
    mean, var = _mean_and_variance(clean_rng)
    # BY is a running worst-so-far series, so its last value is the max drawdown.
    max_dd = f"LOOKUP(2,1/ISNUMBER({by}),{by})"
    # A trade ends on a BS 1 -> not-1 transition (see the year builders).
    bs_cur = f"{pm(month)}!$BS$2:$BS${PADDED_LAST_ROW - 1}"
    bs_next = f"{pm(month)}!$BS$3:$BS${PADDED_LAST_ROW}"
    trade_pnl = f"{pm(month)}!$BU$2:$BU${PADDED_LAST_ROW - 1}"
    trade_count = f'COUNTIFS({bs_cur},1,{bs_next},"<>1")'
    wins = f'COUNTIFS({bs_cur},1,{bs_next},"<>1",{trade_pnl},">0")'
    gross_win = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},">0")'
    gross_loss = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},"<0")'
    return {
        "pnl": f"=LOOKUP(2,1/ISNUMBER({bv}),{bv})",
        "sharpe": f"=IFERROR(({mean})/SQRT({var})*SQRT(252),0)",
        "sortino": f"=IFERROR(({mean})/SQRT(SUMPRODUCT(({clean_rng}<0)*{clean_rng}^2)/{ROWS_IN_USE})*SQRT(252),0)",
        "calmar": f"=IFERROR(({mean})*252/ABS({max_dd}),0)",
        "maxdd": f"={max_dd}",
        "winrate": f"=IFERROR(({wins})/({trade_count}),0)",
        "avgwin": f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},">0"),0)',
        "avgloss": f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},"<0"),0)',
        "pf": f"=IFERROR(({gross_win})/ABS({gross_loss}),0)",
        "trades": f"={trade_count}",
    }[metric]


def portfolio_inception_formula(metric: str) -> str:
    """Portfolio (sum of 8 months) since-inception `metric`. Trade metrics
    have no portfolio definition and return "n/a"."""
    daily, cum, maxdd_rng = calc_range("J"), calc_range("K"), calc_range("N")
    mean, var = _mean_and_variance(daily)
    max_dd = f"LOOKUP(2,1/ISNUMBER({maxdd_rng}),{maxdd_rng})"
    formulas = {
        "pnl": f"=LOOKUP(2,1/ISNUMBER({cum}),{cum})",
        "sharpe": f"=IFERROR(({mean})/SQRT({var})*SQRT(252),0)",
        "sortino": f"=IFERROR(({mean})/SQRT(SUMPRODUCT(({daily}<0)*{daily}^2)/{ROWS_IN_USE})*SQRT(252),0)",
        "calmar": f"=IFERROR(({mean})*252/ABS({max_dd}),0)",
        "maxdd": f"={max_dd}",
    }
    return formulas.get(metric, '="n/a"')


# ---------------------------------------------------------------------------
# Portfolio per calendar year. Same shapes as the per-month year builders,
# on the portfolio daily column J and a portfolio year-reset drawdown (BB:BC).
# ---------------------------------------------------------------------------
PORT_YEARPEAK_COL = "BB"
PORT_YEARDD_COL = "BC"


def _filtered_stats(rng: str, year: str) -> tuple[str, str, str]:
    yr = _dc_year_range()
    days = f"SUMPRODUCT(({yr}={year})*1)"
    mean = f"SUMPRODUCT(({yr}={year})*{rng})/{days}"
    var = f"SUMPRODUCT(({yr}={year})*({rng}-({mean}))^2)/{days}"
    down = f"SUMPRODUCT(({yr}={year})*({rng}<0)*{rng}^2)/{days}"
    return mean, var, down


def portfolio_year_formula(metric: str, year_ref: str) -> str:
    """Portfolio `metric` for the calendar year in `year_ref`. Trade metrics
    have no portfolio definition ("n/a"); the matrix sums Trade Count."""
    yr, rng = _dc_year_range(), calc_range("J")
    mean, var, down = _filtered_stats(rng, year_ref)
    dd = calc_range(PORT_YEARDD_COL)
    maxdd = f"IFERROR(INDEX({dd},MATCH({year_ref},{yr},0)+COUNTIF({yr},{year_ref})-1),0)"
    return {
        "pnl": f"=SUMPRODUCT(({yr}={year_ref})*{rng})",
        "sharpe": f"=IFERROR(({mean})/SQRT({var})*SQRT(252),0)",
        "sortino": f"=IFERROR(({mean})/SQRT({down})*SQRT(252),0)",
        "calmar": f"=IFERROR(({mean})*252/ABS({maxdd}),0)",
        "maxdd": f"={maxdd}",
    }.get(metric, '="n/a"')


# ---------------------------------------------------------------------------
# Metric matrix: every metric x series x period in one hidden block, so the
# Dashboard's tiles, year table, detail table and bar chart are all INDEX
# lookups into figures built by the formulas above.
#   metric k (0-based): header row k*11+1, series s (Portfolio=1) row k*11+1+s
#   periods: FB = All (since inception), FC.. = TABLE_YEARS (row-1 headers)
# ---------------------------------------------------------------------------
MATRIX_LABEL_COL = 157  # FA
MATRIX_FIRST_PERIOD_COL = 158  # FB
MATRIX_BLOCK_ROWS = 11
PERIODS = ["All"] + TABLE_YEARS
MATRIX_LAST_COL = MATRIX_FIRST_PERIOD_COL + len(PERIODS) - 1  # FN
MATRIX_LAST_ROW = MATRIX_BLOCK_ROWS * len(METRICS)  # 110
MATRIX_RANGE = (f"{CALC_SHEET}!${get_column_letter(MATRIX_FIRST_PERIOD_COL)}$1:"
                f"${get_column_letter(MATRIX_LAST_COL)}${MATRIX_LAST_ROW}")

# Dashboard control cells (placed by the Performance Review builder).
MONTH_CELL = "B18"
YEAR_CELL = "D18"
METRIC_CELL = "F18"
ALL_YEARS = "All years"
YEAR_OPTIONS_COL = "EZ"


def _abs(cell: str) -> str:
    col = "".join(ch for ch in cell if ch.isalpha())
    return f"${col}${cell[len(col):]}"


def matrix_address(metric: str, series: str, period) -> str:
    row = METRIC_KEYS.index(metric) * MATRIX_BLOCK_ROWS + 1 + SERIES.index(series) + 1
    return f"{get_column_letter(MATRIX_FIRST_PERIOD_COL + PERIODS.index(period))}{row}"


def matrix_row(metric_index_expr: str, series_index_expr: str) -> str:
    return f"(({metric_index_expr})-1)*{MATRIX_BLOCK_ROWS}+1+({series_index_expr})"


def matrix_lookup(metric_index_expr: str, series_index_expr: str, period_index_expr: str) -> str:
    return f"INDEX({MATRIX_RANGE},{matrix_row(metric_index_expr, series_index_expr)},{period_index_expr})"


def build_metric_matrix(ws) -> None:
    ws.cell(row=1, column=MATRIX_LABEL_COL, value="Metric Matrix")
    for p, period in enumerate(PERIODS):
        ws.cell(row=1, column=MATRIX_FIRST_PERIOD_COL + p, value=period)
    for k, (key, label, _fmt) in enumerate(METRICS):
        header = k * MATRIX_BLOCK_ROWS + 1
        if k:
            ws.cell(row=header, column=MATRIX_LABEL_COL, value=label)
        for s, series in enumerate(SERIES, start=1):
            r = header + s
            ws.cell(row=r, column=MATRIX_LABEL_COL, value=f"{label} | {series}")
            for p, period in enumerate(PERIODS):
                col = MATRIX_FIRST_PERIOD_COL + p
                letter = get_column_letter(col)
                year_ref = f"{letter}$1"
                if series == "Portfolio" and key == "trades":
                    formula = f"=SUM({letter}{r + 1}:{letter}{r + len(MONTHS)})"
                elif series == "Portfolio":
                    formula = (portfolio_inception_formula(key) if period == "All"
                               else portfolio_year_formula(key, year_ref))
                else:
                    formula = (month_inception_formula(key, series) if period == "All"
                               else month_year_formula(key, series, year_ref))
                if period != "All" and formula != '="n/a"':
                    formula = f'=IF({year_ref}>{LATEST_YEAR},"",{formula.removeprefix("=")})'
                ws.cell(row=r, column=col, value=formula)


def build_controls(ws) -> None:
    """Calc control cells the Dashboard dropdowns drive (Z1:Z6) and the Year
    dropdown's option list (EZ2:EZ14)."""
    series_list = "{" + ",".join(f'"{s}"' for s in SERIES) + "}"
    metric_list = "{" + ",".join(f'"{label}"' for _, label, _ in METRICS) + "}"
    month, year, metric = (f"Dashboard!{_abs(c)}" for c in (MONTH_CELL, YEAR_CELL, METRIC_CELL))
    years = f"$A$2:$A${PADDED_LAST_ROW}"
    ws["Z1"] = f"=MATCH({month},{series_list},0)"
    ws["Z2"] = f'=IF({year}="{ALL_YEARS}",1,{year}-{FIRST_YEAR}+2)'
    ws["Z3"] = f"=IF($Z$2=1,0,{year})"
    ws["Z4"] = f"=IF($Z$3=0,1,MATCH($Z$3,{years},0))"
    ws["Z5"] = f"=IF($Z$3=0,$Y$1,$Z$4+COUNTIF({years},$Z$3)-1)"
    ws["Z6"] = f"=MATCH({metric},{metric_list},0)"
    ws[f"{YEAR_OPTIONS_COL}1"] = "Year Options"
    ws[f"{YEAR_OPTIONS_COL}2"] = ALL_YEARS
    for j, y in enumerate(TABLE_YEARS):
        ws[f"{YEAR_OPTIONS_COL}{3 + j}"] = f'=IF({y}<=$Y$2,{y},"")'
