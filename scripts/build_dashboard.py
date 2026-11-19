"""Builds the 'Dashboard' tab (+ a hidden 'Dashboard_Calc' helper sheet) into
the Knox multi-month BOHO/D4 workbook, summarizing current position/signal
and since-inception performance across all 8 contract-month Position Model
tabs.

All dashboard formulas are plain (no dynamic-array/spill formulas, no
implicit-array conditions inside MATCH/LOOKUP), per the lessons in
feedback-johansen-gotchas / feedback-sboxho-gotchas: openpyxl can't safely
write spill metadata, and this Excel/COM environment has been seen to choke
on two-range implicit-array patterns even when Excel 365 dynamic arrays are
otherwise in use. Every "find the current/last value" lookup here uses the
single-condition LOOKUP(2,1/ISNUMBER(range),range) idiom already proven safe
in this exact workbook (Johansen_Monthly's Active End Row pointer), and every
"value on the last day of year Y" lookup uses MATCH(exact)+COUNTIF+INDEX
(three independent single-condition calls), never a two-condition product
inside 1/(...).

Basis convention: every headline performance metric (Sharpe, Sortino,
Calmar, Max Drawdown, Win Rate, Avg $ Win/Loss, Profit Factor, Trade Count,
the cumulative-PnL and PnL-by-year charts) uses the risk-adjusted
(stop-loss-aware) columns -- BT "Effective Daily PnL", BU "Effective Trade
Cumulative PnL", BV "Risk Adjusted Cumulative $ PnL", BY "RA_Max Drawdown" --
never the gross per-leg columns (AV:BJ), because the RA basis is what
actually would have happened once the stop-loss overlay is applied. Gross
per-leg PnL is shown only as a secondary attribution sub-table.

Usage (the workbook is named on the command line, so renaming it never
needs a code change; a timestamped backup is written next to it first):
    python scripts/build_dashboard.py "<Knox workbook>.xlsx"
"""

import argparse
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

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

# Largest row common to ALL 8 months' currently-real data -- used only for
# the portfolio-level (cross-month, row-aligned) staging block, where every
# month must contribute the same date range. Computed dynamically at the
# start of main() via compute_common_real_last_row(), never hand-set: a
# hardcoded value silently goes stale and grows more wrong every day new
# trading rows get appended without a rebuild -- exactly what happened to
# the 2026-08-14 build (hardcoded at 1392), found ~3 months/67 rows stale on
# 2026-11-16. None here is a sentinel so any code path that runs before
# main() populates it fails loudly instead of silently reusing a stale
# number.
COMMON_REAL_LAST_ROW = None

YEARS = [2021, 2022, 2023, 2024, 2025, 2026]
MAX_DAY_INDEX = 260  # comfortably above any year's ~252 trading days

# Dashboard_Calc column layout for the per-month, year-reset cumulative/peak/
# drawdown helper columns (see build_calc_sheet) -- used by the Year x Month
# Max Drawdown/Calmar table formulas below to know which column belongs to
# which month.
YEARCUM_FIRST_COL = 30  # AD
YEARPEAK_FIRST_COL = 38  # AL
YEARDD_FIRST_COL = 46  # AT

# Okabe-Ito colorblind-safe categorical set (7 hues) plus one extra
# well-separated hue for the 8th month; black is reserved for the Portfolio
# aggregate line so it reads as "the reference line," not just another month.
MONTH_COLORS = {
    "January": "E69F00",
    "March": "56B4E9",
    "May": "009E73",
    "July": "F0E442",
    "August": "0072B2",
    "September": "D55E00",
    "October": "CC79A7",
    "December": "6F4E37",
}
PORTFOLIO_COLOR = "000000"

# Dropdown options for the "Filter Charts By" selector, in a fixed order that
# every CHOOSE(...)-based helper formula below matches positionally: index 1
# = Portfolio, indices 2-9 = the 8 months in the same order as MONTHS.
SELECTOR_OPTIONS = ["Portfolio"] + MONTHS

# One diverging pair (blue=positive / vermillion=negative / white midpoint),
# reused for every polarity encoding on this tab (Z-score color scale,
# position data bars, PnL heatmap) for a consistent mental model.
DIVERGING_POS = "0072B2"
DIVERGING_NEG = "D55E00"
DIVERGING_MID = "FFFFFF"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FONT = Font(bold=True, size=12)
BOLD = Font(bold=True)

CURRENCY_FMT = '$#,##0;[RED]($#,##0)'
RATIO_FMT = '0.00'
PCT_FMT = '0.0%'


def pm(month: str) -> str:
    return f"'Position Model - {month}'"


def config_tier_cell(month: str, row: int) -> str:
    return f"Config!{CONFIG_MONTH_COLUMNS[month]}{row}"


def _find_last_real_row(data_wb, month: str) -> int:
    """Last row in <month>'s Position Model tab whose column A holds a real
    trading date, using the same >DATE(2015,1,1) convention the Dashboard's
    own Cockpit 'As Of' formula uses to skip unused template padding (which
    reads back as a literal 0 -- Excel's 1899-12-30 epoch -- not a blank)."""
    ws = data_wb[f"Position Model - {month}"]
    last_row = 1
    for row in ws.iter_rows(min_row=2, max_row=PADDED_LAST_ROW, min_col=1, max_col=1):
        cell = row[0]
        v = cell.value
        if isinstance(v, datetime) and v.year >= 2015:
            last_row = cell.row
    return last_row


def compute_common_real_last_row(target_file: Path) -> int:
    """Dynamically determines COMMON_REAL_LAST_ROW from the live file's own
    cached values, instead of a hand-set constant that silently goes stale
    as more trading days get appended (see the module-level comment on
    COMMON_REAL_LAST_ROW). Requires the file to carry cached formula results
    for column A of each Position Model tab -- true under normal day-to-day
    use, since Excel recalculates and caches on every open/edit/save."""
    data_wb = openpyxl.load_workbook(target_file, data_only=True, read_only=True)
    try:
        last_rows = {month: _find_last_real_row(data_wb, month) for month in MONTHS}
    finally:
        data_wb.close()
    common = min(last_rows.values())
    if len(set(last_rows.values())) > 1:
        print(
            f"Per-month last real row differs: {last_rows} -- using the "
            f"minimum ({common}) so every month contributes the same date "
            "range to the portfolio rollup."
        )
    else:
        print(f"All 8 months' last real row agrees: {common}.")
    return common


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
def _dc_daily_range(month: str) -> str:
    col = get_column_letter(2 + MONTHS.index(month))
    return f"Dashboard_Calc!${col}$2:${col}${COMMON_REAL_LAST_ROW}"


def _dc_year_range() -> str:
    return f"Dashboard_Calc!$A$2:$A${COMMON_REAL_LAST_ROW}"


def _dc_yeardd_range(month: str) -> str:
    col = get_column_letter(YEARDD_FIRST_COL + MONTHS.index(month))
    return f"Dashboard_Calc!${col}$2:${col}${COMMON_REAL_LAST_ROW}"


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


def build_year_month_table(
    ws, start_row: int, title: str, cell_formula_fn, number_format: str,
    color_scale: str | None = "diverging2", total_col: bool = False,
) -> int:
    """Generic Year (rows) x Month (cols) table, matching the layout of the
    original Calendar-Year PnL heatmap. cell_formula_fn(month, year) returns
    a full formula string (leading '='). color_scale: 'diverging2' (min/max),
    'diverging3' (min/0/max), or None (no conditional formatting -- used for
    Trade Count, where color doesn't imply good/bad). Returns the last row
    used, so callers can stack the next section below it."""
    header_row = start_row
    section_title(ws, header_row, title)
    col_row = header_row + 1
    ws.cell(row=col_row, column=1, value="Year")
    for i, month in enumerate(MONTHS):
        ws.cell(row=col_row, column=2 + i, value=month)
    last_col = 1 + len(MONTHS)
    if total_col:
        last_col += 1
        ws.cell(row=col_row, column=last_col, value="Total")
    style_header_row(ws, col_row, 1, last_col)

    first_row = col_row + 1
    for j, year in enumerate(YEARS):
        r = first_row + j
        ws.cell(row=r, column=1, value=year)
        for i, month in enumerate(MONTHS):
            col = 2 + i
            ws.cell(row=r, column=col, value=cell_formula_fn(month, year))
            ws.cell(row=r, column=col).number_format = number_format
        if total_col:
            first_month_col = get_column_letter(2)
            last_month_col = get_column_letter(1 + len(MONTHS))
            tc = 2 + len(MONTHS)
            ws.cell(row=r, column=tc, value=f"=SUM({first_month_col}{r}:{last_month_col}{r})")
            ws.cell(row=r, column=tc).number_format = number_format
            ws.cell(row=r, column=tc).font = BOLD

    last_row = first_row + len(YEARS) - 1
    if color_scale == "diverging3":
        ws.conditional_formatting.add(
            f"B{first_row}:{get_column_letter(1 + len(MONTHS))}{last_row}",
            ColorScaleRule(
                start_type="min", start_color=DIVERGING_NEG,
                mid_type="num", mid_value=0, mid_color=DIVERGING_MID,
                end_type="max", end_color=DIVERGING_POS,
            ),
        )
    elif color_scale == "diverging2":
        ws.conditional_formatting.add(
            f"B{first_row}:{get_column_letter(1 + len(MONTHS))}{last_row}",
            ColorScaleRule(
                start_type="min", start_color=DIVERGING_NEG,
                end_type="max", end_color=DIVERGING_POS,
            ),
        )
    return last_row


def style_header_row(ws, row: int, first_col: int, last_col: int) -> None:
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def section_title(ws, row: int, text: str) -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT


# ---------------------------------------------------------------------------
# Dashboard_Calc: hidden staging sheet for the one thing no single Position
# Model tab has -- the portfolio-level (sum of all 8 months) daily RA PnL
# series, its cumulative/drawdown tracking, and the per-year day-indexed
# slices used by the "PnL by year" overlay chart.
# ---------------------------------------------------------------------------
def build_calc_sheet(wb):
    if "Dashboard_Calc" in wb.sheetnames:
        del wb["Dashboard_Calc"]
    ws = wb.create_sheet("Dashboard_Calc")
    ws.sheet_state = "hidden"

    ws["A1"] = "Year"
    for i, month in enumerate(MONTHS):
        ws.cell(row=1, column=2 + i, value=f"{month} Effective Daily PnL")
    ws["J1"] = "Portfolio Daily PnL"
    ws["K1"] = "Portfolio Cumulative PnL"
    ws["L1"] = "Portfolio Running Peak"
    ws["M1"] = "Portfolio Daily Drawdown"
    ws["N1"] = "Portfolio Max Drawdown"
    ws["O1"] = "Portfolio Calendar-Year Cumulative PnL"

    for r in range(2, COMMON_REAL_LAST_ROW + 1):
        ws.cell(row=r, column=1, value=f"={pm('January')}!CO{r}")
        for i, month in enumerate(MONTHS):
            # ISNUMBER-guarded (not just IFERROR): BT genuinely errors during
            # each month's early warm-up rows (expected/pre-existing, per
            # this project's history), AND at least one source cell
            # ('Position Model - March'!BT17) was found to hold stray literal
            # text ("1.23.26 formula") instead of a formula/number -- IFERROR
            # alone doesn't catch that, since text isn't an error. Without
            # this guard, either case would propagate forever through the
            # running-cumulative chain below (K:O), corrupting every
            # subsequent day for the rest of the sheet. This is a single-cell
            # reference, not a multi-cell range, so it doesn't hit the
            # IF/IFERROR-on-a-range implicit-intersection issue documented
            # in the performance-table formulas below.
            src = f"{pm(month)}!BT{r}"
            ws.cell(row=r, column=2 + i, value=f"=IF(ISNUMBER({src}),{src},0)")
        daily_range = f"B{r}:I{r}"
        ws.cell(row=r, column=10, value=f"=SUM({daily_range})")
        if r == 2:
            ws.cell(row=r, column=11, value="=J2")
            ws.cell(row=r, column=12, value="=K2")
            ws.cell(row=r, column=13, value="=K2-L2")
            ws.cell(row=r, column=14, value="=M2")
            ws.cell(row=r, column=15, value="=J2")
        else:
            ws.cell(row=r, column=11, value=f"=K{r - 1}+J{r}")
            ws.cell(row=r, column=12, value=f"=MAX(L{r - 1},K{r})")
            ws.cell(row=r, column=13, value=f"=K{r}-L{r}")
            ws.cell(row=r, column=14, value=f"=MIN(N{r - 1},M{r})")
            ws.cell(row=r, column=15, value=f"=IF(A{r}<>A{r - 1},J{r},O{r - 1}+J{r})")

    # "Filter Charts By" selector support: Dashboard!O2 holds a Data
    # Validation dropdown (Portfolio or one of the 8 months). Z1 turns that
    # text into a 1-9 index via SELECTOR_OPTIONS' fixed order; AA/AB/AC then
    # mirror the existing J/K/O portfolio-level daily/cumulative/calendar-
    # year-reset-cumulative pattern, but for whichever single series is
    # currently selected -- reusing B:I's already-ISNUMBER-cleaned per-month
    # daily columns and J's portfolio daily column as CHOOSE's 9 options
    # (all single-cell references per row, so this is exactly as safe as the
    # other single-cell CHOOSE/IF idioms already used in this project; no
    # multi-cell range is ever wrapped in IF/IFERROR/CHOOSE here).
    selector_list_literal = '{' + ','.join(f'"{name}"' for name in SELECTOR_OPTIONS) + '}'
    ws["Z1"] = f"=MATCH('Dashboard'!$O$2,{selector_list_literal},0)"
    ws["AA1"] = "Selected Series Daily PnL"
    ws["AB1"] = "Selected Series Cumulative PnL"
    ws["AC1"] = "Selected Series Calendar-Year Cumulative PnL"
    daily_choices = ",".join(["J{r}"] + [get_column_letter(2 + i) + "{r}" for i in range(len(MONTHS))])
    for r in range(2, COMMON_REAL_LAST_ROW + 1):
        ws.cell(row=r, column=27, value=f"=CHOOSE($Z$1,{daily_choices.format(r=r)})")
        if r == 2:
            ws.cell(row=r, column=28, value="=AA2")
            ws.cell(row=r, column=29, value="=AA2")
        else:
            ws.cell(row=r, column=28, value=f"=AB{r - 1}+AA{r}")
            ws.cell(row=r, column=29, value=f"=IF(A{r}<>A{r - 1},AA{r},AC{r - 1}+AA{r})")

    # Per-year, day-indexed slices for the "PnL by year" overlay chart: index
    # 0 is a literal zero (the common starting point every year's line
    # anchors to); index k (1-based) pulls the k-th trading day of that
    # year's calendar-year-reset cumulative PnL for the currently selected
    # series (column AC above -- selector-driven, defaults to Portfolio).
    # Contiguity of same-year rows (the source tabs are date-sorted) makes a
    # single-condition MATCH (first occurrence of the year) + COUNTIF (how
    # many rows have that year) sufficient to locate any day k within that
    # year -- no two-condition array lookup needed.
    ws["Q1"] = "Day Index"
    year_range = f"$A$2:$A${COMMON_REAL_LAST_ROW}"
    cum_range = f"$AC$2:$AC${COMMON_REAL_LAST_ROW}"
    for j, year in enumerate(YEARS):
        col = 18 + j  # R=18 .. W=23
        ws.cell(row=1, column=col, value=year)
    for k in range(0, MAX_DAY_INDEX + 1):
        r = 2 + k
        ws.cell(row=r, column=17, value=k)
        for j, year in enumerate(YEARS):
            col = 18 + j
            col_letter = get_column_letter(col)
            year_cell = f"{col_letter}$1"
            formula = (
                f"=IF($Q{r}=0,0,"
                f"IF($Q{r}>COUNTIF({year_range},{year_cell}),\"\","
                f"INDEX({cum_range},MATCH({year_cell},{year_range},0)+$Q{r}-1)))"
            )
            ws.cell(row=r, column=col, value=formula)

    # Per-month, year-reset cumulative/peak/drawdown (AD:BA, 8 months x 3
    # columns) -- needed for the Max Drawdown/Calmar-by-Year-x-Month tables.
    # Mirrors columns AA:AC's selector-driven pattern (and, before that,
    # J:O's portfolio pattern), but computed independently for EVERY month
    # at once rather than just the currently-selected one, since the
    # Year-x-Month tables show all 8 months simultaneously. Deliberately
    # NOT reusing the live Position Model tabs' own 'CD' (RA Calendar Year
    # Cumulative PnL) column for this: CD has a real pre-existing bug where
    # its year-reset branch discards that year's first day of PnL entirely
    # (resets to a literal 0 instead of that day's own BT value), found
    # while investigating why the PnL-by-year-x-month heatmap didn't tie to
    # the Performance table's since-inception total. Rebuilding this
    # calculation here, from B:I's already-correct cleaned daily PnL, avoids
    # that bug rather than propagating it, and keeps the fix scoped to the
    # dashboard rather than editing the live per-month trading formulas.
    for i, month in enumerate(MONTHS):
        daily_col = get_column_letter(2 + i)
        yc_col = get_column_letter(YEARCUM_FIRST_COL + i)
        yp_col = get_column_letter(YEARPEAK_FIRST_COL + i)
        yd_col = get_column_letter(YEARDD_FIRST_COL + i)
        ws.cell(row=1, column=YEARCUM_FIRST_COL + i, value=f"{month} Year-Reset Cumulative PnL")
        ws.cell(row=1, column=YEARPEAK_FIRST_COL + i, value=f"{month} Year-Reset Running Peak")
        ws.cell(row=1, column=YEARDD_FIRST_COL + i, value=f"{month} Year-Reset Max Drawdown (running)")
        for r in range(2, COMMON_REAL_LAST_ROW + 1):
            if r == 2:
                ws.cell(row=r, column=YEARCUM_FIRST_COL + i, value=f"={daily_col}2")
                ws.cell(row=r, column=YEARPEAK_FIRST_COL + i, value=f"={yc_col}2")
                ws.cell(row=r, column=YEARDD_FIRST_COL + i, value=f"=MIN(0,{yc_col}2-{yp_col}2)")
            else:
                ws.cell(
                    row=r, column=YEARCUM_FIRST_COL + i,
                    value=f"=IF(A{r}<>A{r - 1},{daily_col}{r},{yc_col}{r - 1}+{daily_col}{r})",
                )
                ws.cell(
                    row=r, column=YEARPEAK_FIRST_COL + i,
                    value=f"=IF(A{r}<>A{r - 1},{yc_col}{r},MAX({yp_col}{r - 1},{yc_col}{r}))",
                )
                ws.cell(
                    row=r, column=YEARDD_FIRST_COL + i,
                    value=f"=IF(A{r}<>A{r - 1},MIN(0,{yc_col}{r}-{yp_col}{r}),MIN({yd_col}{r - 1},{yc_col}{r}-{yp_col}{r}))",
                )

    return ws


# ---------------------------------------------------------------------------
# Dashboard: the human-facing tab.
# ---------------------------------------------------------------------------
def build_dashboard_sheet(wb):
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)

    ws["A1"] = "BOHO / D4 Multi-Month Strategy — Dashboard"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = "Data through:"
    ws["A2"].font = BOLD
    last_date_terms = ",".join(
        f'LOOKUP(2,1/({pm(m)}!$A$2:$A${PADDED_LAST_ROW}>DATE(2015,1,1)),{pm(m)}!$A$2:$A${PADDED_LAST_ROW})'
        for m in MONTHS
    )
    ws["B2"] = f"=MAX({last_date_terms})"
    ws["B2"].number_format = "MM/DD/YYYY"

    # "Filter Charts By" selector: drives the two charts marked "Selected
    # Series" below (the cumulative-PnL drill-down and the PnL-by-year
    # overlay) via Dashboard_Calc's CHOOSE-based helper columns. Placed
    # directly above the chart column so its effect is visually obvious;
    # defaults to Portfolio so both charts' out-of-the-box view is unchanged
    # from before this control existed.
    ws["N2"] = "Filter Charts By:"
    ws["N2"].font = BOLD
    ws["O2"] = "Portfolio"
    selector_dv = DataValidation(
        type="list",
        formula1='"' + ",".join(SELECTOR_OPTIONS) + '"',
        allow_blank=False,
    )
    ws.add_data_validation(selector_dv)
    selector_dv.add(ws["O2"])

    # ---------------- Cockpit ----------------
    cockpit_header_row = 4
    section_title(ws, cockpit_header_row, "Current Snapshot")
    header_row = cockpit_header_row + 1
    headers = ["Month", "As Of", "Total Position", "Z-Score", "Dist. to Threshold", "Regime"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, 1, len(headers))

    first_data_row = header_row + 1
    for i, month in enumerate(MONTHS):
        r = first_data_row + i
        rng = f"{pm(month)}!$A$2:$A${PADDED_LAST_ROW}"
        ws.cell(row=r, column=1, value=month)
        # Plain ISNUMBER isn't enough here: several months' unused template
        # padding beyond real data evaluates to literal 0 (not blank), which
        # passes ISNUMBER and would be picked up as the "last" date. A real
        # trading date is always well past this literal epoch-zero, so
        # requiring range>DATE(2015,1,1) excludes the 0-padding while still
        # correctly skipping genuine error cells (LOOKUP's own error-tolerant
        # scan, the same mechanism the ISNUMBER-based idiom relies on).
        ws.cell(row=r, column=2, value=f'=LOOKUP(2,1/({rng}>DATE(2015,1,1)),{rng})')
        ws.cell(row=r, column=2).number_format = "MM/DD/YYYY"

        v_rng = f"{pm(month)}!$V$2:$V${PADDED_LAST_ROW}"
        ws.cell(row=r, column=3, value=f"=LOOKUP(2,1/ISNUMBER({v_rng}),{v_rng})")

        l_rng = f"{pm(month)}!$L$2:$L${PADDED_LAST_ROW}"
        z_formula = f"LOOKUP(2,1/ISNUMBER({l_rng}),{l_rng})"
        ws.cell(row=r, column=4, value=f"={z_formula}")
        ws.cell(row=r, column=4).number_format = RATIO_FMT

        t1 = config_tier_cell(month, 14)
        t2 = config_tier_cell(month, 15)
        t3 = config_tier_cell(month, 16)
        z_cell = f"$D{r}"
        ws.cell(
            row=r,
            column=5,
            value=(
                f"=MIN(ABS({z_cell}-{t1}),ABS({z_cell}-{t2}),ABS({z_cell}-{t3}),"
                f"ABS({z_cell}+{t1}),ABS({z_cell}+{t2}),ABS({z_cell}+{t3}))"
            ),
        )
        ws.cell(row=r, column=5).number_format = RATIO_FMT

        dc_rng = f"{pm(month)}!$DC$2:$DC${PADDED_LAST_ROW}"
        ws.cell(row=r, column=6, value=f'=LOOKUP(2,1/({dc_rng}<>""),{dc_rng})')

    last_cockpit_row = first_data_row + len(MONTHS) - 1

    ws.conditional_formatting.add(
        f"D{first_data_row}:D{last_cockpit_row}",
        ColorScaleRule(
            start_type="min", start_color=DIVERGING_NEG,
            mid_type="num", mid_value=0, mid_color=DIVERGING_MID,
            end_type="max", end_color=DIVERGING_POS,
        ),
    )
    ws.conditional_formatting.add(
        f"C{first_data_row}:C{last_cockpit_row}",
        DataBarRule(start_type="min", end_type="max", color=DIVERGING_POS),
    )
    # Dist. to Threshold: small = close to a regime shift (danger), large =
    # comfortably inside the current regime (safe). Reuses the same
    # diverging pair as Z-score/PnL rather than a fresh red/green scale, for
    # one consistent color language across the tab and to avoid a
    # red/green combination that isn't colorblind-safe.
    ws.conditional_formatting.add(
        f"E{first_data_row}:E{last_cockpit_row}",
        ColorScaleRule(
            start_type="min", start_color=DIVERGING_NEG,
            end_type="max", end_color=DIVERGING_POS,
        ),
    )
    # Regime: categorical, so a color scale doesn't apply -- flag "Extreme"
    # solid vermillion (matches the "danger" end of every other polarity
    # encoding on this tab) and "High" a lighter tint of the same hue;
    # "Normal" is left unstyled since it's the expected/calm state.
    regime_range = f"F{first_data_row}:F{last_cockpit_row}"
    ws.conditional_formatting.add(
        regime_range,
        CellIsRule(operator="equal", formula=['"Extreme"'],
                   fill=PatternFill("solid", fgColor=DIVERGING_NEG),
                   font=Font(color="FFFFFF", bold=True)),
    )
    ws.conditional_formatting.add(
        regime_range,
        CellIsRule(operator="equal", formula=['"High"'],
                   fill=PatternFill("solid", fgColor="F4C7A8")),
    )

    # ---------------- Performance summary ----------------
    perf_header_row = last_cockpit_row + 2
    section_title(ws, perf_header_row, "Performance Since Inception (Risk-Adjusted Basis)")
    perf_col_row = perf_header_row + 1
    perf_headers = [
        "Month", "RA Total PnL", "Sharpe", "Sortino", "Calmar", "Max Drawdown",
        "Win Rate", "Avg $ Win", "Avg $ Loss", "Profit Factor", "Trade Count",
    ]
    for i, h in enumerate(perf_headers, start=1):
        ws.cell(row=perf_col_row, column=i, value=h)
    style_header_row(ws, perf_col_row, 1, len(perf_headers))

    perf_first_row = perf_col_row + 1
    for i, month in enumerate(MONTHS):
        r = perf_first_row + i
        bv = f"{pm(month)}!$BV$2:$BV${PADDED_LAST_ROW}"
        by = f"{pm(month)}!$BY$2:$BY${PADDED_LAST_ROW}"
        # Sharpe/Sortino/Calmar are sourced from Dashboard_Calc's per-month
        # column (already IFERROR-cleaned cell-by-cell, confirmed 0 errors),
        # NOT the raw source column, and use only AVERAGE/STDEV.P/COUNT or
        # bare-range comparison operators. IFERROR(multi-cell-range,0) used
        # as a sub-expression inside SUMPRODUCT was tested in isolation and
        # confirmed to perform row-based implicit intersection against the
        # *formula's own row* in this Excel/COM environment rather than
        # evaluating element-wise -- i.e. it silently returns one wrong
        # scalar (whatever that source row happens to hold) broadcast across
        # the whole array, not a per-element cleanup. That bug, not a math
        # error, is what previously made every month's Sharpe/Sortino/Calmar
        # come out as exactly 0.
        clean_col = get_column_letter(2 + i)
        clean_rng = f"Dashboard_Calc!${clean_col}$2:${clean_col}${COMMON_REAL_LAST_ROW}"

        ws.cell(row=r, column=1, value=month)
        ws.cell(row=r, column=2, value=f"=LOOKUP(2,1/ISNUMBER({bv}),{bv})")
        ws.cell(row=r, column=2).number_format = CURRENCY_FMT

        mean_cell = f"AVERAGE({clean_rng})"
        sharpe = f"=IFERROR(({mean_cell})/STDEV.P({clean_rng})*SQRT(252),0)"
        ws.cell(row=r, column=3, value=sharpe)
        ws.cell(row=r, column=3).number_format = RATIO_FMT

        downside_var = f"SUMPRODUCT(({clean_rng}<0)*{clean_rng}^2)/COUNT({clean_rng})"
        sortino = f"=IFERROR(({mean_cell})/SQRT({downside_var})*SQRT(252),0)"
        ws.cell(row=r, column=4, value=sortino)
        ws.cell(row=r, column=4).number_format = RATIO_FMT

        # BY ("RA_Max Drawdown") is a running worst-so-far series (each row
        # = MIN(prior, today's drawdown)), so it is non-increasing by
        # construction -- its last non-error value already equals the true
        # overall max drawdown. Using the same LOOKUP-last-valid idiom as
        # everywhere else avoids AGGREGATE, which returns #NAME? (genuinely
        # unrecognized) in this Excel/COM environment despite being a
        # long-standing function elsewhere.
        max_dd = f"LOOKUP(2,1/ISNUMBER({by}),{by})"
        calmar = f"=IFERROR(({mean_cell})*252/ABS({max_dd}),0)"
        ws.cell(row=r, column=5, value=calmar)
        ws.cell(row=r, column=5).number_format = RATIO_FMT

        ws.cell(row=r, column=6, value=f"={max_dd}")
        ws.cell(row=r, column=6).number_format = CURRENCY_FMT

        # Trade boundary = a BS ("Effective Trade Active") transition from 1
        # to not-1, i.e. the last day a trade is actually active -- NOT "the
        # day before BN ('New Trade Start Flag') next fires", which is what
        # this used to key off. That original pairing silently assumed
        # trades are always back-to-back with no flat/no-position gap
        # between one closing and the next starting; whenever a gap exists
        # (confirmed via ground-truth recomputation from raw BN/BU data to
        # be the large majority of the time in this data), it picked up a
        # flat day (BU=0) instead of the prior trade's real realized PnL --
        # producing a systematic near-zero-loss bias in every month, most
        # visible in January/December where it zeroed out losses entirely.
        # BS/BU stay row-aligned the same way BN/BU used to (both start at
        # row 2), so this is a drop-in swap of which columns define the
        # boundary, not a new formula shape.
        bs_cur = f"{pm(month)}!$BS$2:$BS${PADDED_LAST_ROW - 1}"
        bs_next = f"{pm(month)}!$BS$3:$BS${PADDED_LAST_ROW}"
        trade_pnl = f"{pm(month)}!$BU$2:$BU${PADDED_LAST_ROW - 1}"
        trade_count = f'COUNTIFS({bs_cur},1,{bs_next},"<>1")'
        wins = f'COUNTIFS({bs_cur},1,{bs_next},"<>1",{trade_pnl},">0")'
        gross_win = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},">0")'
        gross_loss = f'SUMIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},"<0")'

        ws.cell(row=r, column=7, value=f"=IFERROR(({wins})/({trade_count}),0)")
        ws.cell(row=r, column=7).number_format = PCT_FMT
        ws.cell(
            row=r, column=8,
            value=f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},">0"),0)',
        )
        ws.cell(row=r, column=8).number_format = CURRENCY_FMT
        ws.cell(
            row=r, column=9,
            value=f'=IFERROR(AVERAGEIFS({trade_pnl},{bs_cur},1,{bs_next},"<>1",{trade_pnl},"<0"),0)',
        )
        ws.cell(row=r, column=9).number_format = CURRENCY_FMT
        ws.cell(row=r, column=10, value=f"=IFERROR(({gross_win})/ABS({gross_loss}),0)")
        ws.cell(row=r, column=10).number_format = RATIO_FMT
        ws.cell(row=r, column=11, value=f"={trade_count}")

    portfolio_row = perf_first_row + len(MONTHS)
    ws.cell(row=portfolio_row, column=1, value="Portfolio (Sum of 8 Months)").font = BOLD
    calc_daily = f"Dashboard_Calc!$J$2:$J${COMMON_REAL_LAST_ROW}"
    calc_cum = f"Dashboard_Calc!$K$2:$K${COMMON_REAL_LAST_ROW}"
    calc_maxdd = f"Dashboard_Calc!$N$2:$N${COMMON_REAL_LAST_ROW}"
    ws.cell(row=portfolio_row, column=2, value=f"=LOOKUP(2,1/ISNUMBER({calc_cum}),{calc_cum})")
    ws.cell(row=portfolio_row, column=2).number_format = CURRENCY_FMT
    p_mean = f"AVERAGE({calc_daily})"
    p_sharpe = f"=IFERROR(({p_mean})/STDEV.P({calc_daily})*SQRT(252),0)"
    ws.cell(row=portfolio_row, column=3, value=p_sharpe)
    ws.cell(row=portfolio_row, column=3).number_format = RATIO_FMT
    # Bare-range comparison/arithmetic operators only (no IFERROR/IF wrapping
    # a multi-cell range) -- see the per-month comment above for why.
    # calc_daily is already error-free by construction (Dashboard_Calc's J
    # column sums 8 already-cleaned per-month columns), so no cleansing is
    # even needed here.
    p_downside_var = f"SUMPRODUCT(({calc_daily}<0)*{calc_daily}^2)/COUNT({calc_daily})"
    p_sortino = f"=IFERROR(({p_mean})/SQRT({p_downside_var})*SQRT(252),0)"
    ws.cell(row=portfolio_row, column=4, value=p_sortino)
    ws.cell(row=portfolio_row, column=4).number_format = RATIO_FMT
    p_max_dd = f"LOOKUP(2,1/ISNUMBER({calc_maxdd}),{calc_maxdd})"
    p_calmar = f"=IFERROR(({p_mean})*252/ABS({p_max_dd}),0)"
    ws.cell(row=portfolio_row, column=5, value=p_calmar)
    ws.cell(row=portfolio_row, column=5).number_format = RATIO_FMT
    ws.cell(row=portfolio_row, column=6, value=f"={p_max_dd}")
    ws.cell(row=portfolio_row, column=6).number_format = CURRENCY_FMT
    for col in (7, 8, 9, 10, 11):
        ws.cell(row=portfolio_row, column=col, value="n/a")
        ws.cell(row=portfolio_row, column=col).alignment = Alignment(horizontal="center")

    # ---------------- Gross-by-leg attribution (secondary) ----------------
    attr_header_row = portfolio_row + 2
    section_title(ws, attr_header_row, "Gross PnL Attribution by Position Tier (secondary detail, not stop-loss-adjusted)")
    attr_col_row = attr_header_row + 1
    attr_headers = ["Month", "Tier 1 PnL", "Tier 2 PnL", "Tier 3 PnL", "Gross Total PnL"]
    for i, h in enumerate(attr_headers, start=1):
        ws.cell(row=attr_col_row, column=i, value=h)
    style_header_row(ws, attr_col_row, 1, len(attr_headers))

    attr_first_row = attr_col_row + 1
    for i, month in enumerate(MONTHS):
        r = attr_first_row + i
        ws.cell(row=r, column=1, value=month)
        for col, letter in ((2, "AZ"), (3, "BE"), (4, "BJ")):
            rng = f"{pm(month)}!${letter}$2:${letter}${PADDED_LAST_ROW}"
            ws.cell(row=r, column=col, value=f"=LOOKUP(2,1/ISNUMBER({rng}),{rng})")
            ws.cell(row=r, column=col).number_format = CURRENCY_FMT
        ws.cell(row=r, column=5, value=f"=B{r}+C{r}+D{r}")
        ws.cell(row=r, column=5).number_format = CURRENCY_FMT

    attr_last_row = attr_first_row + len(MONTHS) - 1

    # ---------------- Year x Month per-metric tables ----------------
    # PnL is rebuilt here to sum straight from Dashboard_Calc's cleaned daily
    # columns rather than each Position Model tab's live 'CD' column: CD has
    # a real pre-existing bug (its year-reset branch discards that year's
    # first trading day of PnL entirely) that silently made this heatmap's
    # yearly totals not tie to the Performance table's since-inception RA
    # Total PnL -- found while investigating exactly that mismatch. Every
    # other table here is new, per-request: Sharpe/Sortino/Calmar/Max
    # Drawdown/Win Rate/Avg Win/Avg Loss/Profit Factor/Trade Count, each
    # broken out by year AND contract month. A "Total" column is included
    # only for PnL and Trade Count -- the only two metrics here that are
    # meaningfully additive across months; summing a Sharpe ratio, win rate,
    # or profit factor across months would be meaningless.
    heat_last_row = build_year_month_table(
        ws, attr_last_row + 2, "Calendar-Year PnL by Contract Month (Risk-Adjusted Basis)",
        _year_month_pnl_formula, CURRENCY_FMT, color_scale="diverging3", total_col=True,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Sharpe by Contract Month (Risk-Adjusted Basis)",
        _year_month_sharpe_formula, RATIO_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Sortino by Contract Month (Risk-Adjusted Basis)",
        _year_month_sortino_formula, RATIO_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Calmar by Contract Month (Risk-Adjusted Basis)",
        _year_month_calmar_formula, RATIO_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Max Drawdown by Contract Month (Risk-Adjusted Basis)",
        _year_month_maxdd_formula, CURRENCY_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Win Rate by Contract Month (Risk-Adjusted Basis)",
        _year_month_winrate_formula, PCT_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Avg $ Win by Contract Month (Risk-Adjusted Basis)",
        _year_month_avgwin_formula, CURRENCY_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Avg $ Loss by Contract Month (Risk-Adjusted Basis)",
        _year_month_avgloss_formula, CURRENCY_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Profit Factor by Contract Month (Risk-Adjusted Basis)",
        _year_month_profitfactor_formula, RATIO_FMT,
    )
    heat_last_row = build_year_month_table(
        ws, heat_last_row + 2, "Calendar-Year Trade Count by Contract Month",
        _year_month_tradecount_formula, "0", color_scale=None, total_col=True,
    )

    # ---------------- Charts ----------------
    build_cumulative_pnl_chart(ws, wb, anchor="N4")
    build_ranking_chart(ws, anchor="N25", perf_first_row=perf_first_row)
    build_pnl_by_year_chart(ws, wb, anchor="N45")
    build_selected_series_chart(ws, wb, anchor="N66")

    for col_letter, width in [("A", 24), ("B", 14), ("C", 14), ("D", 14), ("E", 16),
                               ("F", 12), ("G", 14), ("H", 14), ("I", 14), ("J", 14), ("K", 12),
                               ("N", 18), ("O", 14)]:
        ws.column_dimensions[col_letter].width = width

    return ws


def build_cumulative_pnl_chart(ws, wb, anchor: str) -> None:
    chart = LineChart()
    chart.title = "Cumulative PnL Since Inception (Risk-Adjusted)"
    chart.style = 2
    chart.y_axis.title = "Cumulative PnL ($)"
    chart.y_axis.number_format = "$#,##0"
    chart.x_axis.title = "Date"
    chart.height = 10
    chart.width = 22
    chart.legend.position = "r"
    chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    dates_ws = wb["Position Model - January"]
    cats = Reference(dates_ws, min_col=1, min_row=2, max_row=COMMON_REAL_LAST_ROW)

    for month in MONTHS:
        src_ws = wb[f"Position Model - {month}"]
        data = Reference(src_ws, min_col=74, min_row=2, max_row=COMMON_REAL_LAST_ROW)  # BV=74
        series = Series(data, title=month)
        series.graphicalProperties.line.solidFill = MONTH_COLORS[month]
        series.graphicalProperties.line.width = 15000
        series.marker = Marker(symbol="none")
        series.smooth = False
        chart.series.append(series)

    calc_ws = wb["Dashboard_Calc"]
    port_data = Reference(calc_ws, min_col=11, min_row=2, max_row=COMMON_REAL_LAST_ROW)  # K
    port_series = Series(port_data, title="Portfolio")
    port_series.graphicalProperties.line.solidFill = PORTFOLIO_COLOR
    port_series.graphicalProperties.line.width = 25000
    port_series.marker = Marker(symbol="none")
    port_series.smooth = False
    chart.series.append(port_series)

    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def build_ranking_chart(ws, anchor: str, perf_first_row: int) -> None:
    chart = BarChart()
    chart.type = "col"
    chart.title = "Sharpe Ratio by Contract Month (Risk-Adjusted)"
    chart.y_axis.title = "Sharpe"
    chart.height = 9
    chart.width = 22
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    data = Reference(ws, min_col=3, min_row=perf_first_row, max_row=perf_first_row + len(MONTHS) - 1)
    cats = Reference(ws, min_col=1, min_row=perf_first_row, max_row=perf_first_row + len(MONTHS) - 1)
    series = Series(data, title="Sharpe")
    chart.series.append(series)
    chart.set_categories(cats)

    bar_series = chart.series[0]
    bar_series.graphicalProperties.solidFill = None
    for i, month in enumerate(MONTHS):
        pt = openpyxl.chart.marker.DataPoint(idx=i)
        pt.graphicalProperties.solidFill = MONTH_COLORS[month]
        bar_series.data_points.append(pt)

    ws.add_chart(chart, anchor)


def build_pnl_by_year_chart(ws, wb, anchor: str) -> None:
    calc_ws = wb["Dashboard_Calc"]
    chart = LineChart()
    chart.title = "PnL by Year, Selected Series (Trading-Day-of-Year Overlay)"
    chart.y_axis.title = "Cumulative PnL ($)"
    chart.y_axis.number_format = "$#,##0"
    chart.x_axis.title = "Trading Day of Year"
    chart.height = 10
    chart.width = 22
    chart.legend.position = "r"
    chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    cats = Reference(calc_ws, min_col=17, min_row=2, max_row=2 + MAX_DAY_INDEX)  # Q, no header
    # 261 distinct day-index categories would otherwise each get their own
    # illegible, overlapping tick label -- skip to one label every 20 days.
    chart.x_axis.tickLblSkip = 20
    chart.x_axis.tickMarkSkip = 20
    # Sequential light->dark blue ramp for the five completed years plus black
    # for the current (partial) year, so the in-progress line reads as the
    # reference/foreground line against the historical backdrop.
    year_line_colors = ["9EC5E0", "6BAED6", "3182BD", "08519C", "54278F", "000000"]
    for j, year in enumerate(YEARS):
        col = 18 + j
        data = Reference(calc_ws, min_col=col, min_row=1, max_row=2 + MAX_DAY_INDEX)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.line.solidFill = year_line_colors[j]
        series.graphicalProperties.line.width = 20000 if year != YEARS[-1] else 28000
        series.marker = Marker(symbol="none")
        series.smooth = False
        chart.series.append(series)
    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def build_selected_series_chart(ws, wb, anchor: str) -> None:
    """Single-line cumulative-PnL drill-down driven by the 'Filter Charts By'
    dropdown (Dashboard!O2) via Dashboard_Calc!AB (Selected Series
    Cumulative PnL). Defaults to Portfolio. Complements, rather than
    replaces, the always-all-months overview chart -- this is the "isolate
    one series" view; that one is the "compare all at once" view."""
    chart = LineChart()
    chart.title = "Cumulative PnL, Selected Series (see filter above)"
    chart.style = 2
    chart.y_axis.title = "Cumulative PnL ($)"
    chart.y_axis.number_format = "$#,##0"
    chart.x_axis.title = "Date"
    chart.height = 10
    chart.width = 22
    chart.legend.position = "r"
    chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    dates_ws = wb["Position Model - January"]
    cats = Reference(dates_ws, min_col=1, min_row=2, max_row=COMMON_REAL_LAST_ROW)

    calc_ws = wb["Dashboard_Calc"]
    data = Reference(calc_ws, min_col=28, min_row=1, max_row=COMMON_REAL_LAST_ROW)  # AB, header included
    series = Series(data, title_from_data=True)
    series.graphicalProperties.line.solidFill = PORTFOLIO_COLOR
    series.graphicalProperties.line.width = 20000
    series.marker = Marker(symbol="none")
    series.smooth = False
    chart.series.append(series)
    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def main(target_file: Path) -> None:
    if not target_file.is_file():
        raise SystemExit(f"Workbook not found: {target_file}")
    if target_file.suffix.lower() != ".xlsx":
        raise SystemExit(f"Expected an .xlsx workbook, got: {target_file.name}")

    global COMMON_REAL_LAST_ROW
    COMMON_REAL_LAST_ROW = compute_common_real_last_row(target_file)
    print(f"COMMON_REAL_LAST_ROW computed dynamically: {COMMON_REAL_LAST_ROW}")

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = target_file.with_name(
        target_file.stem + f" - backup before dashboard build {timestamp}.xlsx"
    )
    shutil.copy2(target_file, backup_path)
    print(f"Backup written: {backup_path}")

    wb = openpyxl.load_workbook(target_file)
    build_calc_sheet(wb)
    build_dashboard_sheet(wb)
    wb.save(target_file)
    print(f"Dashboard written to: {target_file}")

    # This project's working directory is OneDrive-synced; opening a
    # just-written large file in a brand-new Excel process immediately after
    # saving has been observed to fail with a COM "disconnected from its
    # clients" error (twice, reproducibly), while the identical open
    # succeeds cleanly once the file has had a few seconds to settle.
    # refresh_via_excel_com's own retry loop re-opens fresh on each attempt,
    # but this head start avoids burning through those retries every run.
    import time
    time.sleep(10)

    refresh_via_excel_com(target_file)
    print("Excel-COM refresh + validation complete.")


def refresh_via_excel_com(path: Path) -> None:
    """Force every formula cell on the two new sheets to be re-tokenized and
    recalculated by Excel, then scan for error cells and save.

    Necessary because openpyxl-written formula cells on these sheets have
    been observed to silently evaluate to a WRONG (but non-error, so
    invisible to an error-cell scan) value the first time Excel opens the
    file -- e.g. a bug-free AVERAGE/STDEV.P formula reading back as exactly
    0 -- and only compute correctly once the identical formula text is
    re-set. Root cause looks like an openpyxl-vs-Excel XML artifact, not a
    formula-logic bug; re-setting every formula cell to itself before
    trusting any value is a cheap, reliable defense against it, and is
    otherwise the same real-recalculation validation pattern used
    throughout this project (see feedback-johansen-gotchas /
    feedback-sboxho-gotchas).
    """
    import subprocess
    import time
    import win32com.client as win32
    import win32process

    def _retry(fn, attempts=8, delay=1.0):
        """Retry a COM call through transient 'Call was rejected by callee'
        errors -- Excel can reject calls made immediately after Open() on a
        large workbook while it's still settling, before it's fully ready to
        accept further automation calls."""
        last_exc = None
        for _ in range(attempts):
            try:
                return fn()
            except Exception as exc:
                last_exc = exc
                time.sleep(delay)
        raise last_exc

    def _running_excel_pids():
        out = subprocess.run(["tasklist"], capture_output=True, text=True).stdout
        return {
            int(line.split()[1])
            for line in out.splitlines()
            if line.upper().startswith("EXCEL.EXE")
        }

    pids_before = _running_excel_pids()
    print(f"Excel PIDs running before automation: {sorted(pids_before)}")

    # DispatchEx -- NOT Dispatch/gencache.EnsureDispatch -- forces a brand-new,
    # separate Excel process rather than attaching to any already-running
    # instance. Required whenever the user may have their own Excel windows
    # open: attaching to an existing instance would apply this function's
    # Visible=False, CalculateFullRebuild() (application-wide, affects every
    # open workbook), and eventual Quit() to the user's own session.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    my_pid = _retry(lambda: win32process.GetWindowThreadProcessId(excel.Hwnd)[1])
    if my_pid in pids_before:
        excel.Quit()
        raise RuntimeError(
            f"Excel automation landed in a pre-existing process (PID {my_pid}) "
            "instead of a fresh one -- aborted before opening the file, to "
            "avoid touching the user's other open Excel instances."
        )
    print(f"Automation isolated to new Excel process PID {my_pid} (pre-existing: {sorted(pids_before)})")

    def _open_and_prime():
        # Open AND touch the workbook as one retry unit: a "disconnected
        # from its clients" COM error (RPC_E_DISCONNECTED) has been seen
        # here right after Open() on a freshly-saved file on this
        # OneDrive-synced folder -- once a workbook handle is in that state,
        # retrying calls on the SAME handle keeps failing identically, so
        # each attempt must close any stale handle and re-open from scratch
        # rather than just re-querying it.
        candidate = excel.Workbooks.Open(str(path))
        try:
            candidate.Worksheets(1).Name  # force-touch to detect disconnection now
        except Exception:
            try:
                candidate.Close(SaveChanges=False)
            except Exception:
                pass
            raise
        return candidate

    wb = _retry(_open_and_prime, attempts=6, delay=5.0)
    try:
        xlCellTypeFormulas = -4123
        xlErrors = 16

        for sheet_name in ("Dashboard", "Dashboard_Calc"):
            ws = _retry(lambda sn=sheet_name: wb.Worksheets(sn))
            try:
                frng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            except Exception:
                continue
            # Bulk read/write per contiguous area (one COM round trip per
            # area, not per cell) -- Dashboard_Calc alone has ~20,000+
            # formula cells, and per-cell .Formula get/set over COM was slow
            # enough to look hung. Range.Formula accepts/returns a 2D tuple
            # for a multi-cell area in a single call.
            for area in frng.Areas:
                try:
                    area.Formula = area.Formula
                except Exception as exc:
                    print(f"WARN: could not re-set {sheet_name}!{area.Address}: {exc}")

        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()

        for sheet_name in ("Dashboard", "Dashboard_Calc"):
            ws = wb.Worksheets(sheet_name)
            try:
                rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
                count = rng.Count
            except Exception:
                count = 0
            if count:
                raise RuntimeError(f"{sheet_name}: {count} error cells after refresh")
            print(f"{sheet_name}: 0 error cells after refresh")

        wb.Save()
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception as exc:
            print(f"WARN: wb.Close raised (already saved above, so likely harmless): {exc}")
        excel.Quit()

    pids_after = _running_excel_pids()
    missing = pids_before - pids_after
    if missing:
        print(f"WARNING: pre-existing Excel PID(s) no longer running: {sorted(missing)} -- check the user's other open files.")
    else:
        print(f"Confirmed: all pre-existing Excel processes ({sorted(pids_before)}) still running.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build the Dashboard tabs into a Knox multi-month workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the Knox workbook (.xlsx) to update")
    main(parser.parse_args().workbook.resolve())
