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
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.chart.data_source import AxDataSource, NumRef
from openpyxl.chart.series import SeriesLabel, StrRef
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from config_map import CONFIG_MONTH_COLUMNS
from multi_month import MONTHS
from dashboard_metrics import (  # noqa: F401  (re-exported for tests and callers)
    CALC_SHEET, CHART_YEAR_COUNT, FIRST_YEAR, LAST_TABLE_YEAR, LATEST_YEAR, MAX_DAY_INDEX,
    PADDED_LAST_ROW, ROWS_IN_USE, TABLE_YEARS, VALID_COL, YEARCUM_FIRST_COL, YEARDD_FIRST_COL,
    YEARPEAK_FIRST_COL, _mean_and_variance, _year_month_avgloss_formula, _year_month_avgwin_formula,
    _year_month_calmar_formula, _year_month_maxdd_formula, _year_month_pnl_formula,
    _year_month_profitfactor_formula, _year_month_sharpe_formula, _year_month_sortino_formula,
    _year_month_tradecount_formula, _year_month_winrate_formula, calc_range, config_tier_cell,
    dynamic_range_name, METRIC_KEYS, month_inception_formula, pm, portfolio_inception_formula,
)
from dashboard_metrics import (
    PORT_YEARDD_COL, PORT_YEARPEAK_COL, build_controls, build_metric_matrix,
)
from dashboard_today import build_today_block, build_today_section
from dashboard_style import (
    BOLD, CURRENCY_FMT, DIVERGING_MID, DIVERGING_NEG, DIVERGING_POS, HEADER_FILL, HEADER_FONT,
    MONTH_COLORS, PCT_FMT, PORTFOLIO_COLOR, RATIO_FMT, SECTION_FONT, section_title, style_header_row,
)



# Dropdown options for the "Filter Charts By" selector, in a fixed order that
# every CHOOSE(...)-based helper formula below matches positionally: index 1
# = Portfolio, indices 2-9 = the 8 months in the same order as MONTHS.
SELECTOR_OPTIONS = ["Portfolio"] + MONTHS


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
    for j, year in enumerate(TABLE_YEARS):
        r = first_row + j
        # The row shows its year (and its numbers) once the data reaches it.
        ws.cell(row=r, column=1, value=f'=IF({year}<={LATEST_YEAR},{year},"")')
        year_ref = f"$A{r}"
        for i, month in enumerate(MONTHS):
            col = 2 + i
            body = cell_formula_fn(month, year_ref).removeprefix("=")
            ws.cell(row=r, column=col, value=f'=IF({year_ref}="","",{body})')
            ws.cell(row=r, column=col).number_format = number_format
        if total_col:
            first_month_col = get_column_letter(2)
            last_month_col = get_column_letter(1 + len(MONTHS))
            tc = 2 + len(MONTHS)
            ws.cell(row=r, column=tc,
                    value=f'=IF({year_ref}="","",SUM({first_month_col}{r}:{last_month_col}{r}))')
            ws.cell(row=r, column=tc).number_format = number_format
            ws.cell(row=r, column=tc).font = BOLD

    last_row = first_row + len(TABLE_YEARS) - 1
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
    ws[f"{VALID_COL}1"] = "Row In Use (all 8 months dated)"
    ws["X1"] = "Rows In Use"
    ws["Y1"] = f"=SUM({VALID_COL}2:{VALID_COL}{PADDED_LAST_ROW})"
    ws["X2"] = "Latest Year"
    ws["Y2"] = f"=MAX(A2:A{PADDED_LAST_ROW})"

    for r in range(2, PADDED_LAST_ROW + 1):
        # In use once every month's tab shows a real trading date on this row
        # (the same >DATE(2015,1,1) test the Snapshot's As Of uses). Rows
        # past the data are 0 and contribute nothing anywhere below.
        dated = ",".join(f"{pm(m)}!A{r}>DATE(2015,1,1)" for m in MONTHS)
        ws[f"{VALID_COL}{r}"] = f"=IFERROR(IF(AND({dated}),1,0),0)"
        ws.cell(row=r, column=1, value=f"=IF(${VALID_COL}{r}=1,{pm('January')}!CO{r},0)")
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
            ws.cell(row=r, column=2 + i, value=f"=IF(${VALID_COL}{r}=1,IF(ISNUMBER({src}),{src},0),0)")
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
    build_controls(ws)  # Z1:Z6 dropdown indices + Year option list
    ws["AA1"] = "Selected Series Daily PnL"
    ws["AB1"] = "Selected Series Cumulative PnL"
    ws["AC1"] = "Selected Series Calendar-Year Cumulative PnL"
    daily_choices = ",".join(["J{r}"] + [get_column_letter(2 + i) + "{r}" for i in range(len(MONTHS))])
    for r in range(2, PADDED_LAST_ROW + 1):
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
    year_range = f"$A$2:$A${PADDED_LAST_ROW}"
    cum_range = f"$AC$2:$AC${PADDED_LAST_ROW}"
    # The latest CHART_YEAR_COUNT years, oldest first (R=18 .. W=23): the
    # window slides forward on its own when the first day of a new year is
    # pasted, so the newest (black, current-year) line is always last.
    for j in range(CHART_YEAR_COUNT):
        col = 18 + j
        ws.cell(row=1, column=col, value=f"=$Y$2-{CHART_YEAR_COUNT - 1 - j}")
    for k in range(0, MAX_DAY_INDEX + 1):
        r = 2 + k
        ws.cell(row=r, column=17, value=k)
        for j in range(CHART_YEAR_COUNT):
            col = 18 + j
            col_letter = get_column_letter(col)
            year_cell = f"{col_letter}$1"
            formula = (
                f"=IF($Q{r}=0,0,"
                f"IF($Q{r}>COUNTIF({year_range},{year_cell}),NA(),"
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
        for r in range(2, PADDED_LAST_ROW + 1):
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

    # Portfolio year-reset running peak / running max drawdown (BB:BC), same
    # shape as the per-month AL/AT columns, for the Portfolio-by-year metrics.
    ws[f"{PORT_YEARPEAK_COL}1"] = "Portfolio Year-Reset Running Peak"
    ws[f"{PORT_YEARDD_COL}1"] = "Portfolio Year-Reset Max Drawdown (running)"
    pk, dd = PORT_YEARPEAK_COL, PORT_YEARDD_COL
    for r in range(2, PADDED_LAST_ROW + 1):
        if r == 2:
            ws[f"{pk}2"] = "=O2"
            ws[f"{dd}2"] = f"=MIN(0,O2-{pk}2)"
        else:
            ws[f"{pk}{r}"] = f"=IF(A{r}<>A{r - 1},O{r},MAX({pk}{r - 1},O{r}))"
            ws[f"{dd}{r}"] = f"=IF(A{r}<>A{r - 1},MIN(0,O{r}-{pk}{r}),MIN({dd}{r - 1},O{r}-{pk}{r}))"

    # Selected series over the selected period (BD:BF) for the equity and
    # drawdown charts: All years = since-inception cumulative (AB); a single
    # year = that year's year-reset cumulative (AC), peak resetting each year.
    ws["BD1"] = "Selected Series Cumulative PnL (Selected Period)"
    ws["BE1"] = "Selected Series Running Peak (Selected Period)"
    ws["BF1"] = "Selected Series Drawdown (Selected Period)"
    for r in range(2, PADDED_LAST_ROW + 1):
        ws[f"BD{r}"] = f"=IF($Z$3=0,AB{r},AC{r})"
        ws[f"BE{r}"] = "=BD2" if r == 2 else f"=IF(AND($Z$3<>0,A{r}<>A{r - 1}),BD{r},MAX(BE{r - 1},BD{r}))"
        ws[f"BF{r}"] = f"=MIN(0,BD{r}-BE{r})"

    # "Selected year" overlay line (BG, day-index rows like R:W): the chosen
    # year drawn on top of the grey history lines; #N/A (no line) for All years.
    ws["BG1"] = '=IF($Z$3=0,"",$Z$3)'
    for k in range(0, MAX_DAY_INDEX + 1):
        r = 2 + k
        ws[f"BG{r}"] = (
            f"=IF($Z$3=0,NA(),IF($Q{r}=0,0,IF($Q{r}>COUNTIF({year_range},$Z$3),NA(),"
            f"INDEX({cum_range},MATCH($Z$3,{year_range},0)+$Q{r}-1))))"
        )

    build_metric_matrix(ws)
    build_today_block(ws)

    # Sheet-scoped names the date-axis charts plot, each ending at the last
    # row in use, so the charts grow as rows are pasted.
    for name, text in chart_range_names().items():
        ws.defined_names[name] = DefinedName(name, attr_text=text)

    return ws


def _period_range_name(col: str, sheet: str = CALC_SHEET) -> str:
    """A column's rows for the selected period: Dashboard_Calc!Z4..Z5 are the
    first/last offsets into rows 2..PADDED_LAST_ROW (all rows in use when the
    Year dropdown says All years)."""
    ref = f"'{sheet}'" if " " in sheet else sheet
    full = f"{ref}!${col}$2:${col}${PADDED_LAST_ROW}"
    return f"INDEX({full},{CALC_SHEET}!$Z$4):INDEX({full},{CALC_SHEET}!$Z$5)"


# Dashboard_Calc ranges that hold #N/A by design (chart gaps), skipped by the
# post-build error scan.
EXPECTED_NA_RANGES = ("R2:W262", "BG2:BG262", "EX2:EY9")


def errors_outside(addresses: list[str], excluded: tuple[str, ...]) -> list[str]:
    """The cell addresses not inside any of the `excluded` ranges."""
    boxes = [range_boundaries(rng) for rng in excluded]
    out = []
    for addr in addresses:
        col, row = coordinate_to_tuple(addr)[1], coordinate_to_tuple(addr)[0]
        if not any(c1 <= col <= c2 and r1 <= row <= r2 for c1, r1, c2, r2 in boxes):
            out.append(addr)
    return out


def chart_range_names() -> dict[str, str]:
    names = {
        "Chart_SelDates": _period_range_name("A", "Position Model - January"),
        "Chart_SelEquity": _period_range_name("BD"),
        "Chart_SelDD": _period_range_name("BF"),
    }
    # Legacy names for the pre-redesign charts (removed with them in Task 6).
    names["Chart_Dates"] = dynamic_range_name("A", "Position Model - January")
    for month in MONTHS:
        names[f"Chart_{month}"] = dynamic_range_name("BV", f"Position Model - {month}")  # RA cumulative
    names["Chart_Portfolio"] = dynamic_range_name("K")
    names["Chart_Selected"] = dynamic_range_name("AB")
    return names


def _plot_name(series: Series, name: str) -> None:
    """Point a chart series' values at a Dashboard_Calc-scoped name."""
    series.val.numRef.f = f"{CALC_SHEET}!{name}"
    series.val.numRef.numCache = None


def _plot_categories(chart) -> None:
    for series in chart.series:
        series.cat = AxDataSource(numRef=NumRef(f=f"{CALC_SHEET}!Chart_Dates"))


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

    # ---------------- Today ----------------
    last_cockpit_row = build_today_section(ws, first_row=4)

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
    fmts = [CURRENCY_FMT, RATIO_FMT, RATIO_FMT, RATIO_FMT, CURRENCY_FMT,
            PCT_FMT, CURRENCY_FMT, CURRENCY_FMT, RATIO_FMT, None]
    for i, month in enumerate(MONTHS):
        r = perf_first_row + i
        ws.cell(row=r, column=1, value=month)
        for j, key in enumerate(METRIC_KEYS):
            cell = ws.cell(row=r, column=2 + j, value=month_inception_formula(key, month))
            if fmts[j]:
                cell.number_format = fmts[j]

    portfolio_row = perf_first_row + len(MONTHS)
    ws.cell(row=portfolio_row, column=1, value="Portfolio (Sum of 8 Months)").font = BOLD
    for j, key in enumerate(METRIC_KEYS[:5]):
        cell = ws.cell(row=portfolio_row, column=2 + j, value=portfolio_inception_formula(key))
        cell.number_format = fmts[j]
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

    # Each series is built on a placeholder range, then pointed at its named
    # range (chart_range_names), which ends at the last row in use.
    calc_ws = wb[CALC_SHEET]
    placeholder = Reference(calc_ws, min_col=11, min_row=2, max_row=3)

    for month in MONTHS:
        series = Series(placeholder, title=month)
        _plot_name(series, f"Chart_{month}")
        series.graphicalProperties.line.solidFill = MONTH_COLORS[month]
        series.graphicalProperties.line.width = 15000
        series.marker = Marker(symbol="none")
        series.smooth = False
        chart.series.append(series)

    port_series = Series(placeholder, title="Portfolio")
    _plot_name(port_series, "Chart_Portfolio")
    port_series.graphicalProperties.line.solidFill = PORTFOLIO_COLOR
    port_series.graphicalProperties.line.width = 25000
    port_series.marker = Marker(symbol="none")
    port_series.smooth = False
    chart.series.append(port_series)

    _plot_categories(chart)
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
    # Sequential light->dark ramp for the earlier years plus black for the
    # latest (usually partial) year, so the in-progress line reads as the
    # reference/foreground line against the historical backdrop. The year
    # headers (R1:W1) are formulas, so the window slides with the data.
    year_line_colors = ["9EC5E0", "6BAED6", "3182BD", "08519C", "54278F", "000000"]
    assert len(year_line_colors) == CHART_YEAR_COUNT
    for j in range(CHART_YEAR_COUNT):
        col = 18 + j
        data = Reference(calc_ws, min_col=col, min_row=1, max_row=2 + MAX_DAY_INDEX)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.line.solidFill = year_line_colors[j]
        series.graphicalProperties.line.width = 20000 if j != CHART_YEAR_COUNT - 1 else 28000
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

    calc_ws = wb[CALC_SHEET]
    placeholder = Reference(calc_ws, min_col=28, min_row=2, max_row=3)  # AB
    series = Series(placeholder)
    series.tx = SeriesLabel(strRef=StrRef(f"{CALC_SHEET}!$AB$1"))
    _plot_name(series, "Chart_Selected")
    series.graphicalProperties.line.solidFill = PORTFOLIO_COLOR
    series.graphicalProperties.line.width = 20000
    series.marker = Marker(symbol="none")
    series.smooth = False
    chart.series.append(series)
    _plot_categories(chart)
    ws.add_chart(chart, anchor)


def main(target_file: Path) -> None:
    if not target_file.is_file():
        raise SystemExit(f"Workbook not found: {target_file}")
    if target_file.suffix.lower() != ".xlsx":
        raise SystemExit(f"Expected an .xlsx workbook, got: {target_file.name}")

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
                addresses = [c.Address.replace("$", "") for area in rng.Areas for c in area.Cells]
            except Exception:
                addresses = []
            excluded = EXPECTED_NA_RANGES if sheet_name == CALC_SHEET else ()
            count = len(errors_outside(addresses, excluded))
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
