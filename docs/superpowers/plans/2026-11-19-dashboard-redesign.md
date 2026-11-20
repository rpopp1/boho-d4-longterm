# Knox Dashboard Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild the Knox workbook's Dashboard tab as a one-screen "Today" view on top plus a presentable "Performance Review" section driven by Month / Year / Metric dropdowns.

**Architecture:** Every performance number comes from one hidden **metric matrix** in `Dashboard_Calc`. It has 10 metrics × 9 series (Portfolio + 8 months) × 13 periods (All + 2021–2032), and is built from the *existing* formula builders, so the figures can't drift from today's. The Dashboard tiles, the year×month table, the detail table and the month bar chart all `INDEX` into that matrix with the dropdown indices. The "Today" section reads a small per-month block in `Dashboard_Calc` (the last row's values) and builds its text with formula builders that are tested in real Excel.

**Tech Stack:** Python 3, openpyxl (write), pywin32 Excel COM (`excel_com.isolated_workbook`) for recalc/verification, pytest.

**Spec:** `docs/superpowers/specs/2026-11-19-dashboard-redesign-design.md`

## Global Constraints

- No change to any Position Model, Config or strategy formula. Only the `Dashboard` and `Dashboard_Calc` sheets are written.
- Plain formulas only. No dynamic-array/spill formulas, no `AGGREGATE`, no `IF`/`IFERROR`/`CHOOSE` wrapping a multi-cell range inside `SUMPRODUCT`. Last values use `LOOKUP(2,1/…)` or `INDEX(…,MATCH(…))`.
- Auto-extension stays: `Dashboard_Calc` formulas run to row `PADDED_LAST_ROW = 3001`, `Dashboard_Calc!$Y$1` = rows in use, `$Y$2` = latest year.
- Performance uses the risk-adjusted effective columns (BT/BU/BV/BY), as today.
- Colors: navy headers `1F3864`; diverging pair `0072B2` (positive) / `D55E00` (negative); heatmap tints light enough for black text; month colors = existing `MONTH_COLORS`; black = Portfolio.
- Never write to the user's Sandbox until they have seen screenshots and said OK. Build to a copy in the scratchpad.
- Run only the dashboard tests (`scripts/tests/test_build_dashboard.py`, `scripts/tests/test_dashboard_*.py`). The full suite takes about 65 min and has one known unrelated failure.
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. Branch: `dashboard-redesign` (based on `profit-take-reentry`).

## Review Focus

1. **A Year selection where a month has no trades, or a year past the data** → cells show blank / 0, never `#DIV/0!` or `#N/A` on the Dashboard. *(test in Task 2)*
2. **The first pasted day of a new year (2027)** → 2027 appears in the Year dropdown list, the year table and the overlay window, with no rebuild. *(test in Task 2)*
3. **Month = Portfolio** → trade metrics show "n/a" (Trade Count shows the sum of the 8 months) and the month bar chart still shows all 8 months. *(test in Task 2)*
4. **A partially filled trade, or Z exactly on a threshold** → the Next-signal text names the next add tier and the exit, and never says "adds at 99". *(test in Task 4)*
5. **A profit-take exit with the re-entry switch off (Config row 31 = 0)** → "waiting for signal reset", not a re-entry date. *(test in Task 4)*

---

### Task 1: Split the style and metric-formula builders out of build_dashboard.py

Pure refactor: same formulas, same output.

**Files:**
- Create: `scripts/dashboard_style.py`, `scripts/dashboard_metrics.py`
- Modify: `scripts/build_dashboard.py` (remove the moved code, import it instead)
- Test: `scripts/tests/test_dashboard_metrics.py`

**Interfaces:**
- Produces, in `dashboard_style.py`:
  - every existing palette/format constant: `MONTH_COLORS`, `PORTFOLIO_COLOR`, `DIVERGING_POS/NEG/MID`, `HEADER_FILL`, `HEADER_FONT`, `SECTION_FONT`, `BOLD`, `CURRENCY_FMT`, `RATIO_FMT`, `PCT_FMT`
  - `style_header_row(ws, row, first_col, last_col)` and `section_title(ws, row, text)`
  - new constants:
    - `HEAT_CURRENCY_FMT = '$#,##0;($#,##0)'` (no red font)
    - `HEAT_POS = "9DC3E6"`, `HEAT_NEG = "F4B183"`, `HEAT_MID = "FFFFFF"`
    - `TITLE_FONT` (bold 18), `TILE_LABEL_FONT` (9, grey `595959`), `TILE_VALUE_FONT` (bold 16)
- Produces, in `dashboard_metrics.py`:
  - `METRICS: list[tuple[str, str, str]]` = (key, label, number_format), in this order:
    - `pnl` "Total PnL"
    - `sharpe` "Sharpe"
    - `sortino` "Sortino"
    - `calmar` "Calmar"
    - `maxdd` "Max Drawdown"
    - `winrate` "Win Rate"
    - `avgwin` "Avg $ Win"
    - `avgloss` "Avg $ Loss"
    - `pf` "Profit Factor"
    - `trades` "Trade Count"
  - `TRADE_METRICS = {"winrate","avgwin","avgloss","pf","trades"}`
  - `SERIES = ["Portfolio"] + MONTHS`
  - `month_year_formula(metric: str, month: str, year_ref: str) -> str`: dispatches to the moved `_year_month_*` builders, byte-identical output.
  - `month_inception_formula(metric: str, month: str) -> str`: the exact formula text the Performance Since Inception table writes today for that month and column (lift it out of `build_dashboard_sheet` lines 688–756).
  - `portfolio_inception_formula(metric: str) -> str`: today's Portfolio-row text for `pnl`…`maxdd`. Task 2 handles the trade metrics.
  - The helpers `pm`, `calc_range`, `_mean_and_variance` and the constants `CALC_SHEET`, `PADDED_LAST_ROW`, `ROWS_IN_USE`, `LATEST_YEAR`, `VALID_COL`, `FIRST_YEAR`, `TABLE_YEARS`, `YEARDD_FIRST_COL` move here too. `build_dashboard` re-imports them so existing tests keep importing from `build_dashboard`.

- [ ] **Step 1: Write the failing test** `test_dashboard_metrics.py`:
  - `test_metrics_order_and_labels`: assert `[k for k,_,_ in METRICS] == ["pnl","sharpe","sortino","calmar","maxdd","winrate","avgwin","avgloss","pf","trades"]`.
  - `test_month_inception_formulas_match_current_table`: build with the pre-refactor `build_dashboard_sheet` (snapshot the strings now into a fixture dict `EXPECTED_DEC`, taken from the current build: row for December, columns B..K) and assert `month_inception_formula(k, "December") == EXPECTED_DEC[k]` for each of the 10 keys.
  - `test_month_year_formula_dispatch`: `month_year_formula("sharpe","March","$A12") == _year_month_sharpe_formula("March","$A12")` (import the private from `dashboard_metrics`).
- [ ] **Step 2: Run** `pytest scripts/tests/test_dashboard_metrics.py -v`. Expected: FAIL (module not found).
- [ ] **Step 3: Move the code.** Create both modules. `build_dashboard_sheet` calls `month_inception_formula`/`portfolio_inception_formula` instead of inline strings.
- [ ] **Step 4: Run** `pytest scripts/tests/test_dashboard_metrics.py scripts/tests/test_build_dashboard.py -v`. Expected: all PASS.
- [ ] **Step 5: Commit** `git commit -m "Split Dashboard styles and metric formula builders into their own modules"`

---

### Task 2: Metric matrix, Portfolio-by-year metrics and dropdown plumbing in Dashboard_Calc

**Files:**
- Modify: `scripts/dashboard_metrics.py`, `scripts/build_dashboard.py` (`build_calc_sheet`)
- Test: `scripts/tests/test_dashboard_metrics.py`

**Interfaces:**
- Consumes: Task 1 builders.
- Produces, in `dashboard_metrics.py`:
  - Calc columns `PORT_YEARPEAK_COL = "BB"` and `PORT_YEARDD_COL = "BC"`. These are the Portfolio year-reset running peak and running max drawdown, built from O and J with the same shape as the per-month AL/AT columns.
  - `portfolio_year_formula(metric: str, year_ref: str) -> str`:
    - `pnl`/`sharpe`/`sortino`/`calmar`: the `_year_month_*` shapes on column J.
    - `maxdd`: `INDEX` on BC.
    - Trade metrics return `'="n/a"'`, except `trades`, which the matrix sums.
  - Matrix layout:
    - `MATRIX_LABEL_COL = "FA"`, periods in `FB:FN`: FB = All, FC.. = `TABLE_YEARS`.
    - `MATRIX_BLOCK_ROWS = 11`. Metric k (0-based) has its header on row `k*11+1` and series s (1-based, Portfolio = 1) on row `k*11+1+s`.
    - `MATRIX_RANGE = "Dashboard_Calc!$FB$1:$FN$110"`.
  - `matrix_row(metric_index_expr: str, series_index_expr: str) -> str`: the Excel expression `(({m})-1)*11+1+({s})`, with the metric index 1-based.
  - `matrix_lookup(metric_index_expr, series_index_expr, period_index_expr) -> str`: returns `INDEX(MATRIX_RANGE, matrix_row(...), {p})`.
  - `build_metric_matrix(calc_ws) -> None`. Year cells are `=IF({year}>{LATEST_YEAR},"",<formula>)`, where `{year}` is the header cell of the period column, e.g. `FC$1`. Portfolio `trades` cells = `SUM` of the 8 month cells below them.
  - Control cells, all in Calc:
    - `Z1` = month index: MATCH of the Dashboard Month cell against `SERIES`. This replaces the old `'Dashboard'!$O$2` reference.
    - `Z2` = period index: `=IF(<Dashboard Year cell>="All years",1,<year>-FIRST_YEAR+2)`.
    - `Z3` = selected year, or 0 for All.
    - `Z4` = first offset into rows 2..3001: All → 1, year → `MATCH(Z3,A2:A3001,0)`.
    - `Z5` = last offset: All → `$Y$1`, year → `Z4+COUNTIF(A2:A3001,Z3)-1`.
    - `Z6` = metric index: MATCH of the Dashboard Metric cell against the labels.
    - `Z7` = heatmap midpoint (Task 5 fills in its formula).
  - Year dropdown list: `EZ2 = "All years"`, `EZ3:EZ14 = IF(year<=$Y$2,year,"")` for `TABLE_YEARS`.
  - Dashboard control cell addresses are module constants in `build_dashboard.py`: `MONTH_CELL = "B18"`, `YEAR_CELL = "D18"`, `METRIC_CELL = "F18"`. Task 5 places them; this task only references them.

- [ ] **Step 1: Write the failing tests:**
  - `test_matrix_month_cells_reuse_existing_builders`: for December/2024 (period column FF; every block uses the row-1 year headers, `FC$1`=2021 … `FN$1`=2032), the matrix cell's formula == `'=IF(FF$1>Dashboard_Calc!$Y$2,"",' + month_year_formula("sharpe","December","FF$1").removeprefix("=") + ')'`. Its All cell (FB) == `month_inception_formula("sharpe","December")`.
  - `test_portfolio_trade_metrics_are_na_and_trades_sum` *(Review Focus 3)*: Portfolio `winrate` cell == `'="n/a"'`. Portfolio `trades` cell starts with `=SUM(` and covers the 8 month rows of the trades block.
  - `test_year_cells_blank_past_latest_year` *(Review Focus 1)*: every year cell starts with `=IF(` and contains `>Dashboard_Calc!$Y$2,""`.
  - `test_year_options_follow_latest_year` *(Review Focus 2)*: `EZ2 == "All years"`; `EZ3 == '=IF(2021<=$Y$2,2021,"")'`; `EZ14` refers to 2032.
  - `test_matrix_lookup_text`: `matrix_lookup("$Z$6","$Z$1","$Z$2") == "INDEX(Dashboard_Calc!$FB$1:$FN$110,(($Z$6)-1)*11+1+($Z$1),$Z$2)"`.
- [ ] **Step 2: Run** `pytest scripts/tests/test_dashboard_metrics.py -v`. Expected: the new tests FAIL.
- [ ] **Step 3: Implement** the items listed in Interfaces. Call `build_metric_matrix` from `build_calc_sheet`.
- [ ] **Step 4: Run** `pytest scripts/tests/test_dashboard_metrics.py scripts/tests/test_build_dashboard.py -v`. Expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "Add the Dashboard metric matrix and Month/Year/Metric control cells"`

---

### Task 3: Period-aware chart series, overlay fix and the error scan

**Files:**
- Modify: `scripts/build_dashboard.py` (`build_calc_sheet`, `chart_range_names`, `refresh_via_excel_com`)
- Test: `scripts/tests/test_build_dashboard.py`

**Interfaces:**
- Consumes: control cells Z1–Z5 (Task 2).
- Produces:
  - Selected-series columns:
    - `BD`: period-aware cumulative. `=IF($Z$3=0,AB{r},AC{r})`.
    - `BE`: running peak. All: `MAX(prev, BD)`. Year: resets when `A{r}<>A{r-1}`, same convention as AL.
    - `BF`: drawdown, `MIN(0,BD-BE)`.
  - `BG`: "Selected year" overlay column on day-index rows 2..262, with header `BG1 = IF($Z$3=0,"",$Z$3)`. The body is the R:W formula shape for year `$Z$3`, and `NA()` when `$Z$3=0`.
  - Overlay columns R:W return `NA()` instead of `""` past a year's last day. This fixes the drop-to-zero bug.
  - `chart_range_names() -> dict[str,str]` now returns exactly `Chart_SelDates` (January col A), `Chart_SelEquity` (BD) and `Chart_SelDD` (BF). Each is `INDEX(col2:col3001,$Z$4):INDEX(col2:col3001,$Z$5)`. The old names are removed along with their charts in Task 6.
  - `EXPECTED_NA_RANGES = ("R2:W262", "BG2:BG262", "EX2:EY9")`. `EX:EY` is the bar-chart helper from Task 6.
  - `count_unexpected_errors(ws, excluded: tuple[str, ...]) -> int` counts error cells outside the excluded ranges, using `Application.Intersect` per error area. `refresh_via_excel_com` uses it for `Dashboard_Calc`; Dashboard stays at strict zero.
- [ ] **Step 1: Write the failing tests** (replacing `test_date_charts_plot_named_ranges_that_end_at_the_last_row_in_use`):
  - `test_chart_names_follow_the_selected_period`: `set(chart_range_names()) == {"Chart_SelDates","Chart_SelEquity","Chart_SelDD"}`, and `Chart_SelEquity == "Dashboard_Calc!INDEX(Dashboard_Calc!$BD$2:$BD$3001,Dashboard_Calc!$Z$4):INDEX(Dashboard_Calc!$BD$2:$BD$3001,Dashboard_Calc!$Z$5)"`. Match the existing `dynamic_range_name` quoting style.
  - `test_overlay_gaps_instead_of_zero`: `calc["R100"].value` contains `NA()` and not `""`.
  - `test_selected_year_overlay_column`: `calc["BG50"].value` starts with `=IF($Z$3=0,NA(),`.
- [ ] **Step 2: Run** the dashboard tests. Expected: the new tests FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** the dashboard tests. Expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "Make Dashboard chart series follow the selected month and year; gap overlay lines past their last day"`

---

### Task 4: "Today" section

**Files:**
- Create: `scripts/dashboard_today.py`
- Modify: `scripts/build_dashboard.py` (call it, drop the old Current Snapshot)
- Test: `scripts/tests/test_dashboard_today.py` (structure) and `scripts/tests/test_dashboard_today_com.py` (real-Excel evaluation of the text formulas)

**Interfaces:**
- Produces:
  - `TODAY_FIELDS: dict[str, str]` maps a field to its Calc column, one row per month, rows 2..9 in `MONTHS` order:
    - `asof` CA, `row` CB, `pos` CC (V), `active` CD (BS), `z` CE (L)
    - `pt_exit` CF (EN), `exit_z` CG (EM), `regime` CH (DC), `pctl` CI (DB), `vol` CJ (CY), `shock` CK (DA), `trade_pnl` CL (BU)
    - legs: `s1` CM (M), `s2` CN (N), `s3` CO (O), `l3` CP (P), `l2` CQ (Q), `l1` CR (R)
    - `reentry` CS (Config row 31)
    - tiers: `t1` CT (row 14), `t2` CU (row 15), `t3` CV (row 16)
    - exits: `exit_short` CW (row 17), `exit_long` CX (`'Position Model - <m>'!$DV$7`)
  - How the block is filled:
    - `asof` uses the existing `LOOKUP(2,1/(A>DATE(2015,1,1)),A)`.
    - `row` = `MATCH(asof, A2:A3001, 0)`.
    - Every other Position Model field is `INDEX(<col>$2:<col>$3001, row)`. Config fields reference the month's Config column.
  - `build_today_block(calc_ws) -> None`
  - `position_formula(refs: dict[str,str]) -> str`: `"Flat"` when pos=0; `"Long"`/`"Short"` when active=1; `"Flat (exited)"` otherwise.
  - `next_signal_formula(refs: dict[str,str]) -> str`. `refs` maps every `TODAY_FIELDS` key to a cell address. Numbers use `TEXT(x,"0.00")`. Cases:
    - **Flat, no signal:** `Short at Z ≥ {min tier} (now {z})` if z ≥ 0, else `Long at Z ≤ -{min tier} (now {z})`.
    - **Short in trade:** `Exit at Z ≤ {exit_short} (now {z})`. If any short leg is 0, append ` · adds at Z ≥ {lowest unfilled tier}`, where s1↔t1, s2↔t2, s3↔t3.
    - **Long in trade:** mirror. `Exit at Z ≥ {exit_long}`, and adds at `Z ≤ -{tier}` for l1/l2/l3.
    - **Exited, profit-take, re-entry on:** `Flat, profit-take exit. Re-entry {m/d of WORKDAY(EOMONTH(asof,0),1)} if Z ≥ {exit_z} (now {z})`. Use `≤` when exit_z < 0.
    - **Exited otherwise:** `Flat, stopped out. Waiting for signal reset at Z ≤ {exit_short} (now {z})` for a short (`≥ {exit_long}` for a long). If pt_exit=1 but re-entry is off, the text starts `Flat, profit-take exit. Waiting for signal reset…` instead.
  - `build_today_section(ws, first_row: int = 4) -> int` returns the last row used.
    - Layout: title row, header row, 8 month rows, Portfolio row.
    - Columns: A Month, B Position, C Signal, D Z-Score, E Regime, F Vol Pctl, G EWMA Vol, H Shock, I Open Trade PnL, J Next Signal (width about 60).
    - Open Trade PnL is blank when not active.
    - Portfolio row: B = `"x long · y short · z flat"`, I = sum of open PnL, E = `"n of 8 High/Extreme"`.
    - Formatting:
      - Z-score: diverging 3-color, using `HEAT_*` tints.
      - Regime: Normal no fill, High `F4C7A8`, Extreme `D55E00` with white bold text.
      - Shock = 1 shows "⚡ shock".
- [ ] **Step 1: Write the failing tests.**
  - `test_dashboard_today.py`:
    - `test_today_block_reads_the_last_row`: Calc `CC2 == "=INDEX('Position Model - January'!$V$2:$V$3001,$CB2)"` and `CS9 == "=Config!$I$31"`.
    - `test_today_rows_in_month_order`: Dashboard A6:A13 == MONTHS.
  - `test_dashboard_today_com.py`: build a scratch workbook with openpyxl. It has one input row per case, holding `TODAY_FIELDS` values with tiers 2.25/2.1/1.75 and exits 0.1/-0.1. Column Z holds `next_signal_formula` and column Y holds `position_formula`. Recalculate in `isolated_workbook`, then assert:
    - flat, z=0.74 → `"Short at Z ≥ 1.75 (now 0.74)"`
    - short, all legs on, z=1.52 → `"Exit at Z ≤ 0.10 (now 1.52)"`
    - short, only s3 on, z=1.80 → `"Exit at Z ≤ 0.10 (now 1.80) · adds at Z ≥ 2.10"` *(Review Focus 4)*
    - long, z=-2.10 exactly on t2 with l3/l2 on → `"… · adds at Z ≤ -2.25"`, and the output contains no `99`
    - exited PT, reentry=1, asof 2026-11-18, exit_z=1.52 → `"Flat, profit-take exit. Re-entry 12/1 if Z ≥ 1.52 (now 1.52)"`
    - exited PT, reentry=0 → starts with `"Flat, profit-take exit. Waiting for signal reset"` *(Review Focus 5)*
    - exited stop, short → `"Flat, stopped out. Waiting for signal reset at Z ≤ 0.10 (now 0.74)"`
    - position text for (pos=-1, active=0) → `"Flat (exited)"`
- [ ] **Step 2: Run** `pytest scripts/tests/test_dashboard_today.py scripts/tests/test_dashboard_today_com.py -v`. Expected: FAIL.
- [ ] **Step 3: Implement** `dashboard_today.py` and wire it in.
- [ ] **Step 4: Run** the same tests plus `test_build_dashboard.py`. Expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "Add the Dashboard Today section: effective position, next-signal cue, regime"`

---

### Task 5: Performance Review section: controls, tiles, year table, detail tables, page setup

**Files:**
- Create: `scripts/dashboard_review.py`
- Modify: `scripts/build_dashboard.py` (`build_dashboard_sheet` becomes the title, then `build_today_section`, `build_review_section`, `build_review_charts`; the old tables are removed)
- Test: `scripts/tests/test_dashboard_review.py`; update `test_build_dashboard.py` (drop `test_calendar_year_tables_reserve_rows_through_the_last_table_year`, whose job moves here)

**Interfaces:**
- Consumes: `matrix_lookup`, `METRICS`, `SERIES`, control cells (Task 2), `build_today_section` (Task 4).
- Produces:
  - `build_review_section(ws, first_row: int) -> dict[str,int]`, which returns the anchors `{"charts": row, "year_table": row, "detail": row, "end": row}` for Task 6 and the print area.
  - Row 17: "Performance Review" title. Row 18: the control labels and dropdowns.
    - `MONTH_CELL` B18: list = SERIES, default "Portfolio".
    - `YEAR_CELL` D18: list = `=Dashboard_Calc!$EZ$2:$EZ$14`, default "All years".
    - `METRIC_CELL` F18: list = METRICS labels, default "Total PnL".
    - A caption in H18 such as `="Portfolio · All years"`.
  - Tiles on rows 20 (label) and 21 (value), columns A..H, one per metric except Avg Win/Avg Loss: pnl, sharpe, sortino, maxdd, calmar, winrate, pf, trades.
    - Value = `=matrix_lookup(<k+1>,"Dashboard_Calc!$Z$1","Dashboard_Calc!$Z$2")`, formatted with the metric's format.
  - Year table starts at `anchors["year_table"]`:
    - title, then header Year | 8 months | Portfolio (A..J)
    - rows for `TABLE_YEARS`, with `A = IF(year<=LATEST_YEAR,year,"")`
    - each cell is `IF($A="","",INDEX(MATRIX_RANGE, matrix_row("Dashboard_Calc!$Z$6", <s>), $A-FIRST_YEAR+2))`, where s is 2..9 for the months and 1 for Portfolio
  - Heatmap on B:J of the year rows:
    - 3-color scale, min `HEAT_NEG`, mid `HEAT_MID` with cfvo type `formula` `Dashboard_Calc!$Z$7`, max `HEAT_POS`.
    - `Z7 = CHOOSE($Z$6,0,0,0,0,MEDIAN(<year table B:J>),0.5,MEDIAN(…),MEDIAN(…),1,MEDIAN(…))`.
    - Number format comes from formula rules `=$F$18="<label>"`, one per metric, each setting that metric's format (`HEAT_CURRENCY_FMT` for $ metrics).
    - Outline rules give a thick border on the selected month column (`COLUMN()-1 = Dashboard_Calc!$Z$1`, with J ↔ Portfolio) and on the selected year row (`$A=Dashboard_Calc!$Z$3`).
  - Detail table: "By Contract Month — <period>". Rows are Portfolio + 8 months, columns are the 10 METRICS labels (A..K), and each cell is `matrix_lookup(k, s, "Dashboard_Calc!$Z$2")`.
  - Tier attribution table: today's formulas unchanged, under the detail table, rows grouped (`outline_level=1`, hidden).
  - Page setup:
    - `showGridLines=False`, freeze panes at `A3`.
    - Print area `A17:J<year table last row>`, landscape, `fitToWidth=1`, `fitToHeight=0`.
    - Column widths: A 22, B–I 15, J 60, K 15.
- [ ] **Step 1: Write the failing tests** (`test_dashboard_review.py`, built with openpyxl as in the existing fixture):
  - `test_controls_have_dropdowns_and_defaults`: B18 == "Portfolio", D18 == "All years", F18 == "Total PnL"; each is covered by a DataValidation whose `formula1` is as specified.
  - `test_tiles_index_the_matrix`: the Sharpe tile value == `"=" + matrix_lookup("2","Dashboard_Calc!$Z$1","Dashboard_Calc!$Z$2")`.
  - `test_year_table_rows_and_portfolio_column`: the first year row's A == `'=IF(2021<=Dashboard_Calc!$Y$2,2021,"")'`; J of that row contains `+(1),` (Portfolio series 1).
  - `test_heatmap_has_no_red_font_and_formula_midpoint`: the year table's color-scale rule has mid cfvo `type == "formula"`; no CF rule or cell number format in the table contains `[RED]`.
  - `test_print_area_and_gridlines`: `ws.print_area` starts with `'Dashboard'!$A$17`; `ws.sheet_view.showGridLines is False`.
- [ ] **Step 2: Run** `pytest scripts/tests/test_dashboard_review.py -v`. Expected: FAIL.
- [ ] **Step 3: Implement** `dashboard_review.py` and rewire `build_dashboard_sheet`.
- [ ] **Step 4: Run** all dashboard tests. Expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "Add the Performance Review section: Month/Year/Metric filters, tiles, one year table"`

---

### Task 6: Review charts

**Files:**
- Modify: `scripts/dashboard_review.py`, `scripts/build_dashboard.py` (delete `build_cumulative_pnl_chart`, `build_ranking_chart`, `build_selected_series_chart` and the old `build_pnl_by_year_chart`)
- Test: `scripts/tests/test_dashboard_review.py`

**Interfaces:**
- Consumes: `chart_range_names` (Task 3); `anchors["charts"]` (Task 5); Calc BG and R:W.
- Produces `build_review_charts(ws, wb, anchor_row: int) -> None`, which builds four charts with no y-axis titles (number formats carry the $) and no x-axis title on date charts:
  1. **Equity:** line `Chart_SelEquity` on `Chart_SelDates`, black, 2.25pt. Anchor `A{anchor_row}`, spanning A..J, height 9cm.
  2. **Drawdown:** area `Chart_SelDD`, fill `D55E00` at about 40% (use a lighter solid `F4B183` if transparency isn't supported), directly under the equity chart, same width, height 5cm.
  3. **PnL by year overlay:** R:W in a grey ramp (`D9D9D9`→`7F7F7F`, 1pt), plus BG as "Selected year" in black 2.5pt. `dispBlanksAs="gap"`. Anchored left half under the drawdown chart.
  4. **Month comparison:** clustered column with `overlap=100`, `gapWidth=60`.
     - Series "Other" = `Dashboard_Calc!$EX$2:$EX$9` in grey `BFBFBF`.
     - Series "Selected" = `$EY$2:$EY$9` in `0072B2`.
     - Categories `$EW$2:$EW$9` (month names).
     - Anchored right half beside the overlay.
     - Helper cells built here in Calc: `EX{2+i} = IF($Z$1={i+2},NA(),v)` and `EY{2+i} = IF($Z$1={i+2},v,NA())`, where `v = matrix_lookup("$Z$6", i+2, "$Z$2")`.
- [ ] **Step 1: Write the failing test** `test_review_charts`:
  - The Dashboard has exactly 4 charts.
  - The plotted refs include `Dashboard_Calc!Chart_SelEquity` and `Dashboard_Calc!Chart_SelDD`.
  - One chart has a series over `$BG$`.
  - The bar chart has 2 series over `$EX$`/`$EY$` with `overlap == 100`.
- [ ] **Step 2: Run** it. Expected: FAIL.
- [ ] **Step 3: Implement** the charts.
- [ ] **Step 4: Run** all dashboard tests. Expected: PASS.
- [ ] **Step 5: Commit** `git commit -m "Add Performance Review charts: equity with drawdown, year overlay, month comparison"`

---

### Task 7: Build on a copy, verify against today's Dashboard, show the user

**Files:**
- Create (scratchpad, one-off): `<scratchpad>/verify_redesign.py`
- No repo code changes unless verification finds a bug. If it does, fix it in the owning task's module with a test.

- [ ] **Step 1: Capture the old values.** Copy the Sandbox to `<scratchpad>/redesign_old.xlsx` and read, with openpyxl `data_only` and `iter_rows(values_only=True)`:
  - the current Performance Since Inception table (9 rows × 10 metrics)
  - the ten year×month tables (6 years × 8 months each)
  Save them as JSON.
- [ ] **Step 2: Build.** Copy the Sandbox to `<scratchpad>/redesign_new.xlsx` and run `python scripts/build_dashboard.py "<scratchpad>/redesign_new.xlsx"`. Expected: "Dashboard: 0 error cells", and Dashboard_Calc reports 0 *unexpected* error cells.
- [ ] **Step 3: Compare.** Using `isolated_workbook`, read the matrix `FB1:FN110`:
  - Every old value must equal the matching matrix value, to 1e-6 relative.
  - For the Portfolio by year Sharpe/Sortino/Calmar/Max DD there is no old value. Check these against a pandas recomputation from Calc column J instead.
  - Spot anchors: Portfolio/All PnL = 26,050,469; Sharpe 1.91; max DD −1,519,273; December/2026 PnL = 947,688.
- [ ] **Step 4: Drive the dropdowns.** For (Portfolio, All years), (December, 2026), (May, 2024) and (Portfolio, 2025):
  - Set B18/D18/F18 and recalculate.
  - Assert each tile equals its matrix cell.
  - Assert `Z4:Z5` cover exactly that year's rows.
  - Assert the Dashboard has 0 error cells.
- [ ] **Step 5: Check Today.** For each month, compare the Today row with the raw last row: V, BS, L, DC. January, March and December must read "Flat (exited)". Print the Next Signal texts.
- [ ] **Step 6: Screenshots.** Snapshot `A1:J16` and the review section to PNG (CopyPicture → ChartObject.Export, as done on 11/19) and look at them. Then show the user the screenshots and the comparison summary.
- [ ] **Step 7: After the user OKs:**
  - Ask them to close the Sandbox in Excel.
  - Run `build_dashboard.py` on the Sandbox. It writes its own timestamped backup.
  - Re-run Step 3's anchors on it.
  - Push the branch and open a PR to master. Its body notes that it stacks on `profit-take-reentry`.
  - Update the handoff doc and the `project_knox_dashboard_staleness` memory.
