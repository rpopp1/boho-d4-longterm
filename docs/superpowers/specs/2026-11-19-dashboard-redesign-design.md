# Knox Dashboard redesign — design

Date: 2026-11-19
Status: draft, awaiting approval

## Intent

The Dashboard tab of the Knox multi-month BOHO/D4 workbook has to do two jobs:

1. **Daily check (top screen).** At a glance, one screen and no scrolling:
   where each contract month is positioned, how close each one is to a signal,
   and what volatility regime each one is in.
2. **Performance review (below).** A presentable section the user can show to
   someone else or export, filterable by **contract month** and by **year**.

Today the tab is a functional but hard-to-read build (2026-08-14, auto-extend
fix 2026-11-18). This redesign is about layout, readability and filtering. No
strategy number changes.

## Goals

- The top screen fits one laptop screen (about rows 1–16, columns A–M) and answers
  "what am I holding, how close is the next signal, what's the regime".
- The performance section is driven by two dropdowns, **Month** (Portfolio,
  January … December) and **Year** (All years, 2021 … latest). Every headline
  number, chart and table in the section follows both.
- The section looks presentable: consistent colors, no clipped labels, no
  gridlines, readable heatmaps. It also prints/exports to one landscape page via a set print area.
- Everything keeps auto-extending as rows are pasted (the 2026-11-18 design:
  formulas to row 3001, `Dashboard_Calc!Y1` rows-in-use, named chart ranges).

## Non-goals

- No change to Position Model, Config, or any strategy logic or number.
- No VBA/macros, no PivotTables/slicers (they need manual refresh and can't
  compute Sharpe/drawdown), no dynamic-array spill formulas (project gotcha).
- No separate presentation tab (option B was considered and not chosen).

## Anti-goals (a failure even if it "works")

- Any headline figure that disagrees with today's Dashboard for the same
  selection (e.g. Portfolio/All years must still show $26,050,469, Sharpe 1.91,
  max DD −$1,519,273 on the current Sandbox; December 2026 PnL $947,688).
- A dashboard that goes stale again when rows are pasted.
- The top screen showing a position the strategy isn't actually holding (see
  "Position" below).

## Constraints

- Built by `scripts/build_dashboard.py` (openpyxl write + Excel COM refresh);
  the user's file is the Sandbox (`Knox_Multimonth_BOHO_D4_Strategy - Roll Fix
  Sandbox.xlsx`), which is usually open in Excel — build to a copy, verify,
  then replace with the user's OK and a timestamped backup (existing pattern).
- Formula gotchas from memory: no IFERROR/IF/CHOOSE wrapping a multi-cell
  range inside SUMPRODUCT (implicit intersection), no AGGREGATE (#NAME?), use
  `LOOKUP(2,1/…)` for last-value lookups, plain formulas only.
- Must keep working with the profit-take/re-entry overlay columns (BL:BY,
  EK:EO) — performance uses the effective (RA) columns BT/BU/BV/BY as today.

## Approaches considered

- **A. One tab, top screen + performance section below (chosen by the user).**
  Matches "compact screen at the top" plus a review underneath; a print area
  makes the review section exportable on its own.
- **B. Two tabs (Dashboard + Performance Review).** Cleaner for presenting
  (positions never shown) but a second tab to flip to. Not chosen.
- **C. PivotTables + slicers.** Nice filter UI, but needs manual refresh and
  can't do Sharpe/Sortino/drawdown. Dropped.

## Design

### Top screen — "Today"

Title row with "Data through <date>". One table, one row per month:

| Column | Content |
|---|---|
| Month | |
| Position | **Effective** position: Long / Short / Flat, from the signal (V) gated by Effective Trade Active (BS). When the signal is on but the overlay has exited, show e.g. "Flat (exited)". Found during recon: today the tab shows V, so January, March and December read −1 although BS = 0 (exited, flat). |
| Size | signal position (V), so the tier size is still visible |
| Z-Score | latest L, diverging color, same blue/vermillion pair |
| Next signal | the next Z level that would change the position (next entry tier if flat/partial, the exit level if in a trade), and the distance to it. The builder derives the exact rule from the Position Model's tier/exit formulas (Config rows 14–17) |
| Regime | Regime Label (DC) shown for every state, including Normal; colored Normal/High/Extreme |
| Vol percentile | EWMA Percentile (DB) as a %, plus annualized EWMA vol (CY) |
| Shock | BOHO Shock Flag (DA) shown as a marker when on |
| Open trade PnL | current effective trade PnL (BU) |

Plus a small Portfolio line: net position count long/short/flat, total open trade PnL,
and the share of months in High/Extreme regime.

### Performance section — "Performance Review"

Controls row: **Month** dropdown (Portfolio, 8 months) and **Year** dropdown
(All years, 2021 … latest year with data; options come from a helper list so
new years appear automatically). The existing "Filter Charts By" selector
becomes the Month dropdown.

1. **Headline tiles** (one row of big-number cells): Total PnL, Sharpe,
   Sortino, Max Drawdown, Calmar, Win Rate, Profit Factor, Trade Count — for
   the selected month and year. Portfolio shows n/a for trade stats, as today
   (portfolio trades aren't defined).
2. **Equity curve with drawdown** — cumulative PnL line plus an underwater
   (drawdown) area beneath it, for the selection. "All years" = full history;
   a single year = that year's curve, starting from zero.
3. **PnL by year overlay** — trading-day-of-year cumulative lines, one per year,
   for the selected month (Portfolio by default). Fixes the current bug where lines drop to $0 after
   their last day (blank cells must be #N/A or empty, not 0). The selected year is
   emphasized (thick black); others thin. "All years" = no emphasis.
4. **Month comparison** — bar chart of the chosen metric by contract month
   for the selected year (replaces "Sharpe by Contract Month"). Selected month highlighted.
5. **Year × Month table** — ONE table (years down, months across, + Total
   where additive) with a **Metric** dropdown (PnL, Sharpe, Sortino, Calmar,
   Max DD, Win Rate, Avg Win, Avg Loss, Profit Factor, Trade Count). Replaces
   the ten stacked tables. The selected month column / year row are outlined.
   Heatmap: a single diverging scale centered on a meaningful midpoint (0 for
   PnL/DD/Sharpe-type, 1 for Profit Factor) with light tints so black text
   stays readable; no red font on colored fills.
6. **Per-month summary table** (the current "Performance Since Inception"),
   now for the selected year, kept as the detail table under the charts. The
   gross tier-attribution table moves to the bottom, collapsed (grouped rows).

### Visual system

- One palette: navy headers (current 1F3864), blue/vermillion diverging pair
  (current), month colors from the existing Okabe-Ito set, black = Portfolio.
- Gridlines off, consistent column widths, no clipped headers, axis titles
  that don't overlap ticks (or dropped where the axis is obvious), legends only
  where needed, charts aligned to a grid.
- Remove the empty reserved year rows between sections (the single year
  table grows downward, nothing sits under it except the detail tables).
- Freeze panes under the title row; print area = performance section,
  landscape, fit to 1 page wide.

### Data plumbing (builder's detail)

- Extend `Dashboard_Calc` with the year filter: a selected-year flag column (all
  rows when "All years") feeding the existing mean/variance/drawdown helpers,
  and a year-rebased cumulative/peak/drawdown column for the selected series.
- Month selection keeps the existing CHOOSE-index helper.
- Metric dropdown: the year×month table cells pick from the existing per-year
  formulas via CHOOSE on a metric index.

## Verification

- Portfolio/All years and each month/All years headline figures equal the current
  Dashboard's Performance table on the Sandbox (exact).
- Each month/year cell equals the current ten year×month tables (exact).
- Overlay chart: no drops to zero; legend of selected-series chart shows names.
- Zero error cells on Dashboard/Dashboard_Calc after a full rebuild.
- Paste-forward test: appending a day still updates everything (existing test).
- Screenshots of top screen and review section for the user to review before
  replacing the Sandbox.

## Exited months on the top screen (decided by the user, 2026-11-19)

When the signal is on (V ≠ 0) but the overlay has the month flat (BS = 0),
"Next signal" shows the exit reason and the real way back in, not the Config
entry tier distance:

- Profit-take exit (EN = 1, re-entry switch on): "Flat, profit-take exit.
  Re-entry <first trading day of next month> if Z ≥ <exit Z, EM> (now <Z>)",
  using the overlay's rule: same sign as the exit Z and at least its magnitude
  (shown as "Z ≥ 1.52" or "Z ≤ −1.52").
- Stop exit: "Flat, stopped out. Waiting for signal reset at |Z| ≤ <Config
  row 17> (now <|Z|>)".

Months in a trade, and months with no signal, keep the normal next-threshold
cue.

## Calls made after the builder check (Claude's, for the user to check)

- **Month comparison chart follows the Metric dropdown.** The Metric control
  sits in the controls row next to Month and Year, and drives both the chart and
  the year×month table. Under "All years" the bars show since-inception values,
  the same numbers as the per-month summary table.
- **The year×month table gets a Portfolio column** with year-filtered portfolio
  PnL, Sharpe, Sortino, Calmar and Max DD (from Dashboard_Calc's portfolio daily
  column). Trade metrics show n/a there, as today. The Total column stays for
  PnL and Trade Count only. When Month = Portfolio, the Portfolio column is the
  one outlined.
- **Print area is one page wide and runs as long as it needs:** tiles, the
  three charts and the year table (years with data only). The detail and tier
  tables stay out of the print area.

## Left to the builder

Exact cell positions, tile styling, chart sizes, the "next signal" rule
derived from the formulas, and helper column placement in `Dashboard_Calc`.
