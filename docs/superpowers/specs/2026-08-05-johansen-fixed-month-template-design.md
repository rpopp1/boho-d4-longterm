# Johansen Fixed-Month BOHO/D4 Template Engine — Design

## Context

The live `Johansen_BOHO_D4_Strategy_v08.5 - Live w D4 Hedge.xlsx` workbook runs a Johansen-cointegration mean-reversion strategy between BOHO (SBO/HO crack, $/gal) and the D4 RIN price, specifically on the December BOHO contract, and is actively trading. It has proven out the core approach: a rolling-window Johansen re-estimation, refreshed monthly and applied with a one-month lag (walk-forward, no look-ahead), feeding a 3-tier z-score entry/exit position model with per-leg dollar PnL, stop-loss, EWMA volatility regime detection, and drawdown/PnL tracking.

The next phase of this strategy is to run the same fixed-month approach on one or two of the other ZL-tradeable BOHO contract months (Jan, Mar, May, Jul, Aug, Sep, Oct — ZL/soybean oil does not list Feb/Apr/Jun/Nov, so BOHO's contract-month choices are constrained to ZL's calendar), run concurrently with December, 1-3 books at a time, never all 8 simultaneously (position-limit/risk-seat constraint).

A front-month (continuously rolling) version was considered and explicitly rejected for this phase: it requires real contract-roll mechanics (roll-date determination, back-adjusted vs. raw price-series construction, ZL's calendar gaps relative to HO's, roll-day position handling) that add substantial engineering and statistical risk (non-stationary seasonality within a single spliced series, roll-date z-score contamination) relative to the fixed-month approach, which reuses an already-proven engine per month with no roll logic required.

Running the same engine across multiple months, however, exposed a maintainability problem in the current Dec workbook: contract specs, thresholds, sizing weights, stop-loss, and EWMA parameters are scattered as literal constants across specific absolute cell addresses inside the Position Model sheet, intermixed with thousands of live formulas; and the monthly Johansen refresh depends on a hardcoded "Active End Row" cell reference that must be manually advanced every month. Multiplying this by up to 3 concurrently-run books multiplies both the manual maintenance burden and the risk of one book silently drifting stale or a stray cross-sheet/cross-workbook reference contaminating one book's parameters with another's.

This spec defines a **Config-driven master template workbook** that captures Dec's proven engine logic once, parameterized so that activating a new month is a matter of supplying that month's price data and contract/threshold values — not re-deriving or hand-copying thousands of formulas.

## Goals

- Preserve Dec's exact, already-proven signal and PnL logic — this project parameterizes the engine, it does not redesign the strategy.
- Eliminate scattered literal constants: every contract spec, threshold, sizing weight, and risk parameter lives in one Config sheet per book, referenced via Excel defined names.
- Eliminate the manual monthly "Active End Row" bump via a self-updating last-row formula.
- Eliminate the manual per-year PnL column addition via a dynamic year-breakdown.
- Produce one canonical template workbook that every future live month is instantiated from, so a structural engine fix made once in the template is the single source of truth.
- Never touch or modify the live December workbook.

## Non-goals

- No front-month or contract-roll logic of any kind (no roll-date detection, no back-adjusted price series, no multi-contract handling within a single book).
- No "Active Trading Window" restriction — every book trades year-round, matching how December is actually run.
- No legacy/prototype sheet equivalent to Dec's hidden "Spread" sheet.
- No VBA/macros — pure formulas only, matching Dec's current style, `.xlsx` not `.xlsm`.
- No redesign of the Johansen/position-sizing/PnL calculation logic itself.

## Architecture

### Sheets (in order)

1. **Notes** — usage instructions, updated to describe the Config-driven workflow (replacing Dec's "paste data starting row 2" instructions with "fill in Config, then paste Data").
2. **Config** — new sheet; the single source of truth for every parameter (see below). All values are plain cells; each backs a workbook-scoped Excel defined name.
3. **Data** — Date, BOHO, D4, and the underlying SBO/HO settlement + open pricing block, structurally identical to Dec's Data sheet, but empty of values — populated per live book.
4. **DiffLag** — first differences and lagged levels, identical formula logic to Dec, filled to the template's row capacity.
5. **Matrices** — Johansen covariance matrices (S00/S11/S01/S10) over a rolling window, window length pulled from Config via defined name instead of Dec's hardcoded `63`.
6. **Johansen** — 2×2 eigenvalue/eigenvector solve, identical to Dec's formulas (no parameters to templatize here beyond what Matrices already pulls from Config).
7. **Johansen_Monthly** — month-end hedge-ratio log; the "Active End Row" pointer becomes a self-updating last-row lookup (`LOOKUP(2,1/(range<>""),ROW(range))` pattern) instead of Dec's hardcoded cell reference.
8. **Monthly Beta** — locked monthly hedge ratio carried forward, identical logic to Dec.
9. **Position Model** — the live engine: spread construction from the locked hedge ratio, rolling z-score, 3-tier entries/exits, per-leg dollar PnL, stop-loss, EWMA vol regime/shock detection, drawdown tracking, and the year-by-year PnL breakdown (made dynamic — see below). All literal constants replaced with defined-name references into Config.
10. **Results** — Johansen trace/max-eigen cointegration test summary, identical to Dec.

### Config sheet contents

| Group | Fields |
|---|---|
| Contract identity | Book label, ZL contract month, HO contract month |
| Contract specs | SBO lb/contract, ZL cents→$/gal multiplier, HO gal/contract, D4 multiplier |
| Johansen params | Rolling window length (days) — replaces Dec's hardcoded `63` |
| Signal thresholds | Entry z-thresholds (tiers 1-3), exit threshold, sizing weights (tiers 1-3, default 25/50/25%) |
| Risk params | Stop-loss ($), EWMA lambda, shock-multiplier threshold, risk-level scalar |

Each field is bound to a workbook-scoped defined name (e.g. `ZThreshold_Tier1`, `ContractMult_ZL`, `JohansenWindowLen`, `EWMA_Lambda`, `StopLoss_USD`). Every engine-sheet formula that currently reads a raw absolute cell address in Dec (e.g. `$DV$3`, `$DK$2`, `$DQ$7`) is rewritten to reference the corresponding name. This makes formulas self-documenting and prevents the specific failure mode identified during design: a stray cross-sheet/cross-workbook absolute reference silently pulling in another book's parameter, since a defined name must be explicitly present (or explicitly broken/flagged) in any workbook it's copied into rather than silently resolving to the wrong cell.

### Structural fixes carried into every sheet

1. **Dynamic Active End Row** (Johansen_Monthly): replace Dec's hardcoded `=F69`-style pointer with a `LOOKUP(2,1/(range<>""),ROW(range))` formula that always resolves to the sheet's own latest populated month-end row. No monthly manual edit required, and no risk of one book silently trading on a stale hedge ratio because the pointer wasn't advanced.
2. **Dynamic per-calendar-year PnL breakdown** (Position Model / Results): replace Dec's hardcoded per-year columns (2021, 2022, ... 2026, requiring a new column added every January) with a formula-driven breakdown that derives the set of years present from the Date column (e.g. a YEAR()-driven unique list feeding a SUMIFS-style per-year total), so a new calendar year requires zero manual edits.

### Row capacity

All engine sheets filled down to **3,000 rows** (~12 years of daily data at ~250 trading days/year), versus Dec's current ~1,500-row capacity — enough headroom that a live book won't need formulas manually extended for well over a decade.

## Data flow

Identical to Dec's proven chain: Data → DiffLag (differences/lags) → Matrices (rolling-window covariance matrices, window length from Config) → Johansen (eigen-solve) → Johansen_Monthly (month-end hedge-ratio log, self-updating last row) → Monthly Beta (prior-month-locked ratio, walk-forward) → Position Model (spread, z-score, tiered entries/exits sized and thresholded from Config, per-leg PnL using contract specs from Config, stop-loss/EWMA/drawdown, dynamic year breakdown) → Results (cointegration test summary).

No changes to the statistical or trading logic itself — only to where parameters live and how the month-end/year-end housekeeping resolves.

## Deployment workflow

To activate a new month's live book:
1. `Save As` the template workbook, named `Johansen_<Month>_BOHO_D4_Strategy.xlsx`.
2. Fill in Config: contract identity/specs for that month's ZL/HO contracts, and (once backtested/calibrated per the separate per-month calibration workstream) that month's thresholds and window length.
3. Paste that month's SBO/HO settlement + open price history into Data.
4. Engine sheets (DiffLag through Results) are never hand-edited — they should remain identical to the template across every live book. Periodic manual spot-checks (comparing engine-sheet formula text between a live book and the template) are the intended drift check; no automated tooling is being built for this now.

The live December workbook is not modified, copied from, or referenced by the template or any new book at any point in this project.

## Validation approach

Before any new live month is activated, the template itself must be validated against Dec's known-good output: copy Dec's actual historical Data into a throwaway instance of the template, set Config to match Dec's actual contract specs/thresholds/window length, and compare the resulting Position Model output (hedge ratios, z-scores, positions, PnL) row-for-row against the real Dec workbook's numbers. Any discrepancy indicates a formula-translation defect introduced during templatization, not a genuine strategy difference, and must be resolved before the template is used to build a real new-month book.

## Build approach

Script-generated (Python + openpyxl), not manual Excel editing. A script reads the actual formula patterns from the live Dec workbook, generates the template's sheets with those formulas filled down to 3,000 rows, substitutes defined-name references for the scattered constants per the Config mapping above, and applies the two structural fixes (dynamic end-row, dynamic year breakdown). This is reproducible — a defect found later is fixed in the script and the template regenerated, rather than hand-patched — and avoids manual-entry errors across the ~130-column, 3,000-row Position Model sheet.

## Open items for the implementation plan

- Exact Config sheet cell layout and the full defined-name list (one name per parameter currently scattered in Dec).
- Exact formula translation for the dynamic year-breakdown (YEAR()-driven unique list + aggregation vs. an alternative dynamic-array approach, depending on Excel version/functions already in use elsewhere in Dec, e.g. `FILTER` is already used in Johansen_Monthly).
- Order of sheet construction in the script and how Dec's formulas are extracted (read via openpyxl formula strings, transform cell-address tokens into name tokens, write to new sheets).
- Validation test plan detail: which specific Position Model columns get compared row-for-row against Dec, and acceptable tolerance (should be exact match, given identical inputs).
