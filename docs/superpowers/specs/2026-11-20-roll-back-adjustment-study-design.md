# Roll back-adjustment study (Phase 2) — design

Date: 2026-11-20
Status: draft, awaiting approval

## Intent

Phase 1 (merged 11/18) fixed the PnL on contract-roll days. The **signals**
still run on spliced BOHO prices that jump by the old→new contract gap on each
roll day. That jump reaches:

- the monthly Johansen hedge weights (63-row windows that contain the jump);
- the spread's 20-day Z-score;
- the EWMA volatility regime (CV = daily BOHO change, so the roll day's
  change is the whole gap).

This study measures how much removing the jump would change the signals,
trades and results, in all 8 months. Ryan decides afterwards, from the results,
whether to change the workbook. **This study changes no workbook Ryan uses.**

## Decisions already made (Ryan, 2026-11-20)

- Approach A (hybrid): Python back-adjusts the prices and recomputes the
  historical Johansen weights. A separate test copy of the Sandbox then gets
  the adjusted BOHO prices and the new weights, and Excel recalculates
  everything downstream.
- Back-adjust by **shifting** (adding/subtracting the gap), not by scaling.
- **D4 is not adjusted.** Vintage gaps are a few cents.
- No threshold for "worth changing the workbook" is set now. The results come first.

## Goals

- For each of the 8 months, a current-vs-back-adjusted comparison of:
  - hedge weights, especially in the ~3 months after each roll;
  - Z-score;
  - days with a different position;
  - regime labels;
  - trades added, removed or changed;
  - RA PnL, Sharpe and max drawdown, full history and 2024–26, per month and Portfolio.
- Results in a page Ryan can read (the same style as the earlier Profit-Take
  Test and Regime pages).
- A test workbook Ryan can open, whose new Dashboard shows the back-adjusted results.

## Non-goals

- No change to the Sandbox, the live file or `SBOxHO Historical Pricing - Roll Fix.xlsx`.
- No change to the generator scripts (`build_*.py`).
- No D4 adjustment and no proportional (ratio) adjustment.
- No new stop / profit-take tuning on adjusted data.

## Anti-goals

- Writing to, saving or holding open Ryan's Sandbox or live workbook. The
  study reads them only from a fresh copy.
- Back-adjusting the prices PnL is computed from. PnL must keep using the
  real SBO/HO settles (Position Model B/C ← Data settle block, plus the
  Roll Prices override from Phase 1).
- A comparison where differences could come from a Python bug rather than the
  adjustment. The Johansen re-implementation must first reproduce the stored
  weights from the unadjusted prices.

## Constraints and facts found in recon

- **The historical weights are stored numbers.** `Johansen_Monthly` keeps them
  per month-block as values (e.g. January block H3:I3 = 0.6258677961771613 /
  −0.7799291645453088 for 2021-01-29). Formulas in `Matrices_Johansen_Results`
  compute only the current window (End Row = `Johansen_Monthly!K3`, window =
  `Config!B11` = 63 rows, from DiffLag BOHO/D4 levels and differences: S00/S01/S11
  moment matrices, M = S11⁻¹ S10 S00⁻¹ S01, 2×2 eigenvector normalised; the sign rule is
  in rows 18–26). Position Model `I` (spread) looks the weight up by date
  (`LOOKUP(date−1, CT:CT, CU:CU/CV:CV)` in the month's Johansen_Monthly block).
- **Signal BOHO vs PnL prices.** In `Data`, each month has a first block
  (Date, BOHO, D4, e.g. A:C January … EY:FA December), then a Settle block
  (Date, SBO, HO, BOHO, D4) and an Open block. Position Model D (BOHO Settle,
  used by the spread and CV) reads the first block. B/C (SBO/HO, used for PnL)
  read the Settle block. DiffLag reads the first block. So adjusting
  **only the first-block BOHO column** of each month changes the signals and
  leaves PnL alone. The builder must confirm this by tracing every reference to
  those columns before relying on it.
- Roll Prices holds the old contract's SBO/HO settle (and open) on each roll
  day, per month. The roll gap is the new-contract BOHO minus the old-contract BOHO
  on the roll day. The old-contract BOHO is computed from the old SBO/HO with the same
  formula that produces Data's BOHO from SBO/HO. The builder confirms that formula
  reproduces Data's BOHO on ordinary days.
- Shift convention: the latest prices stay real. Every price before a roll is
  shifted by the sum of the gaps of all later rolls. Today's signal levels then
  stay comparable to the live file.
- Gotchas carried from earlier work:
  - read workbooks with `iter_rows(values_only=True)`, not random `ws.cell()`;
  - use `isolated_workbook` (DispatchEx) for every Excel run;
  - OneDrive AutoSave can persist edits to an open file;
  - FILTER spill formulas in Johansen_Monthly F:G must survive any write.

## Approaches considered

- **A. Hybrid (chosen).** Python only for prices and weights. Excel recalculates the
  rest in a test copy, so downstream logic (spread, Z, regime, tiers, stop,
  $100k profit-take + re-entry, PnL, Dashboard) is the real workbook's.
- **B. Everything in Python.** Fast reruns, but about 100 workbook columns
  re-implemented and validated by hand, and no workbook to open. Not chosen.

## Design

1. **Snapshot.** Copy the Sandbox to the scratchpad, and read inputs only from that
   copy. The Sandbox path is opened read-only, or not at all.
2. **Back-adjusted BOHO per month** (`scripts/phase2/back_adjust.py`): from
   Data's first-block BOHO, the dates and Roll Prices, compute each roll's gap and the
   shifted series. Output a table per month: date, original BOHO, gap applied,
   adjusted BOHO.
3. **Johansen re-implementation** (`scripts/phase2/johansen.py`): the
   Matrices_Johansen_Results calculation as a function of a 63-row window,
   applied at each stored month-end End Row. **Gate:** on the unadjusted series it
   reproduces every stored weight in all 8 blocks (tolerance about 1e-9). Then the same
   function runs on the adjusted series.
4. **Test workbook** (`scripts/phase2/build_test_workbook.py`):
   - Copy the Sandbox to a new file, `Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Back-Adjusted TEST.xlsx`, in the same folder.
   - Write the adjusted first-block BOHO values and the recomputed weight values into that copy only.
   - Fully recalculate in an isolated Excel, check for error cells, save.
   - The script refuses any target path that is the Sandbox, the live file or any existing file.
5. **Comparison** (`scripts/phase2/compare.py`): read both workbooks' values
   per month and write JSON tables for the results page:
   - weights at each month-end, and their change;
   - daily Z difference;
   - days the Total Position differs;
   - regime label differences;
   - trade list match-up (start date, side, exit reason, PnL);
   - per-year and total RA PnL, Sharpe and max DD, by month and Portfolio.
6. **Results page:** a published artifact (Dashboard type, as before) with a
   per-month view and Portfolio totals, and a plain-language summary at the top.

## Verification

- Back-adjusted series: ordinary-day changes are identical to the original,
  the roll-day change has the gap removed, and the latest value is unchanged.
- Johansen gate: the stored weights are reproduced (step 3).
- Test workbook:
  - its PnL inputs (Position Model B/C, Roll Prices) are identical to the Sandbox;
  - with **unadjusted** inputs written through the same pipeline (a dry run), every BT value equals the Sandbox's, which proves the write path changes nothing by itself;
  - zero new error cells.
- The Sandbox's file hash and modified time are unchanged after the study.

## Answers and calls after the builder check

- **Starting file (answered).** The Sandbox *is* the Phase 1 Roll Fix workbook
  (`Knox_Multimonth_BOHO_D4_Strategy - Roll Fix Sandbox.xlsx`). It already
  holds the Roll Prices PnL override, the $100k profit-take + re-entry and
  the redesigned Dashboard, so the test copy starts from it.
- **What else reads Johansen_Monthly (answered by trace, 11/20).** The Position Model tabs
  read only each block's date column and two weight columns (e.g. January
  G/H/I … December CT/CU/CV). Nothing reads a trace statistic, an eigenvalue or a
  cointegration flag from it, so only the weights need recomputing.
- **Roll Prices coverage (answered by trace).** Every month has one row per roll from
  2021 to its 2026 roll (rolls not yet reached are blank). Claude's call: a roll
  inside the data with no Roll Prices row stops the run with a report; no gap is
  estimated.
- **If the Johansen gate fails (Claude's call).** The comparison baseline is the
  dry-run copy: unadjusted prices plus Python-recomputed weights, written through the
  same pipeline. If the gate passes, the dry run equals the Sandbox exactly. If it
  fails, both sides still use the same weight calculation (like for like), and the
  gap between the stored and recomputed weights is reported as its own finding
  (for example, stored weights computed on older data). The run does not stop.

## Left to the builder

The Data column map per month, the BOHO-from-SBO/HO formula, the handling of
months whose first roll falls before the data starts, chart and table
choices on the results page, and helper file layout.
