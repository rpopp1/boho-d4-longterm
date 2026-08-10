# Johansen Multi-Month BOHO/D4 Consolidation — Design

## Context

The [Johansen Fixed-Month Template Engine](2026-08-05-johansen-fixed-month-template-design.md) produces a blank, Config-driven template (`Johansen_Template_Engine.xlsx`) that gets `Save As`-ed into a standalone workbook per contract month (documented in `docs/ACTIVATING_A_NEW_MONTH.md`). Today, `Johansen_FM_BOHO_D4_Strategy.xlsx` (December, the original live book) and `Johansen_January_BOHO_D4_Strategy.xlsx` exist this way.

Running multiple contract months this way means opening and updating a separate workbook per active month every day — pasting the same-shaped price data into N different files, and maintaining N different `Johansen_Monthly` manual log entries. As more of the 8 tradeable BOHO contract months (constrained to ZL/soybean oil's listed months) come online, this manual multiplication becomes the operational bottleneck.

This spec defines a **single consolidated workbook** covering all 8 contract months at once: one place to paste price data, one place to see all 8 months' cointegration diagnostics, and 8 `Position Model` tabs to flip between — replacing the per-month Save-As workflow entirely.

## Goals

- Replace the "Save-As a new workbook per contract month" process with one workbook covering all 8 months.
- One `Data` tab to paste price history into, instead of N separate files.
- 8 independent `Position Model` tabs (one per contract month), each fully self-contained and flippable, each computing off its own month's data and its own month's calibrated parameters.
- Every shared engine stage (`DiffLag`, the cointegration test, `Johansen_Monthly`, `Monthly Beta`) consolidated onto one tab per stage, with 8 side-by-side per-month column-blocks, instead of one tab per stage per month.
- Every `Config` parameter — contract identity, contract specs, and every calibration parameter (thresholds, sizing weights, stop-loss, EWMA lambda, Johansen window length) — independently settable per contract month, so each month can be calibrated separately to its own PnL/Sharpe optimum.
- Preserve today's proven signal and PnL logic exactly — this is a re-architecture of the container, not a change to the strategy.
- Migrate December's and January's existing real historical data into the new workbook so no live history is lost in the transition.

## Non-goals

- **No fix for the known hardcoded-row-boundary defect** in `Position Model`'s summary/Sharpe panel (punch list item 1 in `docs/superpowers/HANDOFF.md`). Each of the 8 `Position Model` tabs will initially carry the same interim limitation the single-month template has today. This gets addressed as a follow-up once the 8-tab structure is confirmed working end-to-end.
- No fix for punch list items 2-4 (the `#N/A` summary cells, header label inconsistencies, validation harness blind spots) — same reasoning, deferred.
- No change to the underlying Johansen/z-score/position-sizing/PnL statistical logic itself.
- No front-month or contract-roll logic (already excluded by the original design spec; still excluded here).
- No automated data feed or API import. "Uploading a price series" means pasting values into the `Data` tab's per-month block — the same manual-paste workflow as today, just consolidated into one tab instead of N files.
- No change to the live December source workbook (`Johansen_FM_BOHO_D4_Strategy.xlsx`) — it remains read-only, untouched, per the existing hard constraint.

## Architecture

### Tabs (in order)

| Tab | Structure |
|---|---|
| `Notes` | Single tab; usage instructions rewritten for the multi-month layout |
| `Config` | 1 tab, 8 columns (one per contract month) — every field independently settable |
| `Data` | 1 tab, 8 column-blocks — each block structurally identical to today's Data sheet (Date/BOHO/D4 + Settlement pricing block + Open pricing block), each with its own independent row range/date axis |
| `DiffLag` | 1 tab, 8 column-blocks |
| `Matrices_Johansen_Results` | 1 tab, 8 column-blocks — each block combines today's `Matrices` + `Johansen` + `Results` content for that month (all three are small, fixed-size dashboards; `DiffLag`/`Monthly Beta` are the large day-by-day chains and get their own tabs) |
| `Johansen_Monthly` | 1 tab, 8 column-blocks — each still needs the recurring manual month-end log entry appended, per block, per month |
| `Monthly Beta` | 1 tab, 8 column-blocks |
| `Position Model — <Month>` ×8 | 8 separate tabs, one per contract month: January, March, May, July, August, September, October, December |

Total: ~15 physical tabs, vs. today's 10-tab template × up to 8 files.

### Per-month block pattern

Every consolidated tab (`Data`, `DiffLag`, `Matrices_Johansen_Results`, `Johansen_Monthly`, `Monthly Beta`) follows the same layout convention:

- Each month gets a fixed-width column-block, in month order (January → December), separated by blank gap columns for readability (exact widths finalized in the implementation plan).
- A header row above each block names the month (e.g. "January", "March", ...), so it's visually obvious where one month's data ends and the next begins.
- Each block's internal column layout is identical to that sheet's current single-month layout — only the starting column offset changes per month.
- Row numbering is independent per block: each month's price history starts at its own row 2, with its own length. Months do not share a calendar; a month with a shorter history simply has fewer populated rows in its block.
- Exact column-letter ranges per block are an implementation-plan-level detail, not fixed here.

`Position Model` does not follow this pattern — each month gets a full standalone tab, laid out exactly as today's single-month `Position Model` sheet.

### Config redesign

Because every `Config` field is now independently calibrated per month, workbook-scoped Excel defined names (e.g. `ZThreshold_Tier1`) go away entirely — a single name can't resolve to 8 different cells. Instead:

- `Config` becomes one tab with one row per parameter (contract identity, contract specs, Johansen window length, thresholds, sizing weights, stop-loss, EWMA lambda, shock multiplier, risk level — the full list in `scripts/config_map.py`'s `CONFIG_LAYOUT`) and one column per contract month.
- Every formula that currently reads a defined name instead reads its own month's `Config` column directly (e.g. May's block reads `Config!D14` for its Tier 1 threshold, where today's single-month template would read the name `ZThreshold_Tier1`).
- All 19 `CONFIG_FIELDS` entries — including the true physical contract-spec constants (`HOGalPerContract`, `ZLMultiplier`, `ZLCentsPerLB`, `D4Multiplier`) that are unlikely to actually differ by month — still get one column per month, for structural uniformity. This keeps the substitution mechanism uniform (every field is "read this month's column," none are "read the one shared name") rather than building a hybrid of per-month-addressed and workbook-named fields.
- This is a parameterization of the existing substitution mechanism, not a new concept: `scripts/config_map.py`'s `CONSTANT_MAP` (currently "Dec address → defined name") becomes "Dec address → (parameter, month) → that month's Config column/row," applied once per month-block instead of once globally.

### Row capacity

Each month's block/tab keeps the existing 3,000-row capacity (~12 years of daily data), independent of the other 7 months. This means the consolidated workbook has roughly 8× the formulas of today's single-month template — a real increase in file size and recalculation time that should be watched, but not a blocker.

## Build pipeline changes

- `scripts/config_map.py`: `CONFIG_LAYOUT` and `CONSTANT_MAP` restructured around `(parameter, month) → cell` instead of `parameter → defined name`.
- Each `build_*.py` (`build_config_sheet.py`, `build_static_sheets.py`, `build_diff_lag.py`, `build_matrices_johansen.py`, `build_johansen_monthly.py`, `build_monthly_beta.py`, `build_position_model.py`) gains a loop over the 8 contract months, writing one column-shifted block per month instead of one block at column A. `build_position_model.py`'s loop writes 8 separate sheets instead of 8 blocks.
- `formula_tools.py` and `sheet_cloner.py` (the row-shifting formula-generation core) need to additionally shift formulas *across* by a per-month column offset, not just down by row — the existing "pattern row, shift down N rows" mechanism gets a matching "pattern column-block, shift right by M columns" dimension.
- `build_template.py` becomes the orchestrator for the month loop, in addition to its current sheet-ordering role.
- A new one-time migration step copies December's and January's real `Data` + `Johansen_Monthly` log values into their respective blocks in the new workbook. The other 6 months' blocks start empty.

## Validation approach

`scripts/validate_against_dec.py` continues to work exactly as it does today — copy Dec's real historical `Data` + `Johansen_Monthly` log into a throwaway copy of the new consolidated template, force two real Excel recalculations via COM, and diff the December `Position Model` tab's output against the real Dec workbook's numbers row-for-row. December remains the only month with a known-good reference to validate against; the other 7 months' blocks are structurally identical (same generated formula pattern) but can't be independently validated until they have their own real trading history.

## Deployment / cutover

1. Regenerate the blank multi-month engine template from the build pipeline (working name: `Johansen_MultiMonth_Template_Engine.xlsx`).
2. Validate December's block against the real live December workbook (see above).
3. One-time `Save As` into the live, ongoing multi-month book (working name TBD) and run the Dec/Jan data migration step.
4. Continue trading December out of the existing live workbook (`Johansen_FM_BOHO_D4_Strategy.xlsx`) until the new workbook's December block is validated and cut over; only then retire the old per-month files (`Johansen_January_BOHO_D4_Strategy.xlsx`) and the old Save-As-per-month workflow.
5. `docs/ACTIVATING_A_NEW_MONTH.md` gets rewritten for the new workflow: activating a previously-empty month becomes "fill in that month's `Config` column, paste that month's price history into that month's `Data` block" — no more `Save As`, no more new file.

## Open items for the implementation plan

- Exact column-letter layout (block width, gap columns, starting offsets) for each consolidated tab.
- Exact tab-naming convention for the 8 `Position Model` sheets and header-row labeling for each month-block.
- Working file names for the regenerated blank template and the live multi-month book.
- Whether/how `formula_tools.py`'s row-shifting regex engine needs hardening to safely handle a second (column) shift dimension without breaking existing row-shift behavior.
- Exact scope of the Dec/Jan migration script (which cells, how validated post-migration).
- Whether existing tests in `scripts/tests/` need per-month parameterization or a representative subset (running the full suite × 8 months could meaningfully lengthen the ~10-15 minute Excel-COM-driven run today).
