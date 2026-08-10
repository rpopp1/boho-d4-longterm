# Handoff: Johansen Multi-Month Consolidated Template Engine

**Last updated:** 2026-08-10
**Status:** Core build pipeline working and verified. Multi-month architecture consolidates all 8 contract months into one workbook. Known limitations documented below still exist independently in each of the 8 `Position Model - <Month>` tabs and must be addressed before using any month's data for real live trading.

## What this project is

A Python/openpyxl build pipeline that reads the live, actively-traded December BOHO/D4 Johansen cointegration strategy workbook (`Johansen_FM_BOHO_D4_Strategy.xlsx`) and generates a parameterized multi-month template (`Johansen_MultiMonth_Template_Engine.xlsx`) where all 8 contract months (January, March, May, July, August, September, October, December) live in one workbook, side-by-side. Each month has its own `Config` column (B-I), its own `Data` block, and its own `Position Model - <Month>` tab, eliminating the need to manage 8 separate files while keeping months' calculations independent and auditable. Originated from the need to consolidate a single-month template (`Johansen_Template_Engine.xlsx`) into a multi-month pipeline; see conversation history for full context.

**Read first, in this order:**
1. `docs/superpowers/specs/2026-08-06-johansen-multi-month-consolidation-design.md` — the multi-month consolidation design spec (goals, architecture, month-block layout, column-offset logic)
2. `docs/superpowers/plans/2026-08-06-johansen-multi-month-consolidation.md` — the multi-month consolidation implementation plan. **This is the most important document** — every task section has inline addenda documenting what was found and why, and the **"Definition of done"** + **"Known limitations"** sections at the very end are the current state of the project (not aspirational). This consolidation built on top of a prior fixed-month template project; earlier design/plan docs are still available if you need deeper context.
3. This handoff doc, for a fast-start summary and pointers.

## Quick orientation for resuming work

```bash
cd "C:\Users\Ryan Popp\OneDrive - Karbone Inc\Desktop\Renewable Fuels Analysis\RINs\BOHO Trading"
git log --oneline | head -30          # full history, one commit per task
cd scripts && py -m pytest tests/ -v  # full suite; takes ~40-50 min. Only test_validate_against_dec.py uses real Excel COM automation; the rest of the runtime is openpyxl building 8-month workbooks (every builder now runs 8x, once per month, instead of once)
```

- **This is a git repo** (initialized specifically for this project). `.xlsx`/`.xls`/`.xlsm` files are gitignored and must never be committed — the generated `Johansen_MultiMonth_Template_Engine.xlsx` and the live source workbook both live outside git tracking.
- **`scripts/DEC_SOURCE_PATH.txt`** holds the absolute path to the live source workbook. Every script reads it read-only. **Never write to that file** — verify its mtime/sha256 unchanged after any session touching this code: mtime `1785880251.4742322`, sha256 `09eef5f8bfc4fd81f4b94e8822b7264cbcb7d1352d1db01f03ab31bb08b6f9d8` (as of 2026-08-06; if these don't match, something touched the live file — stop and investigate before doing anything else).
- A git worktree (`johansen-multi-month-consolidation`) was used for this consolidation project, isolated from the original fixed-month work. All commits are on this worktree's branch.
- The SDD (subagent-driven-development) execution ledger and every task's detailed report live in `.superpowers/sdd/2026-08-06-johansen-multi-month-consolidation/` — **this directory is gitignored**, so it only exists on this machine, not in a fresh clone. It has the full blow-by-blow history if you need more detail than the plan doc's summaries.

## What's solid (verified via real Excel COM validation)

- The full sheet-builder pipeline (`scripts/build_*.py` → `scripts/build_template.py`) produces a 15-sheet multi-month template (Notes, Config, Data, DiffLag, Matrices_Johansen_Results, Johansen_Monthly, Monthly Beta, plus 8 separate `Position Model - <Month>` tabs) that is reproducible from the live December source.
- Each month's blocks (Config column, Data block, Position Model tab, etc.) are correctly offset and isolated from each other using the month-index-based offset logic in `scripts/multi_month.py`.
- Both hard constraints held throughout: live source workbook never written to; no Excel file ever committed to git.
- The prior fixed-month project's Critical defects (accumulator `#N/A` propagation, literal-constant leaks) are fixed in the multi-month template and confirmed via real Excel COM recalculation (not just openpyxl-level unit tests — `scripts/validate_against_dec.py` drives real Excel via `pywin32`/COM to force recalculation and read back live values).
- 70+ tests pass (`scripts/tests/`), including end-to-end replay of December's real historical data through the December block of the multi-month template, diffed against December's own real recalculated output, with an explicit count-bounded set of documented exceptions.
- A one-time migration script (`scripts/migrate_existing_months.py`) exists to carry December's and January's existing real trading history and log entries into their respective blocks in `Johansen_BOHO_D4_Strategy_AllMonths.xlsx`. Its `migrate_month_data` function has dedicated unit test coverage (`scripts/tests/test_migrate_existing_months.py`), and `_read_config_values` now has coverage too. `run_migration` (the end-to-end entry point that actually writes `Johansen_BOHO_D4_Strategy_AllMonths.xlsx` from the real source files) has **not yet been run against the real live files** — running it is a deliberate manual cutover step the user takes when ready, per the design spec's "Deployment/cutover" section, not something already done as part of this project.

## Known limitations — the punch list for next session

Full detail is in the plan doc's "Known limitations" section (search for that heading). Summary: **All four issues below now apply independently to each of the 8 `Position Model - <Month>` tabs** (one per contract month, January through December), not hidden across 8 separate files. **None were fixed by this consolidation** — they're just relocated and now visible as separate instances rather than hidden in 8 different single-month workbooks.

Roughly in priority order:

1. **Most consequential: hardcoded row boundaries in each tab's param-block cells.** Each `Position Model - <Month>` tab still has ~30 cells (the headline "RA 2021 Sharpe"–"RA 2026 Sharpe" cells in rows 33-38 and others) that hardcode that tab's specific historical row boundaries. They compute correctly for that month's replay (so the validation harness can't catch them — right output, wrong reason) but will silently produce wrong numbers if a new month is added or row boundaries shift. **Fix approach:** generalize into the dynamic per-year mechanism already used in other parts of each tab, or clear the hardcoded cells. Add a build-time test asserting no param-block formula contains a bare row-number literal above ~100.
2. **~20 cells in each tab's headline summary panel show `#N/A`** (fixable with `AGGREGATE(...,6,...)` (ignore-errors mode) or an `IFERROR`-guarded range start, the same technique already used for accumulator columns). Do this before shipping any month's data to a real trader.
3. **Minor: header/label inconsistency in each tab** — some cells still reference 24-26 row ranges after being redefined to all-time accumulators. Rename headers and resolve duplicated/conflicting row labels in each tab.
4. **Minor: validation harness blind spots** — accepted-exception count is one aggregate total across all categories, can't detect offsetting drift between months. Would need per-month exception sub-tracking to catch regressions per tab.



## Key files (all in `scripts/`)

| File | Purpose |
|---|---|
| `multi_month.py` | Core multi-month constants: `MONTHS` list (January, March, May, July, August, September, October, December), `column_offset(month, block_width, gap)` for computing each month's block position, `GAP`, `DATA_BLOCK_WIDTH`, `DIFFLAG_BLOCK_WIDTH`, etc. — the source of truth for all month→column-offset calculations |
| `formula_tools.py` | `extract_row_formulas`, `shift_formula_rows` (the core "fill down" regex engine), `rewrite_constants` — unchanged from fixed-month project, still required for the multi-month build |
| `config_map.py` | `CONFIG_FIELDS` (per-month Config layout: row numbers for each field), `CONFIG_MONTH_COLUMNS` (month→column mapping: January=B, March=C, ... December=I), `constant_map_for_month(month)` / `matrices_only_constant_map_for_month(month)` — now returns month-specific Config cell references instead of workbook-scoped defined names |
| `sheet_cloner.py` | `fill_sheet_from_pattern` — the generic per-row filler for building blocks. **Raises `NotImplementedError` on any `ArrayFormula`/dynamic-array spill cell** — openpyxl cannot reliably reproduce Excel's dynamic-array metadata. Any future sheet with a spill formula must be redesigned around this. |
| `build_config_sheet.py` | Multi-month version: builds the shared Config sheet with 8 columns (B-I), one per month, each with the standard field set from `config_map.CONFIG_FIELDS` |
| `build_static_sheets.py` | Builds the shared Notes and Data sheets; `write_notes_sheet()` creates Notes, `write_data_sheet(wb, month)` builds that month's Data block (20 columns wide with 2-column gap between blocks, using `multi_month.column_offset()` for positioning) |
| `build_diff_lag.py` | Multi-month version: builds the shared DiffLag sheet with 8 blocks (one per month), each 8 columns wide, using `multi_month.column_offset()` for positioning |
| `build_matrices_johansen.py` | Multi-month version: builds the combined Matrices_Johansen_Results sheet with 8 mega-blocks (one per month), each containing three sub-sections (Matrices 8 cols, Johansen 6 cols, Results 2 cols) using `multi_month.mjr_offsets()` for positioning |
| `build_johansen_monthly.py` | Multi-month version: builds the shared Johansen_Monthly sheet with 8 blocks (one per month), each 11 columns wide, with a manually-maintained 4-column log entry (EndRow, Date, BOHO_Weight, D4_Weight) per month |
| `build_monthly_beta.py` | Multi-month version: builds the shared Monthly Beta sheet with 8 blocks (one per month), each 8 columns wide, using `multi_month.column_offset()` for positioning |
| `build_position_model.py` | Multi-month version: called once per month (in a loop), generates one `Position Model - <Month>` tab per contract month (January through December), 8 separate tabs total. Each tab still contains the IFERROR-wrapped accumulators, dynamic year-breakdown, and param-block cells (the hardcoded-row limitations still apply independently in each tab) |
| `build_template.py` | Orchestrator; run directly (`py build_template.py`) to regenerate `Johansen_MultiMonth_Template_Engine.xlsx` with all sheets and blocks in place |
| `validate_against_dec.py` | End-to-end validator: copies Dec's real Data + Johansen_Monthly logs into the December block of a throwaway multi-month template copy, forces two real Excel recalculations via COM, diffs December's block output against December's own real output. This is the ultimate proof of correctness — trust it over any unit test. |
| `migrate_existing_months.py` | One-time migration script that reads the old single-month December and January workbooks and carries their real trading history and Johansen_Monthly log entries into the corresponding blocks of `Johansen_BOHO_D4_Strategy_AllMonths.xlsx` — built and unit-tested (`migrate_month_data`, `_read_config_values`) but **not yet run** against the real live files; the user runs it once, manually, when ready to cut over. Not needed for future month activations after that. |

## Hard-won gotchas (don't rediscover these)

- **Dynamic arrays (`FILTER`, etc.) cannot be safely reproduced via openpyxl.** A single-cell `ArrayFormula.ref` makes Excel treat it as a legacy CSE array (returns only the first value). A multi-cell `ref` fixes that but Excel still needs `cm="1"` metadata (which openpyxl can't write) to treat it as a true auto-resizing spill — without it, any oversized declared range gets padded with `#N/A`. **The only reliable fix found: redesign away from spill formulas entirely** (plain per-row mirror formulas, or manually-maintained log columns).
- **Self-referencing accumulator chains need `IFERROR(prev,0)` wraps**, or a single warm-up `#N/A` propagates forever. Don't hardcode recovery values for specific months — wrap the formula instead, so it works for every month independently.
- **A literal-copy loop with `if value is None: continue` silently lets earlier-pass garbage survive** wherever the source is blank. The multi-month builders are loop-based (one pass per month) — verify each builder that does value-copying checks for this pattern and clears cells explicitly when copying nothing.
- **Month-block offsets compound: if a month-index function is wrong, every month after that month lands in the wrong columns.** The authoritative offset calculations are in `scripts/multi_month.py` (`column_offset`, `mjr_offsets`). Don't hardcode month-to-column mappings anywhere else — always call these functions. Test at least January (index 0, offset 0) and December (index 7) to catch off-by-one errors.
- **A narrow validation comparison set gives false confidence.** The multi-month consolidation validation (in `validate_against_dec.py`) still only checks December's block (because December is the only month with a live source workbook to diff against). Future validations of new months will need to be added as those months' real trading history is populated.
- **"Accepted exceptions" need an exact count assertion, not a widened range/ceiling** — otherwise a future regression silently disappears into the tolerance.

## If you're picking this back up

1. Re-read the plan's "Known limitations" section fully before touching code. Remember: all four known issues now apply independently to each of the 8 `Position Model - <Month>` tabs.
2. Decide whether to fix all 4 items in one pass or dispatch them individually. Item 1 (hardcoded row boundaries) is the most consequential and touches the most cells in each tab; items 2-4 are more contained.
3. **Critical:** Check that your fix is generalizable across all 8 `Position Model` tabs, not just December. The `build_position_model.py` builder is called in a loop (once per month); any fix must work for all months independently.
4. Whatever you fix, regenerate the template with `py build_template.py` and re-run `validate_against_dec.py` as the final check — it's caught every real defect in this project so far; trust it over anything else.
5. If you add a genuinely new sheet or column class, check it against the "hard-won gotchas" above before writing any code. In particular, verify that any month-to-column offset calculations go through `scripts/multi_month.py`, not hardcoded anywhere.
