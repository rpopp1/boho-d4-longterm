# Handoff: Johansen Fixed-Month Template Engine

**Last updated:** 2026-08-06
**Status:** Core build pipeline working and verified. Known limitations documented below must be addressed before this template is used to build a real live-trading book for a new month.

## What this project is

A Python/openpyxl build pipeline that reads the live, actively-traded December BOHO/D4 Johansen cointegration strategy workbook (`Johansen_FM_BOHO_D4_Strategy.xlsx`) and generates a parameterized template (`Johansen_Template_Engine.xlsx`) so the same trading engine can be instantiated for a different fixed contract month (e.g. March, July) without hand-copying ~130 columns of formulas. Originated from a conversation comparing front-month vs. fixed-month strategy design — see conversation history for that reasoning; this doc only covers the build itself.

**Read first, in this order:**
1. `docs/superpowers/specs/2026-08-05-johansen-fixed-month-template-design.md` — the original design spec (goals, non-goals, architecture)
2. `docs/superpowers/plans/2026-08-05-johansen-fixed-month-template.md` — the implementation plan, which grew substantially as real defects were discovered mid-build. **This is the most important document** — every task section has inline "Addendum"/"v2 redesign" notes documenting what was wrong and why, and the **"Definition of done"** + **"Known limitations"** sections at the very end are the current, honest state of the project (not aspirational).
3. This handoff doc, for a fast-start summary and pointers.

## Quick orientation for resuming work

```bash
cd "C:\Users\Ryan Popp\OneDrive - Karbone Inc\Desktop\Renewable Fuels Analysis\RINs\BOHO Trading"
git log --oneline | head -30          # full history, one commit per fix/finding
cd scripts && py -m pytest tests/ -v  # full suite; takes ~10-15 min (real Excel COM automation)
```

- **This is a git repo** (initialized specifically for this project). `.xlsx`/`.xls`/`.xlsm` files are gitignored and must never be committed — the generated `Johansen_Template_Engine.xlsx` and the live source workbook both live outside git tracking.
- **`scripts/DEC_SOURCE_PATH.txt`** holds the absolute path to the live source workbook. Every script reads it read-only. **Never write to that file** — verify its mtime/sha256 unchanged after any session touching this code: mtime `1785880251.4742322`, sha256 `09eef5f8bfc4fd81f4b94e8822b7264cbcb7d1352d1db01f03ab31bb08b6f9d8` (as of 2026-08-06; if these don't match, something touched the live file — stop and investigate before doing anything else).
- **No git worktree/branch was used** — all work is directly on `master`, by explicit choice early in the project (fresh repo, nothing to protect at the time).
- The SDD (subagent-driven-development) execution ledger and every task's detailed report live in `.superpowers/sdd/2026-08-05-johansen-fixed-month-template/` — **this directory is gitignored**, so it only exists on this machine, not in a fresh clone. It has the full blow-by-blow history (what each fix round found, exact commands run, formula inspections) if you need more detail than the plan doc's summaries. Key files: `progress.md` (ledger), `task-*-report.md` (per-task), `final-review-fix-wave-report.md`, `targeted-fix-round-2-report.md`.

## What's solid (verified independently across 3 rounds of adversarial review)

- The full sheet-builder pipeline (`scripts/build_*.py` → `scripts/build_template.py`) produces a 10-sheet template (Notes, Config, Data, DiffLag, Matrices, Johansen, Johansen_Monthly, Monthly Beta, Position Model, Results) that is reproducible byte-for-byte from the live source.
- Both hard constraints held throughout: live source workbook never written to; no Excel file ever committed to git.
- The 3 original Critical defects found during build (accumulator `#N/A` propagation, an overly-narrow validation harness, a ~52,000-cell literal-constant leak) are genuinely fixed and confirmed via real Excel COM recalculation (not just openpyxl-level unit tests — openpyxl never calculates formulas, so `scripts/validate_against_dec.py` drives real Excel via `pywin32`/COM to force recalculation and read back live values).
- 73 tests pass (`scripts/tests/`), including an end-to-end replay of Dec's real historical data/parameters through the generated template, diffed against Dec's own real recalculated output, with an explicit count-bounded (`== 414`) set of documented, root-caused exceptions.

## Known limitations — the punch list for next session

Full detail is in the plan doc's "Known limitations" section (search for that heading). Summary, roughly in priority order:

1. **A live instance of the exact defect this project exists to eliminate.** `Position Model!DF33:DF38` ("RA 2021 Sharpe" ... "RA 2026 Sharpe") and ~30 other param-block cells (`DE46`/`DF48`/`DF49`/`DG44/45/48`/`DH44/45/48`, `DF2-5`/`DF8-17`/`DF20-22`/`DF27-29`/`DT3-5`, `DO10`) still hardcode Dec-specific row boundaries (250, 495, 740, 990, 1240, 1262, 1257, 1227, and a literal `C1266`). These compute correctly for Dec-replay (so the validation harness can't catch them — right output, wrong reason) but will silently produce wrong numbers for any other month. **Fix approach:** generalize into the dynamic `EE`/`EF`/`EG`/`EH`/`EI` per-year mechanism already built (see `_write_dynamic_year_breakdown` and the `EE`/`EF` SUMIFS pattern in `scripts/build_position_model.py`), or clear them the way `DJ224:DN227` was cleared. Add a build-time test asserting no param-block formula contains a bare row-number literal above ~100.
2. **20 cells in the headline summary panel (`DE46`, `DF2-6/8-11/14-17/33/39/48`, `DG48`, `DH44/48`) are `#N/A` and were incorrectly documented as "unfixable."** They're actually fixable — wrap the aggregates with `AGGREGATE(...,6,...)` (ignore-errors mode) or start the range after the warm-up window, the same `IFERROR` technique already used for the accumulator columns. Do this before shipping to a real trader.
3. **Header/label inconsistency**: `BZ1:CC1` still say "24-26" after being redefined to since-inception accumulators (matching `BV:BY`) in the last fix round. Rename headers, and resolve `Position Model` row 49 (labelled "24-26" but now duplicates the all-time number) / row 48 (labelled "21-23", currently `#N/A`).
4. Minor: the last fix round introduced 8 new error cells (`BZ:CC` rows 17-18) that are simultaneously excluded from the comparison that would catch them. The accepted-exception count assertion (`== 414`) is one aggregate total, not per-category, so it can't detect offsetting drift. `param_block_value_parity_mismatches()` only checks one direction (can't detect a future leak into cells where Dec is blank).

## Key files (all in `scripts/`)

| File | Purpose |
|---|---|
| `formula_tools.py` | `extract_row_formulas`, `shift_formula_rows` (the core "fill down" regex engine — has had several rounds of hardening, see plan Task 3), `rewrite_constants` |
| `config_map.py` | `CONFIG_LAYOUT` (Config sheet defined names), `CONSTANT_MAP` (generic per-day rewrite map), `MATRICES_ONLY_CONSTANT_MAP` (scoped to Matrices sheet only — do not merge into the generic map, see plan Task 4/7) |
| `sheet_cloner.py` | `fill_sheet_from_pattern` — the generic per-row filler. **Raises `NotImplementedError` on any `ArrayFormula`/dynamic-array spill cell** — openpyxl cannot reliably reproduce Excel's dynamic-array metadata (`cm="1"`), confirmed the hard way (see plan Task 7 Addendum v2). Any future sheet with a spill formula must be redesigned around this (plain per-row formula or manually-maintained column), not patched into this function. |
| `build_position_model.py` | The largest, most complex builder — `ACCUM_SELF_REF_COLS` (IFERROR-wrapped accumulators), `BOOT_END`/bootstrap-phase classification, `PARAM_BLOCK_COLS` (static param block, must be fully cleared beyond Dec's real extent — this is where the ~52,000-cell leak happened), dynamic year-breakdown (`EE`/`EF`/`EG`/`EH`/`EI`) |
| `build_johansen_monthly.py` | Redesigned (v2) to eliminate FILTER spill anchors entirely — column A is a plain `=DiffLag!B{row}` mirror; F/G/H/I are all manually-maintained log columns (the user must append a row monthly — this is documented in the generated Notes sheet) |
| `build_template.py` | Orchestrator; run directly (`py build_template.py`) to regenerate `Johansen_Template_Engine.xlsx` |
| `validate_against_dec.py` | End-to-end validator: copies Dec's real Data + Johansen_Monthly F:I log into a throwaway copy of the template, forces two real Excel recalculations via COM, diffs output. This is the ultimate proof of correctness — trust it over any unit test. |

## Hard-won gotchas (don't rediscover these)

- **Dynamic arrays (`FILTER`, etc.) cannot be safely reproduced via openpyxl.** A single-cell `ArrayFormula.ref` makes Excel treat it as a legacy CSE array (returns only the first value). A multi-cell `ref` fixes that but Excel still needs `cm="1"` metadata (which openpyxl can't write) to treat it as a true auto-resizing spill — without it, any oversized declared range gets padded with `#N/A`. **The only reliable fix found: redesign away from spill formulas entirely** (plain per-row mirror formulas, or manually-maintained log columns).
- **Self-referencing accumulator chains need `IFERROR(prev,0)` wraps**, or a single warm-up `#N/A` propagates forever. Dec's live file "solves" this with a literal hardcoded `0` at specific historical rows — do NOT replicate that literal (it's Dec-specific and won't exist for a new month); wrap the formula instead.
- **A literal-copy loop with `if value is None: continue` silently lets earlier-pass garbage survive** wherever the source is blank. This exact bug shape caused the ~52,000-cell leak (fixed for 5 columns, missed for ~30 others, found by a later review) and needs to be checked for anywhere else a similar "copy from source, else leave whatever was there before" pattern exists.
- **A narrow validation comparison set gives false confidence.** The original harness compared 7 non-accumulating columns; the real defects were all in the ~127 it didn't check. When validating, prefer an opt-out list of exclusions (with each exclusion justified) over an opt-in list of what to check.
- **"Accepted exceptions" need an exact count assertion, not a widened range/ceiling** — otherwise a future regression silently disappears into the tolerance. This project's exception list grew from 23 → 376 → 414 over three rounds; each growth was verified legitimate (traced to new checks catching things for the first time, not to widening old exclusions) specifically because each was checked against an exact-count assertion, not a loose bound.

## If you're picking this back up

1. Re-read the plan's "Known limitations" section fully before touching code.
2. Decide whether to fix all 3 items in one pass or dispatch them individually — item 1 (hardcoded per-year formulas) is the most consequential and touches the most cells; items 2-3 are more contained.
3. Whatever you fix, regenerate the template and re-run `validate_against_dec.py` as the final check — it's caught every real defect in this project so far; trust it over anything else.
4. If you add a genuinely new sheet or column class, check it against the "hard-won gotchas" above before writing any code.
