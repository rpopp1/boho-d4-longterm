# Roll Back-Adjustment Study (Phase 2) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Measure how much removing the contract-roll jump from the BOHO signal prices changes hedge weights, signals, trades and results in all 8 Knox months, without touching Ryan's workbooks.

**Architecture:**
- Python reads a scratch copy of the Sandbox and builds back-adjusted (shifted) first-block BOHO prices per month.
- It recomputes the monthly Johansen weights with a re-implementation gated against the stored weights.
- It writes adjusted prices and weights into a brand-new TEST copy, and Excel recalculates everything downstream.
- A comparison script reads the baseline and the TEST workbook and writes JSON for a results page.

**Tech Stack:** Python 3, openpyxl (read, `iter_rows(values_only=True)`), numpy, pywin32 via `excel_com.isolated_workbook`, pytest.

**Spec:** `docs/superpowers/specs/2026-11-20-roll-back-adjustment-study-design.md`

## Global Constraints

- The Sandbox `Knox_Multimonth_BOHO_D4_Strategy - Roll Fix Sandbox.xlsx` and the live file are never opened for writing. All reads use a copy in the scratchpad.
- New output workbooks only:
  - `Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Back-Adjusted TEST.xlsx` (adjusted)
  - `Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Baseline DRY RUN.xlsx` (dry run)

  Both go in the BOHO Trading folder. The writer refuses any existing path.
- Adjust by **shifting**, so the latest price stays real. Only the first-block BOHO column of each month changes (Data `B`, `X`, … per month, `column_offset(month, 20)`). **D4 is not adjusted.** The Settle/Open blocks and SBO/HO are not touched.
- A roll inside the data with no Roll Prices row, or a roll date missing from Data's dates, stops the run with a message naming the month and the date.
- If the Johansen gate fails, the run continues: the baseline becomes the dry-run workbook (unadjusted prices + Python weights), and the stored-vs-recomputed gap is reported.
- Window = `Config!B11` = 63 rows. Start row = `max(3, end − 63 + 1)`. Moments are divided by T = end − start + 1, and are not demeaned. M = S11⁻¹·S10·S00⁻¹·S01. Use the eigenvector of the larger eigenvalue, v = (M12, λ − M11), normalised; if v₂ > 0, negate both (`Matrices_Johansen_Results` rows 7–26, N18:N19).
- BOHO from legs: `BOHO = SBO × 7.5 / 100 − HO` (SBO in cents/lb, Config row 8 = 7.5). Task 1 confirms this on ordinary days.
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. Branch: `phase2-back-adjust-study`.

## Review Focus

1. **A month whose first roll falls before 2021-01-04, or on the first data row:** no adjustment for it, and no index error. *(test in Task 1)*
2. **A Roll Prices date that is not among Data's dates (holiday or typo):** the run stops naming the month and date. *(test in Task 1)*
3. **An output path that already exists, or is the Sandbox/live file:** refused before any copy. *(test in Task 3)*
4. **Excel recalculation leaving new error cells,** for example a broken FILTER spill in Johansen_Monthly F:G after the H:I writes: the build fails, listing each sheet whose error count exceeds the source copy's. *(test in Task 3)*
5. **A trade whose start moves by a day or two:** counted as "shifted", not as one removed plus one added. *(test in Task 4)*

---

### Task 1: Inputs, BOHO-from-legs and back-adjusted series

**Files:**
- Create: `scripts/phase2/__init__.py` (empty), `scripts/phase2/inputs.py`, `scripts/phase2/back_adjust.py`
- Test: `scripts/tests/test_phase2_back_adjust.py`

**Interfaces:**
- Produces in `inputs.py`:
  - `@dataclass MonthInputs`, with fields:
    - `month: str`
    - `rows: list[int]` (sheet rows of real dates, starting at 2)
    - `dates: list[date]`
    - `boho: list[float]`, `d4: list[float]` (first block)
    - `sbo: list[float]`, `ho: list[float]` (settle block)
    - `rolls: list[tuple[date, float, float]]` (roll date, old SBO, old HO)
    - `stored: list[tuple[int, date, float, float]]` (Johansen_Monthly F, G, H, I: end row, month-end date, BOHO weight, D4 weight)
  - `load_inputs(path: Path) -> dict[str, MonthInputs]`, which reads `Data`, `Roll Prices` and `Johansen_Monthly` with `iter_rows(values_only=True)`.
  - `data_boho_col(month) -> str` (first-block BOHO letter).
  - `johansen_cols(month) -> dict` (`end`, `date`, `w_boho`, `w_d4` letters; January F/G/H/I, then +13 per month).
- Produces in `back_adjust.py`:
  - `boho_from_legs(sbo: float, ho: float) -> float`
  - `roll_gaps(m: MonthInputs) -> list[tuple[date, float]]`: gap = Data BOHO on the roll date − `boho_from_legs(old SBO, old HO)`.
  - `back_adjust(dates: list[date], boho: list[float], gaps: list[tuple[date, float]]) -> list[float]`: each price on a date before a roll gets that roll's gap added, summed over all later rolls.

- [ ] **Step 1: Write the failing tests:**
  - `test_back_adjust_shifts_history_and_keeps_latest`: dates d1..d5 with prices [1.0, 1.1, 2.0, 2.1, 2.2] and gap (d3, 0.85). Expect [1.85, 1.95, 2.0, 2.1, 2.2]. The latest value is unchanged, and the d2→d3 change = 0.05.
  - `test_two_rolls_accumulate`: gaps (d2, 0.5) and (d4, −0.2). Value at d1 = p1 + 0.3, at d3 = p3 − 0.2.
  - `test_roll_on_first_row_or_before_data_is_ignored` *(Review Focus 1)*: a gap at d1, or at a date before d1, leaves the series unchanged.
  - `test_roll_date_missing_from_data_raises` *(Review Focus 2)*: `roll_gaps` on a MonthInputs whose roll date isn't in `dates` raises `ValueError`, and the message contains the month and the ISO date.
  - `test_boho_from_legs_matches_data_on_the_copy` (integration, using `PHASE2_COPY` = scratch copy path from env; skip if unset): for every month, ≥99% of ordinary days satisfy |boho_from_legs(sbo, ho) − boho| < 1e-9. Print the first mismatching dates if not.
  - `test_load_inputs_layout` (integration): January `stored[0]` == (20, date(2021,1,29), 0.6258677961771613, −0.7799291645453088); `data_boho_col("March") == "X"`; `johansen_cols("March")["w_boho"] == "U"`.
- [ ] **Step 2: Run** `python -m pytest scripts/tests/test_phase2_back_adjust.py -q`. Expected: FAIL (module missing).
- [ ] **Step 3: Implement** `inputs.py` and `back_adjust.py`, using `multi_month.column_offset` / `roll_prices_col` for the column maps.
- [ ] **Step 4: Run** the tests with `PHASE2_COPY` set. Expected: PASS. If the BOHO formula test fails, record what the formula really is and fix the constant. Don't loosen the test.
- [ ] **Step 5: Commit** `Add Phase 2 inputs reader and back-adjusted BOHO series`

---

### Task 2: Johansen re-implementation and the gate

**Files:**
- Create: `scripts/phase2/johansen.py`
- Test: `scripts/tests/test_phase2_johansen.py`

**Interfaces:**
- Consumes: `MonthInputs`, `load_inputs` (Task 1).
- Produces:
  - `johansen_weight(boho_by_row: dict[int, float], d4_by_row: dict[int, float], end_row: int, window: int = 63) -> tuple[float, float]`
  - `monthly_weights(m: MonthInputs, boho: list[float] | None = None) -> list[tuple[int, date, float, float]]`: same end rows and dates as `m.stored`; `boho` overrides `m.boho`.
  - `gate_report(m: MonthInputs) -> dict` with `{"max_abs_diff": float, "worst": (date, stored, recomputed)}`.
  - DiffLag rows: dBOHO[r] = boho[r] − boho[r−1], lag[r] = boho[r−1]; the same for D4. Rows ≥ 3.

- [ ] **Step 1: Write the failing tests:**
  - `test_weight_is_unit_length_and_sign_rule`: on a synthetic cointegrated pair (seeded numpy, 80 rows), the result has norm 1 within 1e-12 and a D4 component ≤ 0.
  - `test_window_start_clamps_at_row_3`: end_row 20 uses rows 3..20. Changing row 2's prices doesn't change the result; changing row 3's does.
  - `test_gate_reproduces_stored_weights` (integration, `PHASE2_COPY`): for each month, print `gate_report`. Assert max_abs_diff < 1e-9 for every month. If this fails, the plan's fallback applies (see Global Constraints). Mark the test `xfail(strict=False)` only after recording the failure and its likely cause in the ledger as a ruling.
- [ ] **Step 2: Run** them. Expected: FAIL (module missing).
- [ ] **Step 3: Implement** with numpy 2×2 algebra, mirroring the sheet's formulas exactly. Use the closed-form eigenvalues λ = (tr ± √(tr² − 4·det))/2, not `numpy.linalg.eig`, so the results match bit for bit.
- [ ] **Step 4: Run** them. Expected: PASS, or a gate failure handled as above.
- [ ] **Step 5: Commit** `Add Phase 2 Johansen weight re-implementation and stored-weight gate`

---

### Task 3: Test-workbook builder (dry run and adjusted)

**Files:**
- Create: `scripts/phase2/build_test_workbook.py`
- Test: `scripts/tests/test_phase2_build.py`

**Interfaces:**
- Consumes: Tasks 1–2.
- Produces:
  - `check_output_path(out: Path, protected: list[Path]) -> None`: raises `ValueError` if `out` exists or resolves to a protected path.
  - `new_error_sheets(base: dict[str, int], test: dict[str, int]) -> list[str]`: sheets whose test count is above the base count.
  - `build(source_copy: Path, out: Path, mode: str) -> dict`, where `mode` is `"dry"` (original BOHO + Python weights) or `"adjusted"` (adjusted BOHO + Python weights on adjusted prices). Steps:
    1. `shutil.copy2(source_copy, out)`.
    2. In `isolated_workbook(out, manual_calc=True)`, write each month's BOHO values into the first-block column (rows from `MonthInputs.rows`, as one `Range.Value` per month) and the weights into Johansen_Monthly H/I at the stored rows.
    3. `CalculateFullRebuild`, then count error cells per sheet with `error_cell_count`.
    4. Compare with the source copy's counts (computed once, the same way, in its own isolated session without saving). Raise if `new_error_sheets` isn't empty.
    5. Save, then `reset_calc_mode_to_automatic`.
    6. Write `<out>.manifest.json` with mode, gaps applied per month, and the gate report.
  - CLI: `python scripts/phase2/build_test_workbook.py --source <copy> --mode dry|adjusted`. The output name comes from the mode (Global Constraints). The protected list = the Sandbox + every `*.xlsx` in the folder whose name lacks "Phase 2".
- [ ] **Step 1: Write the failing tests:**
  - `test_refuses_existing_or_protected_paths` *(Review Focus 3)*: a tmp file that exists → ValueError; a new path equal to a protected path → ValueError; a new unrelated path → no error.
  - `test_new_error_sheets` *(Review Focus 4)*: `({"A":5,"B":0}, {"A":5,"B":2,"C":1})` → `["B","C"]`.
- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** the tests. Expected: PASS. Then run the **dry** build on the scratch copy and check, with `iter_rows` on both files:
  - every Position Model B, C and Roll Prices value is identical to the source copy;
  - if the gate passed, every month's BT column is identical to the source copy (max abs diff 0).

  Record the result in the ledger.
- [ ] **Step 5: Commit** `Add Phase 2 test-workbook builder with dry-run and adjusted modes`

---

### Task 4: Comparison

**Files:**
- Create: `scripts/phase2/compare.py`
- Test: `scripts/tests/test_phase2_compare.py`

**Interfaces:**
- Consumes: the two workbooks from Task 3 (the baseline is the source copy if the gate passed, else the dry-run workbook), plus `load_inputs`.
- Produces:
  - `trades(dates, bs, v, bu, en) -> list[dict]` with `start`, `end`, `side` (sign of V at start), `pnl` (BU on the last active day) and `reason` ("profit-take" if EN = 1 at the end, "stop" if it ended with the signal still on, else "signal"). A trade is a run of BS = 1.
  - `match_trades(base: list[dict], test: list[dict], shift_days: int = 3) -> dict` with lists `same`, `changed_pnl`, `shifted`, `added`, `removed`. "Shifted" = the same side starting within `shift_days` trading days.
  - `compare(base_path: Path, test_path: Path, out_dir: Path) -> None`, which writes JSON:
    - `weights.json`: month, month-end, base/test weights, diff, months since roll;
    - `daily_diff.json`: per month, the count of days with position differences, regime-label differences, and Z diff mean/max overall and within 21 trading days after each roll;
    - `trades.json`: the match-up per month;
    - `summary.json`: RA PnL, Sharpe, Sortino, max DD and trade count, per month and Portfolio, for 2021–26 and 2024–26, base vs test. Same formulas as `dashboard_metrics` (population SD, √252) applied to BT.
- [ ] **Step 1: Write the failing tests:**
  - `test_trades_split_on_bs_runs`: BS [0,1,1,0,1,1,1,0] → 2 trades with the right start/end. Reasons: EN = 1 on the first trade's end → "profit-take"; V still non-zero at the second's end → "stop".
  - `test_match_trades_counts_shifted_not_added_removed` *(Review Focus 5)*: base starts [d10 long, d40 short], test [d11 long, d40 short with another pnl] → shifted 1, changed_pnl 1, added 0, removed 0.
  - `test_summary_matches_dashboard_on_baseline` (integration): the baseline's Portfolio 2021–26 total and Sharpe equal 26,050,468.60 and 1.9107606698748356 within 1e-6.
- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** them. Expected: PASS.
- [ ] **Step 5: Commit** `Add Phase 2 baseline-vs-adjusted comparison`

---

### Task 5: Run the study, verify, publish results

**Files:**
- Create (scratchpad): the run log and JSON outputs.
- No repo code changes unless verification finds a bug (fix in the owning task's module, with a test).

- [ ] **Step 1:** Record the Sandbox's SHA-256 and modified time. Copy it to the scratchpad (`p2_source.xlsx`) and set `PHASE2_COPY`.
- [ ] **Step 2:** Run the dry build (Task 3 Step 4 checks) and then the adjusted build. Check, on the adjusted workbook:
  - ordinary-day BOHO changes equal the original's;
  - roll-day changes have the gap removed;
  - each month's last BOHO value is unchanged;
  - PnL inputs (Position Model B/C, Roll Prices) are identical;
  - no new error cells.
- [ ] **Step 3:** Run `compare` and read `summary.json` and `trades.json`. Spot-check two roll days by hand (old/new BOHO, gap, Z before and after).
- [ ] **Step 4:** Publish the results page: Artifact `quickstart` with intent "other". Use the Dashboard type it names, as for the earlier pages. Content:
  - a plain-language summary;
  - Portfolio and per-month base-vs-adjusted table (PnL, Sharpe, max DD, trades, 2021–26 and 2024–26);
  - a per-month chart of weight changes around rolls;
  - a trade match-up table;
  - regime and position difference counts;
  - a method note (shift, D4 untouched, gate result).
- [ ] **Step 5:** Confirm the Sandbox's SHA-256 and modified time are unchanged. Commit the code and update the handoff doc with a Phase 2 results section and Next steps.
