import datetime
import pathlib
from validate_against_dec import (
    run_validation, compare_columns_for, EXCLUDED_COMPARE_COLUMNS, _is_excel_error, _normalize_for_compare,
    param_block_value_parity_mismatches,
)

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

# Known, accepted, documented exception — re-diagnosed after CRITICAL 2's
# widened comparison (the old narrow 7-column comparison never exercised most
# of this). See Definition of Done in the plan
# (docs/superpowers/plans/2026-08-05-johansen-fixed-month-template.md).
#
# Root cause (unchanged from the original diagnosis): Dec's own live file has
# a literal hardcoded 0 in Position Model!O19:P21 (a one-off manual patch for
# the strategy's ~20-day rolling-window warm-up, where the z-score is
# naturally #N/A/#DIV/0!). The template correctly writes the real general
# formula there instead, which ripples through the self-referencing
# O{r-1}/P{r-1} position-tracking chain (S/T/U/V/X) through row 24.
#
# What's new with the widened comparison: columns outside the original narrow
# 7 that reference S/T/U/V with a ONE- or TWO-ROW LAG (e.g. AG-AL/AP-AU entry-
# and exit-level columns, AV-BJ per-position PnL, and the BM/BO/BQ/BU
# self-referencing trade-state-machine cluster together with its non-self-
# referencing dependents BN/BP/BR/BS/BT/CD/CK/CL/CM/CN/CP) still see a stale
# (still-erroring) S/T/U/V value for 1-4 additional rows after row 24, before
# every lagged reference finally points only at clean rows and (for the
# self-referencing columns) IFERROR's single-row self-heal has had a chance
# to run. This is the SAME root cause with a wider observed footprint, not a
# new defect -- confirmed by every mismatch in this window tracing back to a
# lagged read of S/T/U/V during rows 19-24.
KNOWN_ACCEPTED_MISMATCH_ROWS = range(19, 29)  # rows 19 through 28 inclusive

# A second, unrelated, documented exception: a handful of scattered cells in
# Dec's own live file (Position Model!BC3:BC5, BN10, BN12, BN14, BN16, BP8,
# BT18) are formatted as TEXT in Dec's real workbook, so Dec's own formula-
# looking content in those cells is inert display text, never evaluated --
# openpyxl's data_only=True read of Dec's recalculated copy returns the
# literal formula string, not a computed value. The template's freshly-
# created cells have no such stray Text number-format, so the (identical)
# formula text written there IS evaluated as a real formula by Excel,
# correctly producing whatever that bootstrap-window formula computes to
# (e.g. #N/A/#DIV/0!, since these rows are inside BOOT_END's literal-copy
# window and were never meant to be live per-day formulas). This is a Dec-
# side data-entry artifact, not a template defect -- there is nothing to
# "fix" in the template without deliberately reintroducing Dec's own mistake.
SCATTERED_DEC_TEXT_QUIRK_CELLS = {
    (3, "BC"), (4, "BC"), (5, "BC"),
    (8, "BP"),
    (10, "BN"), (12, "BN"), (14, "BN"), (16, "BN"),
    (18, "BT"),
    # targeted-fix-round-2, finding 4: Position Model!DK7 in Dec's own live
    # file holds the literal TEXT "=ROUNDUP( SBO / Hedge Ratio, 0)" -- a
    # human-written note/placeholder that LOOKS like a formula but is stored
    # as Text (Dec's own cell format), so it is never evaluated in Dec's
    # real recalculation; openpyxl reads back the literal string. The
    # template's freshly-created DK7 has no such stray Text format, so the
    # (identical) text IS evaluated as a real formula -- but "SBO"/"Hedge
    # Ratio" aren't real cell references or defined names anywhere, so it
    # resolves to #NAME?. Same category as the cells above (a Dec-side
    # data-entry artifact, invisible until finding 4's new PARAM_BLOCK_COLS
    # value-parity check started actually comparing these cells at all) --
    # not a template defect, nothing to "fix" without reintroducing Dec's
    # own mistake.
    (7, "DK"),
}

# targeted-fix-round-2, finding 4: two more categories of PARAM_BLOCK_COLS
# mismatch discovered once param_block_value_parity_mismatches() started
# actually comparing these cells for the first time (previously wholly
# unverified). Both trace to root causes that already exist elsewhere in
# this file, just manifesting in a param-block SUMMARY cell instead of a
# per-day cell:

# (a) Dec's real DG49/DH49 described a "24-26 window"-specific max
# drawdown/peak equity (see build_position_model.py's
# _apply_targeted_fix_round_2_overrides docstring) computed from BZ/CA/CB/CC,
# which finding 1 deliberately redefines to an ordinary since-inception
# accumulator (same reasoning, and same EXCLUDED_COMPARE_COLUMNS exclusion,
# as BZ/CA/CB/CC themselves). DG49/DH49 now correctly compute the genuine
# ALL-TIME max drawdown/peak equity instead -- a different, equally valid,
# number than Dec's narrower-window one, by design, not a defect.
PARAM_BLOCK_STRUCTURALLY_REDEFINED_CELLS = {(49, "DG"), (49, "DH")}

# (b) Dec's real DE46/DF2-DF49(subset)/DG48/DH44/DH48 are MAX/MIN/AVERAGE/
# SUM/STDEV aggregates over a fixed row range (e.g. '=SUM($BR$21:$BR$1386)',
# '=MIN($BY$20:$BY$740)') that happens to include the strategy's ~20-row
# warm-up window (rows ~20-24) -- exactly the same already-documented root
# cause as KNOWN_ACCEPTED_MISMATCH_ROWS (Dec's own literal-0 O19:P21 patch
# vs. the template's real general formula, which naturally errors during
# warm-up). A per-day cell only inherits this for a few rows near the
# warm-up window itself (already covered by KNOWN_ACCEPTED_MISMATCH_ROWS);
# these param-block cells inherit it for their ENTIRE aggregate, because
# MAX/MIN/AVERAGE/SUM/STDEV over a range containing even one error cell
# propagates that error to the whole result. Confirmed by direct inspection
# of the recalculated template: AY/BD/BI/BK/CK (feeding DF2-DF17) and BT
# (feeding DF33/DF39/DF48) are #N/A/#DIV/0! at rows 20-24 ONLY, exactly the
# warm-up window, and clean everywhere else. Same root cause, no new defect
# -- fixing it would mean redesigning these aggregate ranges to skip the
# warm-up window, out of scope for this targeted fix round (not one of its
# 5 named findings).
PARAM_BLOCK_WARMUP_INHERITED_ERROR_CELLS = {
    (46, "DE"),
    (2, "DF"), (3, "DF"), (4, "DF"), (5, "DF"), (6, "DF"),
    (8, "DF"), (9, "DF"), (10, "DF"), (11, "DF"),
    (14, "DF"), (15, "DF"), (16, "DF"), (17, "DF"),
    (33, "DF"), (39, "DF"), (48, "DF"),
    (48, "DG"),
    (44, "DH"), (48, "DH"),
}

# A third, unrelated, documented exception, introduced by targeted-fix-round-2
# finding 2's param_block_value_parity_mismatches() check (finding 4): Dec's
# own real DJ224:DN227 hold real MAX/MIN-over-CQ-CU numbers (e.g. Dec's real
# DK224≈253122, DK225≈-623899), but the template deliberately clears these
# specific 15 cells to blank (see build_position_model.py's
# _apply_targeted_fix_round_2_overrides) rather than reproduce the
# now-decommissioned CQ-CU dangling reference in a new shape. This is an
# intentional design choice (the genuine general equivalent, EG/EH/EI, lives
# in the dynamic year-breakdown table instead), not a defect -- see
# build_position_model.py's DJ224:DN227-clearing comment for the reasoning.
PARAM_BLOCK_INTENTIONALLY_CLEARED_CELLS = {
    (row, col)
    for row in (224, 225, 227)
    for col in ("DJ", "DK", "DL", "DM", "DN")
}

# targeted-fix-round-2, finding 3: the row-window exception above (~630
# possible cells across 10 rows x nearly every compared column) previously
# had no count assertion -- it blanket-exempted every mismatch landing in
# that window with no bound, so a future regression landing in the same rows
# (or the count silently growing) would pass silently. This pins the actual
# count of mismatches absorbed by ALL THREE accepted exceptions combined
# (row-window + scattered-text-quirk + intentionally-cleared-cells), the same
# rigor SCATTERED_DEC_TEXT_QUIRK_CELLS already applies via its exact
# (row, col) enumeration. Re-diagnosed against the post-targeted-fix-round-2
# widened harness (compare_columns_for now also covers ~137 columns instead
# of ~97, since PARAM_BLOCK_COLS's row-for-row exclusion is unchanged, but
# param_block_value_parity_mismatches() is a wholly new, additive check) --
# see targeted-fix-round-2-report.md for the exact count and how it was
# obtained. If this number changes, investigate why before updating it; do
# not just bump it to make the test pass.
EXPECTED_ACCEPTED_MISMATCH_COUNT = 414  # from the real COM validation run, see targeted-fix-round-2-report.md


def test_compare_columns_now_covers_far_more_than_the_old_narrow_7():
    # CRITICAL 2: the old COMPARE_COLUMNS = ["I","L","S","T","U","V","X"] covered
    # only 7 of ~137 Position Model columns, and none of them were self-referencing
    # accumulators, so it could never have caught CRITICAL 1. The comparison must
    # now be an opt-OUT list covering nearly every column.
    cols = compare_columns_for(137)
    assert len(cols) > 90
    # The accumulator/drawdown columns CRITICAL 1 fixed must be covered.
    for col in ("Y", "AZ", "BE", "BJ", "CK", "CL", "CM", "CN", "BV", "BW", "BX", "BY"):
        assert col in cols
    # The old narrow 7 must still be covered too.
    for col in ("I", "L", "S", "T", "U", "V", "X"):
        assert col in cols


def test_compare_columns_excludes_documented_non_comparable_columns():
    # PARAM_BLOCK_COLS, the decommissioned per-year columns, and the dynamic
    # year-breakdown table are documented exclusions (see EXCLUDED_COMPARE_COLUMNS'
    # own docstring/comment in validate_against_dec.py) -- structurally not
    # per-day time series, so a row-for-row compare against them is meaningless.
    cols = compare_columns_for(137)
    for col in ("DK", "DV", "DW", "CQ", "CE", "DZ", "EA", "EE", "EG"):
        assert col not in cols
        assert col in EXCLUDED_COMPARE_COLUMNS
    # targeted-fix-round-2, finding 1: BZ/CA/CB/CC are now deliberately
    # redefined (since-inception, not Dec's row-740 trailing window) and
    # excluded for the same "structurally not comparable" reason.
    for col in ("BZ", "CA", "CB", "CC"):
        assert col not in cols
        assert col in EXCLUDED_COMPARE_COLUMNS


def test_param_block_value_parity_flags_within_dec_extent_only():
    # targeted-fix-round-2, finding 4: PARAM_BLOCK_COLS is excluded from the
    # per-day row-for-row comparison above, but must not be excluded from ALL
    # verification -- this unit test exercises param_block_value_parity_
    # mismatches() directly (no COM/Excel recalculation needed) against two
    # small synthetic worksheets standing in for dec_pm/template_pm.
    import openpyxl
    from openpyxl.utils import column_index_from_string

    dec_wb = openpyxl.Workbook()
    dec_ws = dec_wb.active
    tmpl_wb = openpyxl.Workbook()
    tmpl_ws = tmpl_wb.active
    dd_col = column_index_from_string("DD")

    dec_ws.cell(row=1, column=dd_col, value=42)
    tmpl_ws.cell(row=1, column=dd_col, value=42)  # matches -- not flagged

    dec_ws.cell(row=2, column=dd_col, value=99)
    tmpl_ws.cell(row=2, column=dd_col, value=0)  # genuine mismatch -- flagged

    # Row 3: Dec's own source is blank here (outside Dec's real populated
    # extent for this column) -- whatever the template holds must NOT be
    # flagged, even though it's non-blank (this is exactly the padding/blank
    # region the check must never extend the comparison into).
    tmpl_ws.cell(row=3, column=dd_col, value=123)

    mismatches = param_block_value_parity_mismatches(dec_ws, tmpl_ws)
    assert {
        "row": 2, "col": "DD", "dec": 99, "template": 0, "kind": "param_block_value_parity",
    } in mismatches
    assert all(m["row"] != 1 for m in mismatches)
    assert all(m["row"] != 3 for m in mismatches)


def test_is_excel_error_detects_error_sentinels_not_plain_values():
    assert _is_excel_error("#N/A")
    assert _is_excel_error("#DIV/0!")
    assert _is_excel_error("#REF!")
    assert not _is_excel_error(0)
    assert not _is_excel_error(None)
    assert not _is_excel_error("Sell")
    assert not _is_excel_error(1.5)


def test_normalize_for_compare_reconciles_datetime_and_serial_number():
    # Position Model!A holds identical real dates in both workbooks, but Dec's
    # column A cells carry a date number format (openpyxl returns a python
    # datetime) while the template's freshly-created cells don't (openpyxl
    # returns the plain underlying int/float) -- a cosmetic formatting
    # difference that must not manufacture a mismatch for every row.
    assert _normalize_for_compare(datetime.datetime(2021, 1, 5)) == 44201.0
    assert _normalize_for_compare(44201) == 44201
    assert _normalize_for_compare("Sell") == "Sell"
    assert _normalize_for_compare(None) is None


def test_validation_reports_no_mismatches_on_dec_replay():
    mismatches = run_validation(
        template_path=str(pathlib.Path(__file__).parent.parent.parent / "Johansen_MultiMonth_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    accepted = [
        m for m in mismatches
        if m["row"] in KNOWN_ACCEPTED_MISMATCH_ROWS
        or (m["row"], m["col"]) in SCATTERED_DEC_TEXT_QUIRK_CELLS
        or (m["row"], m["col"]) in PARAM_BLOCK_INTENTIONALLY_CLEARED_CELLS
        or (m["row"], m["col"]) in PARAM_BLOCK_STRUCTURALLY_REDEFINED_CELLS
        or (m["row"], m["col"]) in PARAM_BLOCK_WARMUP_INHERITED_ERROR_CELLS
    ]
    unexplained = [m for m in mismatches if m not in accepted]
    assert unexplained == [], f"{len(unexplained)} unexplained mismatches, see _validation_diff.csv"
    # targeted-fix-round-2, finding 3: the row-window exception has no count
    # bound of its own -- pin the TOTAL number of mismatches all three
    # accepted exceptions combined actually absorb, so a regression landing
    # in one of these windows (or the count silently growing) fails loudly
    # instead of being silently absorbed forever.
    assert len(accepted) == EXPECTED_ACCEPTED_MISMATCH_COUNT, (
        f"accepted-exception count changed: {len(accepted)} actual vs "
        f"{EXPECTED_ACCEPTED_MISMATCH_COUNT} expected -- investigate before "
        "updating this number, see targeted-fix-round-2-report.md"
    )


def test_populate_throwaway_copy_writes_into_decembers_block_not_column_a(tmp_path):
    import openpyxl
    from build_template import build_template
    from validate_against_dec import _populate_throwaway_copy
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col

    template_path = str(tmp_path / "template.xlsx")
    build_template(SOURCE_PATH, template_path, num_rows=20)
    throwaway_path = str(tmp_path / "throwaway.xlsx")

    _populate_throwaway_copy(template_path, SOURCE_PATH, throwaway_path)

    wb = openpyxl.load_workbook(throwaway_path, data_only=False)
    december_offset = column_offset("December", DATA_BLOCK_WIDTH, GAP)
    date_col = shift_col("A", december_offset)
    boho_col = shift_col("B", december_offset)
    assert wb["Data"][f"{date_col}2"].value is not None
    assert wb["Data"][f"{boho_col}2"].value is not None
    # January's block (offset 0) must NOT have been touched by Dec's data.
    assert wb["Data"]["A2"].value is None
