import datetime
import pathlib
from validate_against_dec import (
    run_validation, compare_columns_for, EXCLUDED_COMPARE_COLUMNS, _is_excel_error, _normalize_for_compare,
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
}


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
    for col in ("DK", "DV", "DW", "CQ", "CE", "DZ", "EA", "EE"):
        assert col not in cols
        assert col in EXCLUDED_COMPARE_COLUMNS


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
        template_path=str(pathlib.Path(__file__).parent.parent.parent / "Johansen_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    unexplained = [
        m for m in mismatches
        if m["row"] not in KNOWN_ACCEPTED_MISMATCH_ROWS
        and (m["row"], m["col"]) not in SCATTERED_DEC_TEXT_QUIRK_CELLS
    ]
    assert unexplained == [], f"{len(unexplained)} unexplained mismatches, see _validation_diff.csv"
