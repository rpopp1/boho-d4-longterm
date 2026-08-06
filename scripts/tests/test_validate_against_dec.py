import pathlib
from validate_against_dec import run_validation, compare_columns_for, EXCLUDED_COMPARE_COLUMNS, _is_excel_error

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

# Known, accepted, documented exception — see Definition of Done in the plan
# (docs/superpowers/plans/2026-08-05-johansen-fixed-month-template.md, commit
# b86fa03). Dec's own live file has a literal hardcoded 0 in Position Model!
# O19:P21 (a one-off manual patch for the strategy's ~20-day rolling-window
# warm-up, where the z-score is naturally #N/A/#DIV/0!). The template
# correctly writes the real general formula there instead, which ripples
# through the self-referencing O{r-1}/P{r-1} position-tracking chain (columns
# S/T/U/V, and X via the tier totals) until it resynchronizes with Dec at row
# 25. Confirmed by direct diagnosis (Task 15, Addendum 3): every row from 25
# through the end of Dec's real data (~1386) matches Dec exactly across all 7
# compared columns; this is specific to Dec's January-2021 startup dates and
# is not something a future month's book would need to replicate.
KNOWN_ACCEPTED_MISMATCH_ROWS = range(19, 25)  # rows 19, 20, 21, 22, 23, 24


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


def test_validation_reports_no_mismatches_on_dec_replay():
    mismatches = run_validation(
        template_path=str(pathlib.Path(__file__).parent.parent.parent / "Johansen_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    unexplained = [m for m in mismatches if m["row"] not in KNOWN_ACCEPTED_MISMATCH_ROWS]
    assert unexplained == [], f"{len(unexplained)} unexplained mismatches, see _validation_diff.csv"
