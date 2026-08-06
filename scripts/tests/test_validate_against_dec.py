import pathlib
from validate_against_dec import run_validation

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


def test_validation_reports_no_mismatches_on_dec_replay():
    mismatches = run_validation(
        template_path=str(pathlib.Path(__file__).parent.parent.parent / "Johansen_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    unexplained = [m for m in mismatches if m["row"] not in KNOWN_ACCEPTED_MISMATCH_ROWS]
    assert unexplained == [], f"{len(unexplained)} unexplained mismatches, see _validation_diff.csv"
