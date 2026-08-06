import pathlib
from validate_against_dec import run_validation

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_validation_reports_no_mismatches_on_dec_replay():
    mismatches = run_validation(
        template_path=str(pathlib.Path(__file__).parent.parent.parent / "Johansen_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    assert mismatches == [], f"{len(mismatches)} mismatched rows, see _validation_diff.csv"
