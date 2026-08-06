import openpyxl, pathlib, tempfile, os
from build_template import build_template

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_build_template_creates_expected_sheets_in_order():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb.sheetnames == [
            "Notes", "Config", "Data", "DiffLag", "Matrices",
            "Johansen", "Johansen_Monthly", "Monthly Beta",
            "Position Model", "Results",
        ]

def test_build_template_data_sheet_has_no_price_values():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb["Data"]["B2"].value is None
        assert wb["Data"]["C2"].value is None

def test_build_template_no_leftover_dec_absolute_constants():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        ws = wb["Position Model"]
        for row in ws.iter_rows(min_row=2, max_row=11):
            for cell in row:
                if isinstance(cell.value, str):
                    assert "$DV$6" not in cell.value
                    assert "$DQ$7" not in cell.value
