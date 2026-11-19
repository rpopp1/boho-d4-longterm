import openpyxl, pathlib, tempfile, os
from build_template import build_template
from multi_month import MONTHS

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_build_template_creates_expected_sheets_in_order():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb.sheetnames == [
            "Notes", "Config", "Data", "DiffLag", "Matrices_Johansen_Results",
            "Johansen_Monthly", "Monthly Beta",
        ] + [f"Position Model - {m}" for m in MONTHS] + ["Roll Prices"]


def test_build_template_data_sheet_has_no_price_values_for_every_month():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
        for month in MONTHS:
            offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
            assert wb["Data"][f"{shift_col('B', offset)}2"].value is None
            assert wb["Data"][f"{shift_col('C', offset)}2"].value is None


def test_build_template_has_no_array_formula_cells_anywhere():
    # MINOR 6: 5 ArrayFormula cells (Matrices!H7, Position Model!DC22/DE43/
    # DF43/DF46) slipped through the plain iter_rows-based literal-copy loops
    # in build_matrices_johansen.py and build_position_model.py, which had no
    # NotImplementedError guard for dynamic-array/spill formulas the way
    # sheet_cloner.fill_sheet_from_pattern does. All were verified to be
    # single-cell (legacy CSE) array formulas, safe to write as plain formula
    # strings. Guard against any regression across the whole built template.
    from openpyxl.worksheet.formula import ArrayFormula

    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=50)
        wb = openpyxl.load_workbook(out, data_only=False)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    assert not isinstance(cell.value, ArrayFormula), (
                        f"{sheet_name}!{cell.coordinate} is an ArrayFormula"
                    )


def test_build_template_no_leftover_dec_absolute_constants_in_any_month():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        for month in MONTHS:
            ws = wb[f"Position Model - {month}"]
            for row in ws.iter_rows(min_row=2, max_row=11):
                for cell in row:
                    if isinstance(cell.value, str):
                        assert "$DV$6" not in cell.value
                        assert "$DQ$7" not in cell.value
