import openpyxl, pathlib, tempfile, os
from build_template import build_template
from migrate_existing_months import migrate_month_data
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_migrate_month_data_writes_data_and_log_into_that_months_block():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December")

        data_offset = column_offset("December", DATA_BLOCK_WIDTH, GAP)
        assert wb["Data"][f"{shift_col('B', data_offset)}2"].value is not None

        jm_offset = column_offset("December", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
        assert wb["Johansen_Monthly"][f"{shift_col('F', jm_offset)}3"].value is not None
        source_wb.close()


def test_migrate_month_data_writes_config_values_when_given():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December", config_values={"ZThreshold_Tier1": 2.25})

        assert wb["Config"]["I14"].value == 2.25  # December's Config column is I
        source_wb.close()


def test_migrate_month_data_does_not_touch_other_months_blocks():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December")

        assert wb["Data"]["B2"].value is None  # January's block, offset 0, untouched
        source_wb.close()
