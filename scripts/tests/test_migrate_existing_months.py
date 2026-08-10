import openpyxl, pathlib, tempfile, os
from build_template import build_template
from migrate_existing_months import migrate_month_data, _read_config_values
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
from config_map import CONFIG_FIELDS

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


def test_read_config_values_reads_back_every_config_field():
    with tempfile.TemporaryDirectory() as tmp:
        template_path = os.path.join(tmp, "single_month_template.xlsx")
        build_template(SOURCE_PATH, template_path, num_rows=10)

        # Simulate a previously-activated single-month template: set a
        # couple of Config!B values directly via openpyxl, as if a user had
        # calibrated that month's book, then verify _read_config_values
        # reads them back correctly for every CONFIG_FIELDS entry.
        wb = openpyxl.load_workbook(template_path, data_only=False)
        ws = wb["Config"]
        row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
        ws[f"B{row_by_field['ContractLabel']}"] = "March 2027"
        ws[f"B{row_by_field['ZThreshold_Tier1']}"] = 2.5
        wb.save(template_path)

        values = _read_config_values(template_path, "B")

        assert set(values.keys()) == {name for name, _row, _default in CONFIG_FIELDS}
        assert values["ContractLabel"] == "March 2027"
        assert values["ZThreshold_Tier1"] == 2.5


def test_migrate_month_data_does_not_touch_other_months_blocks():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December")

        assert wb["Data"]["B2"].value is None  # January's block, offset 0, untouched
        source_wb.close()
