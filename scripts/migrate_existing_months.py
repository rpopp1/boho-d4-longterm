import pathlib
import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col


def migrate_month_data(wb, source_wb, month: str, config_values: dict | None = None) -> None:
    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
    jm_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)

    data_ws = wb["Data"]
    source_data = source_wb["Data"]
    # max_col=14 (not 20): columns O-T hold the S:T reference-block cells
    # that build_static_sheets.py deliberately writes as formulas pointing
    # at that month's Config column (=Config!<col>9, etc.) -- copying the
    # source's frozen static values over those would decouple them from
    # Config for the migrated month. Only Date/BOHO/D4 + Settlement + Open
    # pricing (columns A-N) are actual pasted-value data to carry over.
    for row in source_data.iter_rows(min_row=2, max_row=source_data.max_row, max_col=14):
        for cell in row:
            if cell.value is not None:
                col_letter, _ = coordinate_from_string(cell.coordinate)
                data_ws[f"{shift_col(col_letter, data_offset)}{cell.row}"] = cell.value

    jm_ws = wb["Johansen_Monthly"]
    source_jm = source_wb["Johansen_Monthly"]
    for row in source_jm.iter_rows(min_row=2, max_row=source_jm.max_row, min_col=6, max_col=9):  # F:I
        for cell in row:
            if cell.value is not None:
                col_letter, _ = coordinate_from_string(cell.coordinate)
                jm_ws[f"{shift_col(col_letter, jm_offset)}{cell.row}"] = cell.value

    if config_values:
        config_ws = wb["Config"]
        row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
        month_col = CONFIG_MONTH_COLUMNS[month]
        for field, value in config_values.items():
            config_ws[f"{month_col}{row_by_field[field]}"] = value


def _read_config_values(config_path: str, month_col_in_source: str) -> dict:
    """Read a previously-activated single-month template's own Config
    sheet (column B, the only column it has) and return {field: value} for
    every field in CONFIG_FIELDS, so that month's already-calibrated real
    values can be carried into the multi-month workbook's Config column for
    that same month, rather than reverting to the generic Dec-derived
    defaults build_config_sheet.py starts every month at."""
    wb = openpyxl.load_workbook(config_path, data_only=True)
    ws = wb["Config"]
    row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
    values = {field: ws[f"{month_col_in_source}{row}"].value for field, row in row_by_field.items()}
    wb.close()
    return values


def run_migration(template_path: str, dec_source_path: str, january_source_path: str, output_path: str) -> None:
    wb = openpyxl.load_workbook(template_path, data_only=False)

    dec_source_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    try:
        migrate_month_data(wb, dec_source_wb, "December")
    finally:
        dec_source_wb.close()

    january_source_wb = openpyxl.load_workbook(january_source_path, data_only=True)
    try:
        january_config = _read_config_values(january_source_path, "B")
        migrate_month_data(wb, january_source_wb, "January", config_values=january_config)
    finally:
        january_source_wb.close()

    wb.save(output_path)


if __name__ == "__main__":
    scripts_dir = pathlib.Path(__file__).parent
    root = scripts_dir.parent
    template_path = str(root / "Johansen_MultiMonth_Template_Engine.xlsx")
    dec_source_path = (scripts_dir / "DEC_SOURCE_PATH.txt").read_text().strip()
    january_source_path = str(root / "Johansen_January_BOHO_D4_Strategy.xlsx")
    output_path = str(root / "Johansen_BOHO_D4_Strategy_AllMonths.xlsx")
    run_migration(template_path, dec_source_path, january_source_path, output_path)
    print(f"Migrated workbook written to {output_path}")
