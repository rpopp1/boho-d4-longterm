import openpyxl
from build_config_sheet import write_config_sheet
from multi_month import MONTHS
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS


def test_config_sheet_has_month_header_row():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    assert ws["A1"].value == "Field"
    assert ws["B1"].value == "January"
    assert ws["I1"].value == "December"


def test_config_sheet_every_month_gets_the_default_value():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for month in MONTHS:
        col = CONFIG_MONTH_COLUMNS[month]
        assert ws[f"{col}14"].value == 2.25  # ZThreshold_Tier1 default, every month


def test_config_sheet_identity_fields_default_to_template_placeholder():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for month in MONTHS:
        col = CONFIG_MONTH_COLUMNS[month]
        assert ws[f"{col}2"].value == "TEMPLATE"  # ContractLabel


def test_config_sheet_no_defined_names():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    assert len(wb.defined_names) == 0


def test_config_sheet_labels_column_a():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for name, row, _default in CONFIG_FIELDS:
        assert ws[f"A{row}"].value == name
