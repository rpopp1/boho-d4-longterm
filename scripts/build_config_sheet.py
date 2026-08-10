import openpyxl
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS
from multi_month import MONTHS


def write_config_sheet(wb: "openpyxl.Workbook") -> None:
    if "Config" in wb.sheetnames:
        del wb["Config"]
    ws = wb.create_sheet("Config")
    ws["A1"] = "Field"
    for month in MONTHS:
        ws[f"{CONFIG_MONTH_COLUMNS[month]}1"] = month
    for name, row, default in CONFIG_FIELDS:
        ws[f"A{row}"] = name
        for month in MONTHS:
            ws[f"{CONFIG_MONTH_COLUMNS[month]}{row}"] = default
