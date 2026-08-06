import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from config_map import CONFIG_LAYOUT


def write_config_sheet(wb: "openpyxl.Workbook") -> None:
    if "Config" in wb.sheetnames:
        del wb["Config"]
    ws = wb.create_sheet("Config")
    ws["A1"] = "Field"
    ws["B1"] = "Value"
    for name, cell, default in CONFIG_LAYOUT:
        row = int("".join(ch for ch in cell if ch.isdigit()))
        label_cell = f"A{row}"
        ws[label_cell] = name
        ws[cell] = default
        col_letters = "".join(ch for ch in cell if ch.isalpha())
        ref = f"Config!${col_letters}${row}"
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
