import openpyxl
from build_config_sheet import write_config_sheet

def test_config_sheet_created_with_defined_names():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    assert "Config" in wb.sheetnames
    ws = wb["Config"]
    assert ws["B14"].value == 2.25  # ZThreshold_Tier1 default
    assert "ZThreshold_Tier1" in wb.defined_names
    dn = wb.defined_names["ZThreshold_Tier1"]
    assert dn.attr_text == "Config!$B$14"

def test_all_layout_entries_produce_a_defined_name():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    from config_map import CONFIG_LAYOUT
    for name, cell, _default in CONFIG_LAYOUT:
        assert name in wb.defined_names, f"missing defined name {name}"
        assert wb.defined_names[name].attr_text == f"Config!${cell[0]}${cell[1:]}"
