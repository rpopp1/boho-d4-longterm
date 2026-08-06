import openpyxl
from build_static_sheets import write_notes_sheet, write_data_sheet

def test_notes_sheet_describes_config_workflow():
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    text = " ".join(str(c.value) for c in wb["Notes"]["A"] if c.value)
    assert "Config" in text
    assert "Data" in text

def test_data_sheet_headers_and_no_price_values():
    wb = openpyxl.Workbook()
    write_data_sheet(wb)
    ws = wb["Data"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "BOHO"
    assert ws["C1"].value == "D4"
    assert ws["S2"].value == "ZL cents/LBs"
    assert ws["T2"].value == "=ZLCentsPerLB"
    assert ws["B2"].value is None  # no live price data in the template
