import openpyxl
from build_static_sheets import write_notes_sheet, write_data_sheet

def test_notes_sheet_describes_config_workflow():
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    text = " ".join(str(c.value) for c in wb["Notes"]["A"] if c.value)
    assert "Config" in text
    assert "Data" in text


def test_notes_sheet_documents_johansen_monthly_manual_append_step():
    # MAJOR 5: Johansen_Monthly!F:I is manually maintained (Task 11 v2
    # redesign), not auto-computed. The Notes sheet used to claim the
    # opposite ("Johansen_Monthly ... recompute automatically ... no formulas
    # need editing"), which is an operational hazard for an actively-traded
    # book -- someone could go a month without advancing the log and the
    # Position Model's LOOKUP against Johansen_Monthly!G:I would silently
    # have nothing new to find.
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    text = " ".join(str(c.value) for c in wb["Notes"]["A"] if c.value)
    assert "Johansen_Monthly" in text
    assert "F:I" in text or ("F" in text and "manual" in text.lower())
    assert "manual" in text.lower()
    # Must no longer claim Johansen_Monthly recomputes automatically with
    # nothing to edit -- that claim is false as of the Task 11 v2 redesign.
    assert "Johansen_Monthly, Monthly Beta, and Position" not in text

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
