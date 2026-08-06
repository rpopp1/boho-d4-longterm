import openpyxl
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP
import pathlib

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_fill_diff_lag_pattern_five_rows():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=5,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B3"].value == '=IF(Data!B3="","",Data!A3)'
    assert ws["B6"].value == '=IF(Data!B6="","",Data!A6)'
    assert ws["C4"].value == "=Data!B4"

def test_fill_position_model_rewrites_constants():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="Position Model",
        pattern_row=3,
        target_start_row=3,
        num_rows=3,
        max_col=13,
        constant_map=CONSTANT_MAP,
    )
    m_formula = ws["M4"].value
    assert "$DV$6" not in m_formula
    assert "ZExitThreshold" in m_formula

def test_fill_does_not_duplicate_array_formula_spill_anchor():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="Johansen_Monthly",
        pattern_row=2,
        target_start_row=2,
        num_rows=10,
        max_col=1,  # column A only — the FILTER spill anchor
        constant_map=CONSTANT_MAP,
    )
    from openpyxl.worksheet.formula import ArrayFormula
    a2 = ws["A2"].value
    assert isinstance(a2, ArrayFormula)
    assert "FILTER(DiffLag!$B$2:$B$1048576" in a2.text
    # ref MUST span more than one cell — a single-cell ref makes Excel treat this
    # as a legacy CSE array formula (returns only the first value) instead of a
    # real dynamic-array spill. This exact defect was caught by Task 15's Dec
    # replay validation (silent #N/A cascade), so this assertion guards against
    # regressing it.
    assert a2.ref == "A2:A11"
    # Every other target row must be left untouched — writing the same spill
    # anchor into A3..A11 would corrupt the file with #SPILL! errors in Excel.
    for row in range(3, 12):
        assert ws[f"A{row}"].value is None
