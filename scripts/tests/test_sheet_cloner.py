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
