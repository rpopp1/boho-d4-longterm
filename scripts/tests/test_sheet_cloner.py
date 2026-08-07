import openpyxl
from sheet_cloner import fill_sheet_from_pattern
import pathlib

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

# These tests exist to verify fill_sheet_from_pattern's generic mechanism, not
# config_map.py's specific content, so they use a small local dict rather than
# importing the (now-removed) global CONSTANT_MAP.
_TEST_CONSTANT_MAP = {"$DV$6": "ZExitThreshold", "$DK$5": "HOGalPerContract"}

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
        constant_map=_TEST_CONSTANT_MAP,
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
        constant_map=_TEST_CONSTANT_MAP,
    )
    m_formula = ws["M4"].value
    assert "$DV$6" not in m_formula
    assert "ZExitThreshold" in m_formula

def test_fill_raises_clearly_on_array_formula_instead_of_silently_corrupting():
    wb = openpyxl.Workbook()
    ws = wb.active
    import pytest
    with pytest.raises(NotImplementedError, match="dynamic-array"):
        fill_sheet_from_pattern(
            target_ws=ws,
            source_path=SOURCE_PATH,
            source_sheet="Johansen_Monthly",
            pattern_row=2,
            target_start_row=2,
            num_rows=10,
            max_col=1,  # column A only — was a FILTER spill anchor as of this writing
            constant_map=_TEST_CONSTANT_MAP,
        )

def test_fill_with_column_shift_relocates_block_and_self_references():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=3,
        max_col=8,
        constant_map=_TEST_CONSTANT_MAP,
        column_shift_map={None: (10, None), "Data": (5, "Data")},
    )
    # DiffLag's own column B (source: '=IF(Data!B2="","",Data!A2)') physically
    # lands 10 columns over, at column L (B=2+10=12); its "Data!" cross-refs
    # shift by Data's own delta (5), not DiffLag's (10).
    assert ws["L2"].value == '=IF(Data!G2="","",Data!F2)'
    # DiffLag's own column E (source: '=IF(OR(C2="",C1=""),"",C2-C1)') is a
    # self-reference to DiffLag's own C column (unqualified -- no "Data!").
    # It physically lands at E(5)+10=15='O', and its self-reference shifts
    # by DiffLag's own delta (10): C(3)+10=13='M'.
    assert ws["O2"].value == '=IF(OR(M2="",M1=""),"",M2-M1)'
    assert ws["O3"].value == '=IF(OR(M3="",M2=""),"",M3-M2)'


def test_fill_without_column_shift_map_is_unchanged():
    # column_shift_map omitted entirely -- must behave exactly as before.
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=3,
        max_col=8,
        constant_map=_TEST_CONSTANT_MAP,
    )
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
