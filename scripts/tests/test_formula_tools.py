import pathlib
from formula_tools import extract_row_formulas
from formula_tools import shift_formula_rows
from config_map import CONSTANT_MAP

SOURCE = pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt"
SOURCE_PATH = SOURCE.read_text().strip()

def test_extract_position_model_row2():
    formulas = extract_row_formulas(SOURCE_PATH, "Position Model", row=2, max_col=12)
    assert formulas["A"] == "=Data!A2"
    assert formulas["I"] == (
        "=(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D2)"
        "+(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$I:$I)*$E2)"
    )
    assert formulas["J"] == "=AVERAGE($I$2:I2)"
    assert formulas["K"] == "=_xlfn.STDEV.S($I$2:I2)"
    assert formulas["L"] == "=(I2-J2)/K2"

def test_extract_diff_lag_row2():
    formulas = extract_row_formulas(SOURCE_PATH, "DiffLag", row=2, max_col=8)
    assert formulas["A"] == 1
    assert formulas["B"] == '=IF(Data!B2="","",Data!A2)'
    assert formulas["C"] == "=Data!B2"
    assert formulas["D"] == "=Data!C2"
    assert formulas["E"] == '=IF(OR(C2="",C1=""),"",C2-C1)'

def test_extract_array_formula_unwrapped():
    formulas = extract_row_formulas(SOURCE_PATH, "Johansen_Monthly", row=2, max_col=1)
    assert formulas["A"] == "=_xlfn._xlws.FILTER(DiffLag!$B$2:$B$1048576,DiffLag!$B$2:$B$1048576<>\"\")"


def test_shift_simple_relative_ref():
    assert shift_formula_rows("=Data!A2", 1) == "=Data!A3"
    assert shift_formula_rows("=Data!A2", 5) == "=Data!A7"

def test_shift_leaves_absolute_row_unchanged():
    assert shift_formula_rows("=(I2-$J$2)/K2", 3) == "=(I5-$J$2)/K5"

def test_shift_leaves_absolute_col_but_shifts_row():
    assert shift_formula_rows("=$A2+$B2", 2) == "=$A4+$B4"

def test_shift_expanding_range_pattern():
    # AVERAGE($I$2:I2) filled down 5 rows -> AVERAGE($I$2:I7)
    assert shift_formula_rows("=AVERAGE($I$2:I2)", 5) == "=AVERAGE($I$2:I7)"

def test_shift_cross_sheet_reference():
    assert shift_formula_rows(
        "=IF($L3<=$DV$6,0,IF(AND($M2=0,$L3>=M$1),$DX$3,M2))", 4
    ) == "=IF($L7<=$DV$6,0,IF(AND($M6=0,$L7>=M$1),$DX$3,M6))"

def test_shift_ignores_defined_name_tokens():
    # A bare name with no row digits must never be touched
    assert shift_formula_rows("=ZThreshold_Tier1*2", 10) == "=ZThreshold_Tier1*2"

def test_shift_lookup_with_full_column_refs_unchanged():
    formula = "=(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D2)"
    shifted = shift_formula_rows(formula, 1)
    assert shifted == "=(LOOKUP($A3-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D3)"

def test_shift_does_not_corrupt_trailing_identifier_text():
    assert shift_formula_rows("=A2Value", 1) == "=A2Value"

def test_shift_does_not_corrupt_function_name_collision():
    assert shift_formula_rows("=LOG10(A2)", 1) == "=LOG10(A3)"


def test_rewrite_single_constant():
    from formula_tools import rewrite_constants
    formula = "=IF($L3<=$DV$6,0,IF(AND($M2=0,$L3>=M$1),$DX$3,M2))"
    result = rewrite_constants(formula, CONSTANT_MAP)
    assert "$DV$6" not in result
    assert "ZExitThreshold" in result
    assert "$DX$3" in result  # not in the map, left untouched

def test_rewrite_leaves_unmapped_addresses_untouched():
    from formula_tools import rewrite_constants
    formula = "=Data!A2"
    assert rewrite_constants(formula, CONSTANT_MAP) == "=Data!A2"

def test_rewrite_multiple_constants_in_one_formula():
    from formula_tools import rewrite_constants
    # $DK$5 is HOGalPerContract per Task 4's reconciled map (confirmed against the
    # live workbook's real labels); $DK$2 is a distinct constant (ZLCentsPerLB),
    # so this test uses $DK$5 twice to exercise multi-occurrence rewriting.
    formula = "=S2*($Z2*$DK$5-$AA2*$DK$5)"
    result = rewrite_constants(formula, CONSTANT_MAP)
    assert result == "=S2*($Z2*HOGalPerContract-$AA2*HOGalPerContract)"
