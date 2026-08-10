import pathlib
from formula_tools import extract_row_formulas
from formula_tools import shift_formula_rows

SOURCE = pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt"
SOURCE_PATH = SOURCE.read_text().strip()

# These tests exist to verify rewrite_constants's generic substitution
# mechanism, not config_map.py's specific content, so they use a small local
# dict rather than importing the (now-removed) global CONSTANT_MAP.
_TEST_CONSTANT_MAP = {"$DV$6": "ZExitThreshold", "$DK$5": "HOGalPerContract"}

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
    result = rewrite_constants(formula, _TEST_CONSTANT_MAP)
    assert "$DV$6" not in result
    assert "ZExitThreshold" in result
    assert "$DX$3" in result  # not in the map, left untouched

def test_rewrite_leaves_unmapped_addresses_untouched():
    from formula_tools import rewrite_constants
    formula = "=Data!A2"
    assert rewrite_constants(formula, _TEST_CONSTANT_MAP) == "=Data!A2"

def test_rewrite_multiple_constants_in_one_formula():
    from formula_tools import rewrite_constants
    # $DK$5 is HOGalPerContract per Task 4's reconciled map (confirmed against the
    # live workbook's real labels); $DK$2 is a distinct constant (ZLCentsPerLB),
    # so this test uses $DK$5 twice to exercise multi-occurrence rewriting.
    formula = "=S2*($Z2*$DK$5-$AA2*$DK$5)"
    result = rewrite_constants(formula, _TEST_CONSTANT_MAP)
    assert result == "=S2*($Z2*HOGalPerContract-$AA2*HOGalPerContract)"


def test_unwrap_scalar_array_formula_single_cell_ref():
    # MINOR 6: Matrices!H7 is a live example -- an ArrayFormula whose ref is
    # just its own coordinate ("H7"), a legacy CSE array formula stored for
    # historical reasons that doesn't need array/spill semantics.
    from formula_tools import unwrap_scalar_array_formula
    from openpyxl.worksheet.formula import ArrayFormula

    value = ArrayFormula(ref="H7", text="=INDEX(DiffLag!E:E,$B$2)")
    assert unwrap_scalar_array_formula(value, "H7") == "=INDEX(DiffLag!E:E,$B$2)"


def test_unwrap_scalar_array_formula_raises_on_multi_cell_ref():
    from formula_tools import unwrap_scalar_array_formula
    from openpyxl.worksheet.formula import ArrayFormula
    import pytest

    value = ArrayFormula(ref="H7:H10", text="=SOME_SPILL_FORMULA()")
    with pytest.raises(NotImplementedError):
        unwrap_scalar_array_formula(value, "H7")


def test_unwrap_scalar_array_formula_passes_through_non_array_values():
    from formula_tools import unwrap_scalar_array_formula

    assert unwrap_scalar_array_formula("=A1+B1", "C1") == "=A1+B1"
    assert unwrap_scalar_array_formula(5, "C1") == 5
    assert unwrap_scalar_array_formula(None, "C1") is None


def test_shift_columns_unqualified_self_reference():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=B3+B6", {None: (2, None)}) == "=D3+D6"

def test_shift_columns_cross_sheet_reference_keeps_qualifier():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=Johansen_Monthly!$K$3", {"Johansen_Monthly": (13, "Johansen_Monthly")}
    ) == "=Johansen_Monthly!$X$3"

def test_shift_columns_full_column_reference():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=DiffLag!E:E", {"DiffLag": (5, "DiffLag")}) == "=DiffLag!J:J"

def test_shift_columns_dollar_anchored_full_column_reference():
    # multi-month Task 11: "$G:$G" (anchored) used to not match _COL_REF_RE at
    # all, so shift_formula_columns silently returned the formula UNCHANGED --
    # every month's Position Model would have kept reading January's
    # Johansen_Monthly walk-forward log. The $ anchors must be preserved.
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=Johansen_Monthly!$G:$G", {"Johansen_Monthly": (13, "Johansen_Monthly")}
    ) == "=Johansen_Monthly!$T:$T"


def test_shift_columns_real_position_model_lookup_formula():
    # Verbatim from the live source workbook's Position Model!I2. Only the
    # Johansen_Monthly! refs move; $A2/D2/$E2 are Position Model's own
    # (unqualified) self-references and there is deliberately no None key.
    from formula_tools import shift_formula_columns
    formula = (
        "=(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D2)"
        "+(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$I:$I)*$E2)"
    )
    assert shift_formula_columns(formula, {"Johansen_Monthly": (13, "Johansen_Monthly")}) == (
        "=(LOOKUP($A2-1, Johansen_Monthly!$T:$T, Johansen_Monthly!$U:$U)*D2)"
        "+(LOOKUP($A2-1, Johansen_Monthly!$T:$T, Johansen_Monthly!$V:$V)*$E2)"
    )


def test_shift_columns_mixed_anchoring_full_column_reference():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=Data!A:$C", {"Data": (22, "Data")}) == "=Data!W:$Y"


def test_shift_columns_anchored_row_range_still_matches_cell_alternative():
    # Regression guard: adding the $ to the full-column alternative must not
    # let it swallow ordinary anchored cell ranges like $AH$3:$AH$301.
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=SUMIFS(Data!$B$3:$B$301,Data!$A$3:$A$301,1)", {"Data": (22, "Data")}
    ) == "=SUMIFS(Data!$X$3:$X$301,Data!$W$3:$W$301,1)"


def test_shift_columns_strips_qualifier_after_sheet_merge():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=Matrices!C50", {"Matrices": (10, None)}) == "=M50"

def test_shift_columns_renames_qualifier_after_sheet_merge():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=Johansen!D18", {"Johansen": (10, "Matrices_Johansen_Results")}
    ) == "=Matrices_Johansen_Results!N18"

def test_shift_columns_leaves_unmapped_sheet_untouched():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=Config!C14", {"DiffLag": (5, "DiffLag")}) == "=Config!C14"

def test_shift_columns_leaves_defined_name_tokens_untouched():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=ZThreshold_Tier1*2", {None: (10, None)}) == "=ZThreshold_Tier1*2"

def test_shift_columns_ignores_dollar_anchoring():
    # Unlike shift_formula_rows, $ never blocks a column shift here -- see the
    # shift_formula_columns docstring for why.
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=$B$5", {None: (3, None)}) == "=$E$5"

def test_shift_columns_real_matrices_formula():
    # Verbatim from the live source workbook's Matrices!C7 (read-only inspection).
    from formula_tools import shift_formula_columns
    formula = (
        '=IF($B$5=0,"",SUMPRODUCT(INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2), '
        'INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2))/$B$5)'
    )
    result = shift_formula_columns(formula, {None: (10, None), "DiffLag": (5, "DiffLag")})
    assert result == (
        '=IF($L$5=0,"",SUMPRODUCT(INDEX(DiffLag!J:J,$L$4):INDEX(DiffLag!J:J,$L$2), '
        'INDEX(DiffLag!J:J,$L$4):INDEX(DiffLag!J:J,$L$2))/$L$5)'
    )
