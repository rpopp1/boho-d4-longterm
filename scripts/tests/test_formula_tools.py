import pathlib
from formula_tools import extract_row_formulas

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
