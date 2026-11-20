import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from patch_config_summary import formulas, last_row_expr


def test_formulas_cover_the_six_open_rows():
    assert set(formulas("December")) == {34, 35, 38, 39, 44, 45}


def test_ranges_start_where_the_existing_rows_start_and_end_at_the_last_dated_row():
    f = formulas("March")
    assert "'Position Model - March'!$BT$741:INDEX('Position Model - March'!$BT$1:$BT$3001,last)" in f[34]
    assert "'Position Model - March'!$BT$29:INDEX(" in f[35]
    assert f[38].endswith("(AVERAGE(r)/STDEV.S(r))*SQRT(252))")
    assert "INDEX('Position Model - March'!$CC$1:$CC$3001,last)" in f[45]
    assert last_row_expr("March").startswith("MATCH(LOOKUP(2,1/('Position Model - March'!$A$2:$A$3001>DATE(2015,1,1))")
