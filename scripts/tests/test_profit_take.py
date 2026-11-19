import pathlib
import sys

import openpyxl

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_config_sheet import write_config_sheet
from build_position_model import write_position_model_sheet
from profit_take import (
    FIRST_ROW, HEADERS, HELPER_COLS, REWRITTEN_COLS, config_refs, original_formula, overlay_formulas,
)

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_config_refs_point_at_the_months_own_config_column():
    assert config_refs("December") == {
        "stop": "Config!$I$27", "tp": "Config!$I$30", "reentry": "Config!$I$31"}
    assert config_refs("January")["tp"] == "Config!$B$30"


def test_overlay_covers_the_rewritten_and_helper_columns():
    assert set(overlay_formulas("December", 500)) == set(REWRITTEN_COLS) | set(HELPER_COLS)


def test_original_formula_matches_the_live_workbook_shapes():
    assert original_formula("BP", "December", 1000, absolute_stop=False) == \
        "=IF(AND(BM1000=1,BO1000<=Config!I27,BO999>Config!I27),1,0)"
    assert original_formula("BP", "December", 1461) == \
        "=IF(AND(BM1461=1,BO1461<=Config!$I$27,BO1460>Config!$I$27),1,0)"
    assert original_formula("BT", "December", 21) == "=IF(BS21=1,BL21,IF(BQ21<>BQ20,BL21,0))"


def test_profit_take_is_off_when_its_config_cell_is_blank():
    f = overlay_formulas("December", 500)
    # N() turns a blank or text Config cell into 0, so the profit-take never fires.
    assert "AND(N(Config!$I$30)>0,BO500>=Config!$I$30)" in f["BP"]
    # The stop keeps its first-crossing test; a re-entry day crosses from zero.
    assert "AND(BO500<=Config!$I$27,IF(EO500=1,0,BO499)>Config!$I$27)" in f["BP"]


def test_reentry_requires_switch_month_start_profit_exit_and_stronger_z():
    eo = overlay_formulas("March", 300)["EO"]
    for part in ("Config!$C$31=1", "BN300<>1", "IFERROR(BQ299,0)=1", "IFERROR(BM299,0)=1", "BL300<>0",
                 "EK300=1", "IFERROR(EN299,0)=1", "SIGN(L300)=SIGN(EM299)", "ABS(L300)>=ABS(EM299)"):
        assert part in eo
    assert eo.startswith("=IFERROR(") and eo.endswith(",1,0),0)")


def test_a_reentry_starts_a_fresh_sub_trade():
    f = overlay_formulas("December", 500)
    assert "OR(BN500=1,EO500=1),BL500" in f["BO"]
    assert f["BQ"] == "=IF(BN500=1,0,IF(EO500=1,BP500,IF(IFERROR(BQ499,0)=1,1,BP500)))"
    assert "OR(BQ500<>BQ499,EO500=1),BL500" in f["BT"]
    assert "OR(BN500=1,EO500=1),BT500" in f["BU"]


def test_exit_bookkeeping_stores_z_and_reason():
    f = overlay_formulas("December", 500)
    assert f["EL"] == "=IFERROR(IF(AND(BQ500=1,OR(IFERROR(BQ499,0)=0,EO500=1)),1,0),0)"
    assert f["EM"] == '=IF(EL500=1,IF(ISNUMBER(L500),L500,""),IF(IFERROR(BN500,0)=1,"",IFERROR(EM499,"")))'
    assert "IF(AND(N(Config!$I$30)>0,BO500>=Config!$I$30),1,0)" in f["EN"]


def test_generator_writes_the_overlay_and_config_rows():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    assert wb["Config"]["A30"].value == "ProfitTakeUSD"
    assert wb["Config"]["I30"].value is None
    assert wb["Config"]["A31"].value == "ReentryMonthStart"
    assert wb["Config"]["I31"].value == 0

    write_position_model_sheet(wb, SOURCE_PATH, "December", num_rows=40)
    ws = wb["Position Model - December"]
    for col, header in HEADERS.items():
        assert ws[f"{col}1"].value == header
    for col, formula in overlay_formulas("December", FIRST_ROW).items():
        assert ws[f"{col}{FIRST_ROW}"].value == formula
    assert ws["BP41"].value == overlay_formulas("December", 41)["BP"]
    assert ws[f"EO{FIRST_ROW - 1}"].value is None
