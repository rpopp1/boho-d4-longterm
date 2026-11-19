import pathlib
import sys

import openpyxl

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_roll_prices import roll_prices_formulas
from multi_month import MONTHS, ROLL_PRICES_HEADERS, roll_prices_col
from roll_fix import PRICE_CHANGE_COLS, price_change_formulas, write_roll_prices_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_roll_prices_blocks_are_contiguous_and_ordered():
    # Each month's 8-column block, one blank column between months.
    assert roll_prices_col("January", "Year") == "A"
    assert roll_prices_col("January", "HO Open") == "H"
    assert roll_prices_col("March", "Year") == "J"
    assert roll_prices_col("December", "Roll Date") == "BM"
    assert roll_prices_col("December", "SBO Settle") == "BO"


def test_price_change_formulas_cover_exactly_z_to_af():
    assert set(price_change_formulas("December", 300)) == set(PRICE_CHANGE_COLS)


def test_price_change_formulas_use_the_months_own_roll_prices_block():
    z = price_change_formulas("December", 300)["Z"]
    assert "'Roll Prices'!$BO$3:$BO$17" in z      # December SBO Settle
    assert "MATCH($A300,'Roll Prices'!$BM$3:$BM$17,0)" in z
    assert "'Roll Prices'!$D$3:$D$17" in price_change_formulas("January", 300)["Z"]


def test_price_change_formulas_fall_back_to_the_plain_day_over_day_change():
    f = price_change_formulas("March", 300)
    # Non-roll day: the held price is the row's own price, minus the previous row's.
    assert f["Z"].endswith(",B300))-B299")
    assert f["AA"].endswith(",C300))-C299")
    assert f["AF"].endswith(",E300))-E299")
    # SBO is cents on Roll Prices but $/lb in the Position Model; HO and D4 aren't scaled.
    assert "held/100" in f["Z"] and "held/100" not in f["AA"]
    # Intraday and overnight legs use the held open, not the Data open.
    assert f["AB"].count("LET(") == 2 and f["AB"].endswith(",F300))")
    assert f["AD"].endswith(",F300))-B299")


def test_file_syntax_prefixes_let_for_openpyxl_written_workbooks():
    f = price_change_formulas("July", 10, file_syntax=True)["Z"]
    assert f.startswith("=_xlfn.LET(_xlpm.held,")
    assert "IF(N(_xlpm.held)>0,_xlpm.held/100,B10)" in f
    assert "_xlfn" not in price_change_formulas("July", 10)["Z"]


def test_roll_prices_formulas_january_holds_next_years_contract():
    jan = roll_prices_formulas("January", 3, 2021)
    assert jan["Year"] == 2021
    assert '"F"&TEXT(MOD((A3+1),100),"00")' in jan["Old Contract"]
    mar = roll_prices_formulas("March", 3, 2021)
    assert '"H"&TEXT(MOD((J3+0),100),"00")' in mar["Old Contract"]
    # Roll date comes from the feed's roll table row for the month.
    assert "$FX$4" in mar["Roll Date"] and "$FY$4" in mar["Roll Date"]
    assert "$FX$10" in roll_prices_formulas("December", 3, 2021)["Roll Date"]


def test_roll_prices_formulas_blank_until_a_roll_date_exists():
    f = roll_prices_formulas("December", 5, 2023)
    for h in ("SBO Settle", "HO Settle", "D4", "SBO Open", "HO Open"):
        assert f[h].startswith('=IF(BM5="","",IFERROR(')


def test_write_roll_prices_sheet_has_headers_and_years_for_every_month():
    wb = openpyxl.Workbook()
    write_roll_prices_sheet(wb)
    ws = wb["Roll Prices"]
    for month in MONTHS:
        assert [ws[f"{roll_prices_col(month, h)}2"].value for h in ROLL_PRICES_HEADERS] == ROLL_PRICES_HEADERS
        assert ws[f"{roll_prices_col(month, 'Year')}3"].value == 2021
        assert ws[f"{roll_prices_col(month, 'Roll Date')}3"].value is None


def test_position_model_writes_roll_aware_price_changes():
    from build_position_model import write_position_model_sheet

    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Position Model - March"]
    for col, formula in price_change_formulas("March", 5, file_syntax=True).items():
        assert ws[f"{col}5"].value == formula
    # Price levels still read March's own (relocated) Data block.
    assert ws["B5"].value.startswith("=Data!")
