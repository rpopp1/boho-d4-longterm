import pathlib
import sys

import openpyxl
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_dashboard import CALC_SHEET, build_calc_sheet, build_dashboard_sheet
from dashboard_today import TODAY_FIELDS
from multi_month import MONTHS


@pytest.fixture(scope="module")
def built():
    wb = openpyxl.Workbook()
    build_calc_sheet(wb)
    build_dashboard_sheet(wb)
    return wb


def test_today_block_reads_the_last_row(built):
    calc = built[CALC_SHEET]
    assert TODAY_FIELDS["pos"] == "CC" and TODAY_FIELDS["row"] == "CB"
    assert calc["CC2"].value == "=INDEX('Position Model - January'!$V$2:$V$3001,$CB2)"
    assert calc["CB2"].value == "=MATCH($CA2,'Position Model - January'!$A$2:$A$3001,0)"
    assert calc["CS9"].value == "=Config!$I$31"
    assert calc["CX9"].value == "='Position Model - December'!$DV$7"


def test_today_rows_in_month_order(built):
    dash = built["Dashboard"]
    assert [dash[f"A{r}"].value for r in range(6, 14)] == MONTHS
    assert dash["A14"].value == "Portfolio"
    assert dash["J5"].value == "Next Signal"
