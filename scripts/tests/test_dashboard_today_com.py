"""Evaluates the Today section's text formulas in a real (isolated) Excel."""
import datetime
import pathlib
import sys
import tempfile

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from dashboard_today import TODAY_FIELDS, next_signal_formula, position_formula

pytest.importorskip("win32com.client")
from excel_com import isolated_workbook  # noqa: E402

BASE = dict(asof=datetime.date(2026, 11, 18), row=1, pos=0, active=0, z=0.74, pt_exit=0, exit_z=None,
            regime="Normal", pctl=0.5, vol=1.0, shock=0, trade_pnl=0,
            s1=0, s2=0, s3=0, l3=0, l2=0, l1=0, reentry=1,
            t1=2.25, t2=2.1, t3=1.75, exit_short=0.1, exit_long=-0.1)

CASES = {
    "flat": dict(z=0.74),
    "short_full": dict(pos=-1, active=1, z=1.52, s1=-0.25, s2=-0.5, s3=-0.25),
    "short_partial": dict(pos=-0.25, active=1, z=1.80, s3=-0.25),
    "long_on_tier": dict(pos=0.75, active=1, z=-2.10, l3=0.25, l2=0.5),
    "pt_reentry": dict(pos=-1, active=0, z=1.52, pt_exit=1, exit_z=1.52, s1=-0.25, s2=-0.5, s3=-0.25),
    "pt_no_reentry": dict(pos=-1, active=0, z=1.52, pt_exit=1, exit_z=1.52, reentry=0,
                          s1=-0.25, s2=-0.5, s3=-0.25),
    "stopped_short": dict(pos=-1, active=0, z=0.74, s1=-0.25, s2=-0.5, s3=-0.25),
}


@pytest.fixture(scope="module")
def results():
    keys = list(TODAY_FIELDS)
    wb = openpyxl.Workbook()
    ws = wb.active
    names = list(CASES)
    for i, name in enumerate(names):
        r = i + 1
        vals = {**BASE, **CASES[name]}
        refs = {k: f"{get_column_letter(1 + j)}{r}" for j, k in enumerate(keys)}
        for j, k in enumerate(keys):
            ws.cell(row=r, column=1 + j, value=vals[k])
        ws[f"AY{r}"] = position_formula(refs)
        ws[f"AZ{r}"] = next_signal_formula(refs)
    path = pathlib.Path(tempfile.mkdtemp()) / "today_cases.xlsx"
    wb.save(path)
    out = {}
    with isolated_workbook(path) as (excel, xwb):
        excel.CalculateFull()
        sh = xwb.Worksheets(1)
        for i, name in enumerate(names):
            out[name] = (sh.Range(f"AY{i + 1}").Value, sh.Range(f"AZ{i + 1}").Value)
        xwb.Saved = True
    return out


def test_flat_no_signal(results):
    assert results["flat"] == ("Flat", "Short at Z ≥ 1.75 (now 0.74)")


def test_short_in_trade(results):
    assert results["short_full"] == ("Short", "Exit at Z ≤ 0.10 (now 1.52)")


def test_partial_short_names_next_add_tier(results):
    assert results["short_partial"][1] == "Exit at Z ≤ 0.10 (now 1.80) · adds at Z ≥ 2.10"


def test_long_on_a_threshold(results):
    pos, text = results["long_on_tier"]
    assert pos == "Long"
    assert text == "Exit at Z ≥ -0.10 (now -2.10) · adds at Z ≤ -2.25"
    assert "99" not in text


def test_profit_take_exit_with_reentry(results):
    assert results["pt_reentry"] == ("Flat (exited)",
                                     "Flat, profit-take exit. Re-entry 12/1 if Z ≥ 1.52 (now 1.52)")


def test_profit_take_exit_without_reentry(results):
    assert results["pt_no_reentry"][1].startswith("Flat, profit-take exit. Waiting for signal reset")


def test_stopped_out_short(results):
    assert results["stopped_short"] == ("Flat (exited)",
                                        "Flat, stopped out. Waiting for signal reset at Z ≤ 0.10 (now 0.74)")
