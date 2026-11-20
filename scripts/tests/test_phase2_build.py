import pathlib
import sys

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from phase2.build_test_workbook import check_output_path, new_error_sheets, output_name, protected_paths


def test_refuses_existing_or_protected_paths(tmp_path):
    existing = tmp_path / "exists.xlsx"
    existing.write_bytes(b"x")
    with pytest.raises(ValueError):
        check_output_path(existing, [])
    protected = tmp_path / "Sandbox.xlsx"
    with pytest.raises(ValueError):
        check_output_path(protected, [protected])
    check_output_path(tmp_path / "new.xlsx", [protected])


def test_protected_paths_cover_every_non_phase2_workbook(tmp_path):
    for name in ("Knox_Multimonth_BOHO_D4_Strategy - Roll Fix Sandbox.xlsx", "Live.xlsx",
                 "Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Baseline DRY RUN.xlsx"):
        (tmp_path / name).write_bytes(b"x")
    names = {p.name for p in protected_paths(tmp_path)}
    assert names == {"Knox_Multimonth_BOHO_D4_Strategy - Roll Fix Sandbox.xlsx", "Live.xlsx"}


def test_output_names_follow_the_mode(tmp_path):
    assert output_name(tmp_path, "adjusted").name == \
        "Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Back-Adjusted TEST.xlsx"
    assert output_name(tmp_path, "dry").name == "Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Baseline DRY RUN.xlsx"
    assert output_name(tmp_path, "jump_only").name ==         "Knox_Multimonth_BOHO_D4_Strategy - Phase 2 Jump-Only TEST.xlsx"
    with pytest.raises(ValueError):
        output_name(tmp_path, "other")


def test_jump_only_keeps_stored_weights_and_adjusts_prices(monkeypatch):
    from datetime import date
    import phase2.build_test_workbook as b
    from phase2.inputs import MonthInputs
    d = [date(2021, 1, 4), date(2021, 1, 5), date(2021, 1, 6)]
    m = MonthInputs("May", [2, 3, 4], d, [1.0, 1.1, 2.0], [1.0] * 3, [0.0] * 3, [0.0] * 3,
                    rolls=[(d[2], 20.0, 0.5)], stored=[(4, d[2], 0.6, -0.8)])
    monkeypatch.setattr(b, "load_inputs", lambda _p: {"May": m})
    writes, _ = b.plan_writes(None, "jump_only")
    rows, boho, weights = writes["May"]
    gap = 2.0 - (20.0 * 7.5 / 100 - 0.5)
    assert boho == pytest.approx([1.0 + gap, 1.1 + gap, 2.0])
    assert weights == [(4, d[2], 0.6, -0.8)]


def test_new_error_sheets():
    assert new_error_sheets({"A": 5, "B": 0}, {"A": 5, "B": 2, "C": 1}) == ["B", "C"]
