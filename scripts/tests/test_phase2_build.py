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
    with pytest.raises(ValueError):
        output_name(tmp_path, "other")


def test_new_error_sheets():
    assert new_error_sheets({"A": 5, "B": 0}, {"A": 5, "B": 2, "C": 1}) == ["B", "C"]
