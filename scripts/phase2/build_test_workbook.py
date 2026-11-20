"""Builds a NEW Knox workbook for the Phase 2 study from a scratch copy of
the Sandbox:

  dry       original BOHO prices + the stored weights written back through the
            same path (proves the write path changes nothing by itself)
  adjusted  back-adjusted first-block BOHO prices + Johansen weights
            recomputed on them
  jump_only back-adjusted prices + the stored weights kept
  local     the roll gap removed only inside windows that contain a roll:
            Z-score windows (Position Model L values), Johansen windows
            (weights) and the roll day's BOHO change (CV); all else identical

Only Data's first-block BOHO column and Johansen_Monthly's stored weight
columns are written. The output must be a new file; Ryan's workbooks (every
.xlsx in the folder without "Phase 2" in its name) are refused.

Usage:
  python scripts/phase2/build_test_workbook.py --source <scratch copy.xlsx> --mode dry|adjusted|jump_only|local
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl.utils.cell import get_column_letter, range_boundaries  # noqa: E402

from excel_com import isolated_workbook, reset_calc_mode_to_automatic, retry, sheet_names  # noqa: E402
from phase2.back_adjust import back_adjust, roll_gaps  # noqa: E402
from phase2.inputs import JM_FIRST_ROW, data_boho_col, johansen_cols, load_inputs  # noqa: E402
from phase2.johansen import gate_report, monthly_weights  # noqa: E402
from phase2.local import load_signal_inputs, local_weights, roll_day_changes, rows_in_roll_windows, z_series  # noqa: E402

FOLDER = Path(__file__).resolve().parents[2]
STEM = "Knox_Multimonth_BOHO_D4_Strategy - Phase 2"
OUTPUT_NAMES = {"adjusted": f"{STEM} Back-Adjusted TEST.xlsx", "dry": f"{STEM} Baseline DRY RUN.xlsx",
                "jump_only": f"{STEM} Jump-Only TEST.xlsx", "local": f"{STEM} Window-Local TEST.xlsx"}


def output_name(folder: Path, mode: str) -> Path:
    if mode not in OUTPUT_NAMES:
        raise ValueError(f"mode must be one of {sorted(OUTPUT_NAMES)}, got {mode!r}")
    return folder / OUTPUT_NAMES[mode]


def protected_paths(folder: Path) -> list:
    return [p for p in folder.glob("*.xlsx") if "Phase 2" not in p.name]


def check_output_path(out: Path, protected: list) -> None:
    target = out.resolve()
    if any(target == p.resolve() for p in protected):
        raise ValueError(f"Refusing to write to a protected workbook: {out}")
    if out.exists():
        raise ValueError(f"Refusing to overwrite an existing file: {out}")


def check_source(source: Path, protected: list) -> None:
    """The source must be a scratch copy, never one of Ryan's workbooks:
    Excel opens it to count error cells."""
    if any(source.resolve() == p.resolve() for p in protected):
        raise ValueError(f"Refusing to use a protected workbook as the source; pass a copy: {source}")


def new_error_sheets(base: dict, test: dict) -> list:
    return sorted(name for name, n in test.items() if n > base.get(name, 0))


def expand_addresses(address: str) -> set:
    """'$L$2:$L$4,$M$9' -> {'L2','L3','L4','M9'}."""
    cells = set()
    for part in address.replace("$", "").split(","):
        c1, r1, c2, r2 = range_boundaries(part if ":" in part else f"{part}:{part}")
        cells |= {f"{get_column_letter(c)}{r}" for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}
    return cells


def new_error_cells(base: dict, test: dict) -> dict:
    """{sheet: sorted cells that hold an error in `test` but not in `base`}."""
    out = {}
    for name, cells in test.items():
        extra = sorted(cells - base.get(name, set()))
        if extra:
            out[name] = extra
    return out


def _error_cells(excel, wb) -> dict:
    excel.CalculateFullRebuild()
    out = {}
    for name in sheet_names(wb):
        try:
            rng = wb.Worksheets(name).UsedRange.SpecialCells(-4123, 16)  # formulas, errors
            out[name] = set().union(*(expand_addresses(a.Address) for a in rng.Areas))
        except Exception:  # SpecialCells raises when nothing matches
            out[name] = set()
    return out


def plan_local(signals: dict) -> dict:
    """{month: {"weights", "z": {row: Z}, "cv": {row: change}, "gaps"}} for the
    window-local test: only values inside windows that hold a roll change."""
    plan = {}
    for month, s in signals.items():
        gaps = roll_gaps(s.m)
        weights = local_weights(s.m, gaps)
        z = z_series(s.m, s.windows, gaps, weights)
        keep = set(rows_in_roll_windows(s.m, s.windows, gaps))
        plan[month] = {"weights": weights, "gaps": gaps, "cv": roll_day_changes(s.m, gaps),
                       "z": {r: v for r, v in zip(s.m.rows, z) if r in keep and v is not None}}
    return plan


def _runs(rows: list) -> list:
    """Consecutive runs of row numbers: [(first, last), ...]."""
    out = []
    for r in sorted(rows):
        if out and r == out[-1][1] + 1:
            out[-1][1] = r
        else:
            out.append([r, r])
    return [tuple(x) for x in out]


def _write_local(wb, plan: dict) -> None:
    jm = wb.Worksheets("Johansen_Monthly")
    for month, p in plan.items():
        jc = johansen_cols(month)
        last = JM_FIRST_ROW + len(p["weights"]) - 1
        jm.Range(f"{jc['w_boho']}{JM_FIRST_ROW}:{jc['w_d4']}{last}").Value = \
            [[wb_, wd] for _e, _d, wb_, wd in p["weights"]]
        pm = wb.Worksheets(f"Position Model - {month}")
        for a, b in _runs(list(p["z"])):
            pm.Range(f"L{a}:L{b}").Value = [[p["z"][r]] for r in range(a, b + 1)]
        for r, v in p["cv"].items():
            pm.Range(f"CV{r}").Value = v


def plan_writes(source: Path, mode: str) -> tuple:
    """{month: (boho values, weight rows)} plus the manifest."""
    inputs = load_inputs(source)
    writes, manifest = {}, {"mode": mode, "months": {}}
    for month, m in inputs.items():
        gaps = roll_gaps(m) if mode in ("adjusted", "jump_only") else []
        boho = back_adjust(m.dates, m.boho, gaps) if gaps else list(m.boho)
        weights = (monthly_weights(m, boho) if mode == "adjusted"
                   else [(e, d, wb, wd) for e, d, wb, wd in m.stored])
        writes[month] = (m.rows, boho, weights)
        manifest["months"][month] = {
            "gaps": [(d.isoformat(), g) for d, g in gaps],
            "gate": {k: (str(v) if k == "worst" else v) for k, v in gate_report(m).items()},
        }
    return writes, manifest


def build(source: Path, out: Path, mode: str) -> dict:
    check_source(source, protected_paths(FOLDER))
    check_output_path(out, protected_paths(out.parent))
    plan, writes = None, {}
    if mode == "local":
        plan = plan_local(load_signal_inputs(source))
        stored = {m: v.stored for m, v in load_inputs(source).items()}
        manifest = {"mode": mode, "months": {
            m: {"gaps": [(d.isoformat(), g) for d, g in p["gaps"]], "z_rows_rewritten": len(p["z"]),
                "weights_changed": sum(1 for a, b in zip(p["weights"], stored[m]) if a != b)}
            for m, p in plan.items()}}
    else:
        writes, manifest = plan_writes(source, mode)

    with isolated_workbook(source, manual_calc=True) as (excel, wb):
        base_cells = retry(lambda: _error_cells(excel, wb))
        wb.Saved = True

    shutil.copy2(source, out)
    with isolated_workbook(out, manual_calc=True) as (excel, wb):
        data, jm = wb.Worksheets("Data"), wb.Worksheets("Johansen_Monthly")
        if plan:
            _write_local(wb, plan)
        for month, (rows, boho, weights) in writes.items():
            if rows != list(range(rows[0], rows[0] + len(rows))):
                raise RuntimeError(f"{month}: Data rows are not contiguous")
            col = data_boho_col(month)
            data.Range(f"{col}{rows[0]}:{col}{rows[-1]}").Value = [[v] for v in boho]
            jc = johansen_cols(month)
            last = JM_FIRST_ROW + len(weights) - 1
            jm.Range(f"{jc['w_boho']}{JM_FIRST_ROW}:{jc['w_d4']}{last}").Value = \
                [[wb_, wd] for _e, _d, wb_, wd in weights]
        test_cells = _error_cells(excel, wb)
        bad = new_error_cells(base_cells, test_cells)
        if bad:
            raise RuntimeError("New error cells after recalculation: "
                               + str({k: v[:5] for k, v in bad.items()}))
        wb.Save()
    reset_calc_mode_to_automatic(out)
    manifest["errors"] = {
        "base": {k: len(v) for k, v in base_cells.items()},
        "test": {k: len(v) for k, v in test_cells.items()},
        "cleared": {k: sorted(base_cells[k] - test_cells.get(k, set()))[:20] for k in base_cells
                    if base_cells[k] - test_cells.get(k, set())},
    }
    out.with_suffix(".manifest.json").write_text(json.dumps(manifest, indent=1, default=str))
    return manifest


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", required=True, type=Path)
    ap.add_argument("--mode", required=True, choices=sorted(OUTPUT_NAMES))
    args = ap.parse_args(argv)
    out = output_name(FOLDER, args.mode)
    manifest = build(args.source.resolve(), out, args.mode)
    print(f"Wrote {out}")
    print("error cells (base -> test):", {k: (manifest['errors']['base'][k], v)
                                          for k, v in manifest["errors"]["test"].items() if v})


if __name__ == "__main__":
    main()
