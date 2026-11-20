"""Builds a NEW Knox workbook for the Phase 2 study from a scratch copy of
the Sandbox:

  dry       original BOHO prices + the stored weights written back through the
            same path (proves the write path changes nothing by itself)
  adjusted  back-adjusted first-block BOHO prices + Johansen weights
            recomputed on them
  jump_only back-adjusted prices + the stored weights kept (isolates the
            roll jump in the Z-score and regime from the weight changes)

Only Data's first-block BOHO column and Johansen_Monthly's stored weight
columns are written. The output must be a new file; Ryan's workbooks (every
.xlsx in the folder without "Phase 2" in its name) are refused.

Usage:
  python scripts/phase2/build_test_workbook.py --source <scratch copy.xlsx> --mode dry|adjusted|jump_only
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from excel_com import error_cell_count, isolated_workbook, reset_calc_mode_to_automatic, retry, sheet_names  # noqa: E402
from phase2.back_adjust import back_adjust, roll_gaps  # noqa: E402
from phase2.inputs import JM_FIRST_ROW, data_boho_col, johansen_cols, load_inputs  # noqa: E402
from phase2.johansen import gate_report, monthly_weights  # noqa: E402

FOLDER = Path(__file__).resolve().parents[2]
STEM = "Knox_Multimonth_BOHO_D4_Strategy - Phase 2"
OUTPUT_NAMES = {"adjusted": f"{STEM} Back-Adjusted TEST.xlsx", "dry": f"{STEM} Baseline DRY RUN.xlsx",
                "jump_only": f"{STEM} Jump-Only TEST.xlsx"}


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


def new_error_sheets(base: dict, test: dict) -> list:
    return sorted(name for name, n in test.items() if n > base.get(name, 0))


def _error_counts(excel, wb) -> dict:
    excel.CalculateFullRebuild()
    return {name: error_cell_count(wb.Worksheets(name)) for name in sheet_names(wb)}


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
    check_output_path(out, protected_paths(out.parent))
    writes, manifest = plan_writes(source, mode)

    with isolated_workbook(source, manual_calc=True) as (excel, wb):
        base_errors = retry(lambda: _error_counts(excel, wb))
        wb.Saved = True

    shutil.copy2(source, out)
    with isolated_workbook(out, manual_calc=True) as (excel, wb):
        data, jm = wb.Worksheets("Data"), wb.Worksheets("Johansen_Monthly")
        for month, (rows, boho, weights) in writes.items():
            if rows != list(range(rows[0], rows[0] + len(rows))):
                raise RuntimeError(f"{month}: Data rows are not contiguous")
            col = data_boho_col(month)
            data.Range(f"{col}{rows[0]}:{col}{rows[-1]}").Value = [[v] for v in boho]
            jc = johansen_cols(month)
            last = JM_FIRST_ROW + len(weights) - 1
            jm.Range(f"{jc['w_boho']}{JM_FIRST_ROW}:{jc['w_d4']}{last}").Value = \
                [[wb_, wd] for _e, _d, wb_, wd in weights]
        test_errors = _error_counts(excel, wb)
        bad = new_error_sheets(base_errors, test_errors)
        if bad:
            raise RuntimeError(f"New error cells after recalculation in: {bad} "
                               f"(base {[base_errors.get(b) for b in bad]}, test {[test_errors[b] for b in bad]})")
        wb.Save()
    reset_calc_mode_to_automatic(out)
    manifest["errors"] = {"base": base_errors, "test": test_errors}
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
