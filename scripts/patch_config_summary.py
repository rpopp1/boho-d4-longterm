"""Makes the Config sheet's 2024-26 and since-inception summary rows extend
themselves as rows are pasted, instead of ending at a fixed row.

Rows 34/35 (PnL), 38/39 (Sharpe) and 44/45 (2024-26 max equity / max
drawdown), all 8 month columns. Each formula finds its month's last dated
Position Model row (the same LOOKUP(2,1/(date>DATE(2015,1,1))) test the
Dashboard uses). The closed 2021-23 rows (33, 37, 41, 42) are left alone.

Usage:
  python scripts/patch_config_summary.py <knox workbook.xlsx> [--dry-run]
A timestamped backup is written next to the workbook first (not in --dry-run).
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config_map import CONFIG_MONTH_COLUMNS  # noqa: E402
from excel_com import error_cell_count, isolated_workbook, reset_calc_mode_to_automatic, retry, sheet_names  # noqa: E402
from multi_month import MONTHS  # noqa: E402

PADDED_LAST_ROW = 3001
FIRST_ROW_ALL, FIRST_ROW_2024 = 29, 741  # 2/11/2021 (after warm-up) and 1/2/2024, as the existing rows use


def _pm(month: str) -> str:
    return f"'Position Model - {month}'"


def last_row_expr(month: str) -> str:
    dates = f"{_pm(month)}!$A$2:$A${PADDED_LAST_ROW}"
    return f"MATCH(LOOKUP(2,1/({dates}>DATE(2015,1,1)),{dates}),{dates},0)+1"


def _rng(month: str, col: str, first: int) -> str:
    return f"{_pm(month)}!${col}${first}:INDEX({_pm(month)}!${col}$1:${col}${PADDED_LAST_ROW},last)"


def formulas(month: str) -> dict:
    """{config row: formula} for `month`."""
    lr = last_row_expr(month)
    sharpe = "(AVERAGE(r)/STDEV.S(r))*SQRT(252)"
    return {
        34: f"=LET(last,{lr},SUM({_rng(month, 'BT', FIRST_ROW_2024)}))",
        35: f"=LET(last,{lr},SUM({_rng(month, 'BT', FIRST_ROW_ALL)}))",
        38: f"=LET(last,{lr},r,{_rng(month, 'BT', FIRST_ROW_2024)},{sharpe})",
        39: f"=LET(last,{lr},r,{_rng(month, 'BT', FIRST_ROW_ALL)},{sharpe})",
        44: f"=LET(last,{lr},INDEX({_pm(month)}!$CA$1:$CA${PADDED_LAST_ROW},last))",
        45: f"=LET(last,{lr},INDEX({_pm(month)}!$CC$1:$CC${PADDED_LAST_ROW},last))",
    }


def apply(path: Path) -> dict:
    with isolated_workbook(path, manual_calc=True) as (excel, wb):
        excel.CalculateFullRebuild()
        before = {n: error_cell_count(wb.Worksheets(n)) for n in sheet_names(wb)}
        cfg = wb.Worksheets("Config")
        for month in MONTHS:
            col = CONFIG_MONTH_COLUMNS[month]
            for row, f in formulas(month).items():
                cfg.Range(f"{col}{row}").Formula2 = f
        excel.CalculateFullRebuild()
        after = {n: error_cell_count(wb.Worksheets(n)) for n in sheet_names(wb)}
        values = {m: {r: cfg.Range(f"{CONFIG_MONTH_COLUMNS[m]}{r}").Value for r in (34, 35, 38, 39, 44, 45)}
                  for m in MONTHS}
        if after != before:
            raise RuntimeError(f"Error-cell counts changed: before {before}, after {after}")
        wb.Save()
    reset_calc_mode_to_automatic(path)
    return values


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--dry-run", action="store_true", help="the workbook is a scratch copy; skip the backup")
    a = ap.parse_args(argv)
    path = a.workbook.resolve()
    if not a.dry_run:
        backup = path.with_name(path.stem + f" - backup before config summary {datetime.now():%Y%m%d-%H%M%S}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup written: {backup}")
    for m, v in retry(lambda: apply(path)).items():
        print(m, {r: round(x, 4) if isinstance(x, float) else x for r, x in v.items()})


if __name__ == "__main__":
    main()
