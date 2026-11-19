"""Adds the profit-take and month-start re-entry overlay (profit_take.py) to an
existing Knox multi-month workbook.

1. Labels Config rows 30 (ProfitTakeUSD) and 31 (ReentryMonthStart) and sets
   them for the months given on the command line (others stay off).
2. In each Position Model tab, checks every BO/BP/BQ/BT/BU cell from row 21 to
   the last data row still has its original stop-overlay formula and EK:EO
   are empty, so nothing hand-edited is overwritten, then writes the overlay.
3. Fully recalculates in an isolated Excel process, checks for error cells and
   saves.

Usage:
  python patch_knox_profit_take.py <knox workbook.xlsx> [--profit-take 100000]
      [--reentry] [--months December,January | all]

With no --profit-take/--reentry the overlay is installed but switched off:
every cell computes exactly as before. Run it on a copy; it edits in place.
"""

import argparse
from pathlib import Path

from config_map import CONFIG_MONTH_COLUMNS
from excel_com import error_cell_count, isolated_workbook, reset_calc_mode_to_automatic, retry, sheet_names
from multi_month import MONTHS
from profit_take import FIRST_ROW, HEADERS, HELPER_COLS, REWRITTEN_COLS, original_formula, overlay_formulas

PADDED_LAST_ROW = 3001
CONFIG_LABELS = {30: "ProfitTakeUSD", 31: "ReentryMonthStart"}


def parse_months(text: str) -> list:
    if text.strip().lower() == "all":
        return list(MONTHS)
    months = [m.strip().capitalize() for m in text.split(",") if m.strip()]
    unknown = [m for m in months if m not in MONTHS]
    if unknown:
        raise SystemExit(f"Unknown month(s): {unknown}. Choose from {MONTHS} or 'all'.")
    return months


def check_config(ws) -> None:
    for row, label in CONFIG_LABELS.items():
        current = ws.Range(f"A{row}").Value
        if current not in (None, "", label):
            raise RuntimeError(f"Config!A{row} holds {current!r}; expected it empty or {label!r}.")


def write_config(ws, profit_take, reentry: bool, months: list) -> None:
    for row, label in CONFIG_LABELS.items():
        ws.Range(f"A{row}").Value = label
    for month in MONTHS:
        col = CONFIG_MONTH_COLUMNS[month]
        on = month in months
        ws.Range(f"{col}30").Value = profit_take if (on and profit_take) else ""
        ws.Range(f"{col}31").Value = 1 if (on and reentry) else 0


def check_original_shape(ws, month: str) -> int:
    """Return the last row of the stop overlay, after checking every
    rewritten cell from FIRST_ROW to it still holds its original formula and
    the helper columns are empty. Raises otherwise."""
    bo = ws.Range(f"BO{FIRST_ROW}:BO{PADDED_LAST_ROW}").Formula
    last = max((FIRST_ROW + i for i, (v,) in enumerate(bo) if v), default=FIRST_ROW - 1)
    if last < FIRST_ROW:
        raise RuntimeError(f"Position Model - {month}: no stop-overlay formulas found in column BO.")
    bad = []
    for col in REWRITTEN_COLS:
        vals = ws.Range(f"{col}{FIRST_ROW}:{col}{PADDED_LAST_ROW}").Formula
        for i, (actual,) in enumerate(vals):
            r = FIRST_ROW + i
            if r > last:
                if actual:
                    bad.append(f"{col}{r}: {actual!r} (below the last row)")
                continue
            allowed = {original_formula(col, month, r), original_formula(col, month, r, absolute_stop=False)}
            if actual not in allowed:
                bad.append(f"{col}{r}: {actual!r}")
    for col in HELPER_COLS:
        vals = ws.Range(f"{col}1:{col}{PADDED_LAST_ROW}").Formula
        bad += [f"{col}{i + 1}: {v!r} (should be empty)" for i, (v,) in enumerate(vals) if v]
    if bad:
        raise RuntimeError(f"Position Model - {month}: {len(bad)} unexpected cells, e.g. {bad[:3]}")
    return last


def patch_position_model(ws, month: str) -> int:
    last = check_original_shape(ws, month)
    for col, header in HEADERS.items():
        ws.Range(f"{col}1").Value = header
    # One formula string per column range: Excel shifts the relative row
    # references down the range; the Config references are absolute.
    for col, formula in overlay_formulas(month, FIRST_ROW).items():
        ws.Range(f"{col}{FIRST_ROW}:{col}{last}").Formula2 = formula
    return last


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--profit-take", type=float, default=None, help="trade-level profit-take in $, e.g. 100000")
    ap.add_argument("--reentry", action="store_true", help="turn on the profit-only month-start re-entry")
    ap.add_argument("--months", default="all", help="comma-separated months to switch on, or 'all' (default)")
    args = ap.parse_args(argv)
    path = Path(args.workbook).resolve()
    if not path.exists():
        raise SystemExit(f"Workbook not found: {path}")
    months = parse_months(args.months)

    with isolated_workbook(path, manual_calc=True) as (excel, wb):
        config = wb.Worksheets("Config")
        retry(lambda: check_config(config))
        lasts = {}
        for month in MONTHS:  # check every tab before writing anything
            ws = wb.Worksheets(f"Position Model - {month}")
            lasts[month] = retry(lambda ws=ws, m=month: check_original_shape(ws, m))
        retry(lambda: write_config(config, args.profit_take, args.reentry, months))
        for month in MONTHS:
            ws = wb.Worksheets(f"Position Model - {month}")
            last = retry(lambda ws=ws, m=month: patch_position_model(ws, m))
            print(f"Position Model - {month}: overlay written for rows {FIRST_ROW}-{last}")
        excel.CalculateFullRebuild()
        errors = {name: error_cell_count(wb.Worksheets(name)) for name in sheet_names(wb)
                  if name.startswith("Position Model") or name.startswith("Dashboard")}
        print("error cells:", errors)
        wb.Save()
    reset_calc_mode_to_automatic(path)
    state = (f"profit-take {args.profit_take:,.0f}" if args.profit_take else "profit-take off") + \
            (", re-entry on" if args.reentry else ", re-entry off")
    print(f"Overlay applied to {path} ({state} for {', '.join(months)})")


if __name__ == "__main__":
    main()
