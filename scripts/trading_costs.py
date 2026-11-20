"""Trading-cost test: exchange/clearing fees (and optional slippage) charged on
every change in the position the Knox strategy actually holds, plus rolling
a held position on each contract roll day. Reads workbook copies only.

A full position is 84 soybean oil (ZL) and 16 heating oil (HO) contracts
(Position Model DK8/DK9 at risk level 1); tiers are 25/50/25% of it. D4 has
no exchange fee here.

Usage:
  python scripts/trading_costs.py --workbook <copy.xlsx> [--compare <copy.xlsx>] [--until 2026-11-18] --out <dir>
"""

import argparse
import json
import sys
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent))
from multi_month import MONTHS  # noqa: E402
from phase2.back_adjust import roll_gaps  # noqa: E402
from phase2.compare import summary_stats  # noqa: E402
from phase2.inputs import load_inputs  # noqa: E402

SBO_CONTRACTS, HO_CONTRACTS = 84, 16
FEE_SBO, FEE_HO = 3.66, 3.01          # $ per contract per side (Ryan, 2026-11-20)
TICK_SBO, TICK_HO = 0.0001 * 60000, 0.0001 * 42000  # 0.01 c/lb x 60,000 lb; $0.0001/gal x 42,000 gal
COSTS_PER_SIDE = {
    "fees": SBO_CONTRACTS * FEE_SBO + HO_CONTRACTS * FEE_HO,
    "slippage_1tick": SBO_CONTRACTS * TICK_SBO + HO_CONTRACTS * TICK_HO,
}
PERIODS = {"2021-26": 2021, "2024-26": 2024}


def _num(x) -> float:
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0


def held_tiers(tiers: list, bs: list) -> list:
    """Tier fractions actually held on each day: the previous row's tier
    positions while Effective Trade Active is 1 (the day's PnL comes from the
    position carried in), otherwise flat."""
    out = []
    for i, active in enumerate(bs):
        out.append(tiers[i - 1] if active == 1 and i > 0 else (0, 0, 0))
    return out


def daily_costs(dates: list, held: list, roll_dates: list, per_side: float) -> list:
    """$ cost per day: |change| in every tier traded one side each, plus two
    sides for each held tier on a roll day."""
    rolls = set(roll_dates)
    out, prev = [], (0, 0, 0)
    for d, h in zip(dates, held):
        traded = sum(abs(a - b) for a, b in zip(h, prev))
        if d in rolls:
            traded += 2 * sum(abs(a) for a in h)
        out.append(traded * per_side)
        prev = h
    return out


def read_positions(path: Path, until: date) -> dict:
    cols = {k: column_index_from_string(c) - 1 for k, c in
            {"date": "A", "s": "S", "t": "T", "u": "U", "bs": "BS", "bt": "BT"}.items()}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = {}
        for m in MONTHS:
            rec = {"dates": [], "tiers": [], "bs": [], "bt": []}
            for r in wb[f"Position Model - {m}"].iter_rows(min_row=2, max_col=cols["bt"] + 1, values_only=True):
                d = r[cols["date"]]
                if not hasattr(d, "year") or d.year < 2015:
                    continue
                d = d.date() if hasattr(d, "date") else d
                if d > until:
                    break
                rec["dates"].append(d)
                rec["tiers"].append(tuple(_num(r[cols[k]]) for k in ("s", "t", "u")))
                rec["bs"].append(_num(r[cols["bs"]]))
                rec["bt"].append(_num(r[cols["bt"]]))
            out[m] = rec
        return out
    finally:
        wb.close()


def run(path: Path, until: date) -> dict:
    pos = read_positions(path, until)
    rolls = {m: [d for d, _g in roll_gaps(v)] for m, v in load_inputs(path).items()}
    res = {}
    for m, p in pos.items():
        held = held_tiers(p["tiers"], p["bs"])
        fees = daily_costs(p["dates"], held, [], COSTS_PER_SIDE["fees"])
        roll_fees = [a - b for a, b in zip(daily_costs(p["dates"], held, rolls[m], COSTS_PER_SIDE["fees"]), fees)]
        slip = daily_costs(p["dates"], held, rolls[m], COSTS_PER_SIDE["slippage_1tick"])
        res[m] = {"dates": p["dates"], "gross": p["bt"], "fees": fees, "roll_fees": roll_fees, "slippage": slip,
                  "contract_sides": sum(sum(abs(a - b) for a, b in zip(h, q)) for h, q in zip(held[1:], held[:-1]))}
    return res


def summarize(res: dict) -> dict:
    out = {}
    series = {m: v for m, v in res.items()}
    dates = res[MONTHS[0]]["dates"]
    port = {k: [sum(v[k][i] for v in series.values()) for i in range(len(dates))]
            for k in ("gross", "fees", "roll_fees", "slippage")}
    series["Portfolio"] = {"dates": dates, **port}
    for name, v in series.items():
        out[name] = {}
        for p, y in PERIODS.items():
            idx = [i for i, d in enumerate(v["dates"]) if d.year >= y]
            g = [v["gross"][i] for i in idx]
            f = [v["fees"][i] + v["roll_fees"][i] for i in idx]
            s = [v["slippage"][i] for i in idx]
            out[name][p] = {
                "gross": summary_stats(g),
                "after_fees": summary_stats([a - b for a, b in zip(g, f)]),
                "after_fees_and_slippage": summary_stats([a - b - c for a, b, c in zip(g, f, s)]),
                "fees_trading": sum(v["fees"][i] for i in idx),
                "fees_rolls": sum(v["roll_fees"][i] for i in idx),
                "slippage": sum(s),
            }
    return out


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True, type=Path)
    ap.add_argument("--compare", type=Path)
    ap.add_argument("--until", default=None)
    ap.add_argument("--out", required=True, type=Path)
    a = ap.parse_args(argv)
    until = date.fromisoformat(a.until) if a.until else date(2100, 1, 1)
    a.out.mkdir(parents=True, exist_ok=True)
    for label, path in (("main", a.workbook), ("compare", a.compare)):
        if path:
            s = summarize(run(path, until))
            (a.out / f"costs_{label}.json").write_text(json.dumps(s, indent=1))
            p = s["Portfolio"]["2021-26"]
            print(label, "gross", round(p["gross"]["pnl"]), round(p["gross"]["sharpe"], 3),
                  "| fees", round(p["fees_trading"]), "rolls", round(p["fees_rolls"]),
                  "| after fees", round(p["after_fees"]["pnl"]), round(p["after_fees"]["sharpe"], 3),
                  "| + slippage", round(p["slippage"]), "->", round(p["after_fees_and_slippage"]["pnl"]),
                  round(p["after_fees_and_slippage"]["sharpe"], 3))


if __name__ == "__main__":
    main()
