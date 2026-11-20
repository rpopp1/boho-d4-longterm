"""Baseline vs back-adjusted comparison for the Phase 2 study. Reads two
Knox workbooks' saved values (Position Model tabs + Johansen_Monthly) and
writes JSON tables for the results page.

Usage:
  python scripts/phase2/compare.py --base <baseline.xlsx> --test <TEST.xlsx> --out <dir>
"""

import argparse
import json
import math
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from multi_month import MONTHS  # noqa: E402
from phase2.back_adjust import roll_gaps  # noqa: E402
from phase2.inputs import load_inputs  # noqa: E402

_COLS = {"date": "A", "boho": "D", "z": "L", "v": "V", "bs": "BS", "bt": "BT", "bu": "BU",
         "regime": "DC", "en": "EN"}
POST_ROLL_DAYS = 21
PERIODS = {"2021-26": 2021, "2024-26": 2024}


def _num(x) -> float:
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0


def _read_tab(rows) -> dict:
    idx = {k: column_index_from_string(c) - 1 for k, c in _COLS.items()}
    out = {k: [] for k in _COLS}
    for r in rows:
        d = r[idx["date"]]
        if not isinstance(d, (datetime, date)) or d.year < 2015:
            continue
        for k, i in idx.items():
            out[k].append(r[i] if i < len(r) else None)
    out["dates"] = [d.date() if isinstance(d, datetime) else d for d in out.pop("date")]
    for k in ("bt", "bu", "v", "bs", "en"):
        out[k] = [_num(x) for x in out[k]]
    return out


def read_workbook(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        last = column_index_from_string("EN")
        return {m: _read_tab(wb[f"Position Model - {m}"].iter_rows(min_row=2, max_col=last, values_only=True))
                for m in MONTHS}
    finally:
        wb.close()


def read_month(path: Path, month: str) -> dict:
    return read_workbook(path)[month]


def trades(dates, bs, v, bu, en) -> list:
    """Runs of Effective Trade Active (BS = 1). Reason, from the day after the
    run: profit-take if Exit Was Profit-Take is set, stop if the signal is
    still on, otherwise the signal closed it."""
    out, i, n = [], 0, len(dates)
    while i < n:
        if bs[i] != 1:
            i += 1
            continue
        j = i
        while j + 1 < n and bs[j + 1] == 1:
            j += 1
        after = j + 1 if j + 1 < n else j
        if en[j] == 1 or en[after] == 1:
            reason = "profit-take"
        elif after != j and v[after] != 0:
            reason = "stop"
        else:
            reason = "signal"
        out.append({"start": dates[i], "end": dates[j], "side": 1 if v[i] > 0 else -1,
                    "pnl": bu[j], "reason": reason})
        i = j + 1
    return out


def match_trades(base: list, test: list, dates: list, shift_days: int = 3) -> dict:
    pos = {d: k for k, d in enumerate(dates)}
    res = {"same": [], "changed_pnl": [], "shifted": [], "added": [], "removed": []}
    left = list(test)
    unmatched = []
    for b in base:
        hit = next((t for t in left if t["start"] == b["start"] and t["side"] == b["side"]), None)
        if hit is None:
            unmatched.append(b)
            continue
        left.remove(hit)
        res["same" if abs(hit["pnl"] - b["pnl"]) < 0.5 else "changed_pnl"].append((b, hit))
    for b in unmatched:
        near = [t for t in left if t["side"] == b["side"] and abs(pos[t["start"]] - pos[b["start"]]) <= shift_days]
        if near:
            hit = min(near, key=lambda t: abs(pos[t["start"]] - pos[b["start"]]))
            left.remove(hit)
            res["shifted"].append((b, hit))
        else:
            res["removed"].append(b)
    res["added"] = left
    return res


def summary_stats(bt: list) -> dict:
    n = len(bt)
    if not n:
        return {"pnl": 0.0, "sharpe": 0.0, "sortino": 0.0, "maxdd": 0.0}
    mean = sum(bt) / n
    sd = math.sqrt(sum((x - mean) ** 2 for x in bt) / n)
    down = math.sqrt(sum(x * x for x in bt if x < 0) / n)
    cum, peak, maxdd = 0.0, None, 0.0
    for x in bt:
        cum += x
        peak = cum if peak is None else max(peak, cum)
        maxdd = min(maxdd, cum - peak)
    return {"pnl": sum(bt), "sharpe": mean / sd * math.sqrt(252) if sd else 0.0,
            "sortino": mean / down * math.sqrt(252) if down else 0.0, "maxdd": maxdd}


def _period(month_data: dict, first_year: int) -> list:
    return [x for d, x in zip(month_data["dates"], month_data["bt"]) if d.year >= first_year]


def compare(base_path: Path, test_path: Path, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    base, test = read_workbook(base_path), read_workbook(test_path)
    bin_, tin = load_inputs(base_path), load_inputs(test_path)

    weights, daily, trade_tab, summary = [], {}, {}, {}
    for m in MONTHS:
        rolls = [d for d, _g in roll_gaps(bin_[m])]
        for (end, d, bw, bd), (_e, _d, tw, td) in zip(bin_[m].stored, tin[m].stored):
            prior = [r for r in rolls if r <= d]
            since = (d.year - prior[-1].year) * 12 + d.month - prior[-1].month if prior else None
            weights.append({"month": m, "date": d.isoformat(), "base_boho": bw, "test_boho": tw,
                            "base_d4": bd, "test_d4": td, "diff_boho": tw - bw, "months_since_roll": since})

        b, t = base[m], test[m]
        assert b["dates"] == t["dates"], m
        zd = [abs(x - y) for x, y in zip(b["z"], t["z"]) if isinstance(x, (int, float)) and isinstance(y, (int, float))]
        pos = {d: k for k, d in enumerate(b["dates"])}
        post = set()
        for r in rolls:
            if r in pos:
                post.update(range(pos[r], min(pos[r] + POST_ROLL_DAYS, len(b["dates"]))))
        zpost = [abs(b["z"][k] - t["z"][k]) for k in post
                 if isinstance(b["z"][k], (int, float)) and isinstance(t["z"][k], (int, float))]
        daily[m] = {
            "position_diff_days": sum(1 for x, y in zip(b["v"], t["v"]) if x != y),
            "regime_diff_days": sum(1 for x, y in zip(b["regime"], t["regime"]) if x != y),
            "z_abs_diff_mean": sum(zd) / len(zd) if zd else 0.0,
            "z_abs_diff_max": max(zd) if zd else 0.0,
            "z_abs_diff_mean_post_roll": sum(zpost) / len(zpost) if zpost else 0.0,
            "days": len(b["dates"]),
        }
        bt_, tt_ = (trades(x["dates"], x["bs"], x["v"], x["bu"], x["en"]) for x in (b, t))
        mt = match_trades(bt_, tt_, b["dates"])
        trade_tab[m] = {k: [_jsonable(x) for x in v] for k, v in mt.items()}
        trade_tab[m]["counts"] = {k: len(v) for k, v in mt.items()} | {"base": len(bt_), "test": len(tt_)}
        summary[m] = {p: {"base": summary_stats(_period(b, y)), "test": summary_stats(_period(t, y)),
                          "base_trades": sum(1 for x in bt_ if x["start"].year >= y),
                          "test_trades": sum(1 for x in tt_ if x["start"].year >= y)}
                      for p, y in PERIODS.items()}

    port = {}
    for p, y in PERIODS.items():
        bsum = [sum(v) for v in zip(*(_period(base[m], y) for m in MONTHS))]
        tsum = [sum(v) for v in zip(*(_period(test[m], y) for m in MONTHS))]
        port[p] = {"base": summary_stats(bsum), "test": summary_stats(tsum),
                   "base_trades": sum(summary[m][p]["base_trades"] for m in MONTHS),
                   "test_trades": sum(summary[m][p]["test_trades"] for m in MONTHS)}
    summary["Portfolio"] = port

    tables = {"weights": weights, "daily_diff": daily, "trades": trade_tab, "summary": summary}
    for name, obj in tables.items():
        (out_dir / f"{name}.json").write_text(json.dumps(obj, indent=1, default=str))
    return tables


def _jsonable(x):
    if isinstance(x, tuple):
        return [_jsonable(y) for y in x]
    return {k: (v.isoformat() if isinstance(v, date) else v) for k, v in x.items()}


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--base", required=True, type=Path)
    ap.add_argument("--test", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    a = ap.parse_args(argv)
    t = compare(a.base, a.test, a.out)
    for k, v in t["summary"]["Portfolio"].items():
        print(k, "base", {x: round(y, 2) for x, y in v["base"].items()}, "test", {x: round(y, 2) for x, y in v["test"].items()})


if __name__ == "__main__":
    main()
