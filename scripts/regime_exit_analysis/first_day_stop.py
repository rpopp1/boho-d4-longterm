# First-day stop test: let the trade-level stop fire on a trade's first day.
# Usage: python first_day_stop.py <workbook copy.xlsx> <out.json>
# Current rules ($100k profit-take + profit-only month-start re-entry) vs the same with
# first_day_stop='all' and 'fresh'. The baseline must reproduce the workbook's BT exactly.
import json
import sys

import numpy as np
import pandas as pd

import exit_backtest as eb

OUT = sys.argv[2]
PT = 100000
RUNS = {"current": None, "day1_all": "all", "day1_fresh": "fresh"}


def stats(s):
    s = pd.Series(s)
    cum = s.cumsum()
    sd = s.std(ddof=0)
    return {"pnl": float(s.sum()), "sharpe": float(s.mean() / sd * np.sqrt(252)) if sd else 0.0,
            "maxdd": float((cum - cum.cummax()).min())}


res, series, changed = {}, {k: [] for k in RUNS}, {}
for m in eb.MONTHS:
    df = eb.load(m)
    out = {}
    for name, mode in RUNS.items():
        eff, tr = eb.simulate(df, stop=eb.STOP_BASE[m], profit=PT, reentry_month="profit_only", first_day_stop=mode)
        out[name] = (eff, tr)
        series[name].append(pd.Series(eff, index=df.Date))
    diff = np.abs(out["current"][0] - df.BT.values).max()
    assert diff < 1e-6, (m, diff)
    res[m] = {}
    for name in RUNS:
        eff, tr = out[name]
        dates = df.Date
        res[m][name] = {"2021-26": stats(eff), "2024-26": stats(eff[dates.dt.year.values >= 2024]),
                        "trades": int(len(tr)), "stops": int((tr.reason == "stop").sum())}
    # trades whose outcome differs (keyed by start date)
    base = out["current"][1].set_index("start")
    for name in ("day1_all", "day1_fresh"):
        t = out[name][1].set_index("start")
        rows = []
        for d in base.index.intersection(t.index):
            a, b = base.loc[d], t.loc[d]
            a = a.iloc[0] if isinstance(a, pd.DataFrame) else a
            b = b.iloc[0] if isinstance(b, pd.DataFrame) else b
            if abs(a.pnl - b.pnl) > 0.5 or a.reason != b.reason:
                rows.append({"start": str(d.date()), "pnl_current": round(float(a.pnl)), "pnl_new": round(float(b.pnl)),
                             "reason_current": a.reason, "reason_new": b.reason})
        changed.setdefault(name, {})[m] = rows
    print(m, "validation max diff", diff, {k: (round(v["2021-26"]["pnl"]), round(v["2021-26"]["sharpe"], 3)) for k, v in res[m].items()})

port = {}
for name in RUNS:
    s = pd.concat(series[name], axis=1).fillna(0).sum(axis=1)
    port[name] = {"2021-26": stats(s.values), "2024-26": stats(s[s.index.year >= 2024].values)}
res["Portfolio"] = port
json.dump({"results": res, "changed": changed}, open(OUT, "w"), indent=1)
for name, v in port.items():
    print("Portfolio", name, {p: {k: round(x, 3) for k, x in d.items()} for p, d in v.items()})
