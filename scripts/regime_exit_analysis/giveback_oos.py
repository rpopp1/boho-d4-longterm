# Out-of-sample test of a trailing give-back stop on the Knox Position Model tabs.
# Usage: python giveback_oos.py <exit_backtest.py> <workbook.xlsx> <outdir>
# Rule: once a trade's cumulative PnL has reached a peak >= floor, exit if it retraces
# `frac` of that peak (e.g. 0.4 = give back 40% of the peak gain). A naive (frac, floor=0)
# version was tested first and lost to the current rule everywhere (noise-level peaks of a
# few thousand dollars trivially "give back" 40-80%); the floor requires a real gain before
# the trailing stop arms, which is what made the vol-scaled stop's and profit-take's levels
# meaningful too.
# Same OOS method as tp_oos.py: choose (frac, floor) on 2021-2023, score on 2024-2026,
# plus a year-by-year walk-forward. Compared against both the current stop and the
# $100k December profit-take / 1.0x all-month profit-take picked in tp_oos.py.
import json, runpy, sys
from pathlib import Path
import numpy as np
import pandas as pd

lib, src, outdir = sys.argv[1], sys.argv[2], Path(sys.argv[3])
sys.argv = [lib, src]
g = runpy.run_path(lib, run_name="lib")
load, simulate, SB, MONTHS = g["load"], g["simulate"], g["STOP_BASE"], g["MONTHS"]

FLOORS = [50000, 75000, 100000, 150000]
FRACS = [0.3, 0.4, 0.5, 0.6, 0.7]
GB = [None] + [(f, fl) for fl in FLOORS for f in FRACS]
YEARS = [2021, 2022, 2023, 2024, 2025, 2026]
IS, OOS = [2021, 2022, 2023], [2024, 2025, 2026]
TP_DEC, TP_K = 100000, 1.0  # fixed reference profit-take levels from tp_oos.py


def sharpe(s):
    return float(s.mean() / s.std() * np.sqrt(252)) if s.std() > 0 else 0.0


def maxdd(s):
    c = s.cumsum()
    return float((c - c.cummax()).min())


def score(s):
    return {"total": float(s.sum()), "sharpe": sharpe(s), "max_dd": maxdd(s)}


def window(s, years):
    return s[s.index.year.isin(years)]


def label(g_):
    return "None" if g_ is None else f"{g_[0]:.0%} gb / ${g_[1]//1000}k floor"


data = {m: load(m) for m in MONTHS}
paths_gb, paths_tp = {}, {}
for m in MONTHS:
    for gv in GB:
        eff, _ = simulate(data[m], stop=SB[m], giveback=gv)
        paths_gb[(m, gv)] = pd.Series(eff, index=data[m].Date)
    eff, _ = simulate(data[m], stop=SB[m], profit=TP_DEC if m == "December" else TP_K * -SB[m])
    paths_tp[m] = pd.Series(eff, index=data[m].Date)


def pooled_gb(gv, years):
    return pd.concat([window(paths_gb[(m, gv)], years) for m in MONTHS], axis=1).fillna(0).sum(axis=1)


def pooled_tp(years):
    return pd.concat([window(paths_tp[m], years) for m in MONTHS], axis=1).fillna(0).sum(axis=1)


grid, picks, wf, curves = [], [], [], []

# --- December alone ---
for gv in GB:
    s = paths_gb[("December", gv)]
    for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
        grid.append({"test": "dec", "level": label(gv), "frac": None if gv is None else gv[0],
                      "floor": None if gv is None else gv[1], "period": per, **score(window(s, yrs))})
dec_best = max(GB[1:], key=lambda gv: sharpe(window(paths_gb[("December", gv)], IS)))

# --- all 8 months, one (frac, floor) ---
for gv in GB:
    for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
        grid.append({"test": "all", "level": label(gv), "frac": None if gv is None else gv[0],
                      "floor": None if gv is None else gv[1], "period": per, **score(pooled_gb(gv, yrs))})
all_best = max(GB[1:], key=lambda gv: sharpe(pooled_gb(gv, IS)))

for test, chosen, cur_s, gb_s, tp_s in [
    ("dec", label(dec_best), window(paths_gb[("December", None)], OOS), window(paths_gb[("December", dec_best)], OOS),
     window(paths_tp["December"], OOS)),
    ("all", label(all_best), pooled_gb(None, OOS), pooled_gb(all_best, OOS), pooled_tp(OOS)),
]:
    picks.append({"test": test, "chosen": chosen, "rule": "Current", **score(cur_s)})
    picks.append({"test": test, "chosen": chosen, "rule": "Give-back", **score(gb_s)})
    picks.append({"test": test, "chosen": chosen, "rule": "Profit-take", **score(tp_s)})

# --- walk-forward ---
for y in YEARS[2:]:
    prior = [x for x in YEARS if x < y]
    gv_dec = max(GB[1:], key=lambda gv: sharpe(window(paths_gb[("December", gv)], prior)))
    gv_all = max(GB[1:], key=lambda gv: sharpe(pooled_gb(gv, prior)))
    for test, chosen, cur_s, gb_s in [
        ("dec", label(gv_dec), window(paths_gb[("December", None)], [y]), window(paths_gb[("December", gv_dec)], [y])),
        ("all", label(gv_all), pooled_gb(None, [y]), pooled_gb(gv_all, [y])),
    ]:
        wf.append({"test": test, "year": y, "chosen": chosen, "current": float(cur_s.sum()), "giveback": float(gb_s.sum()),
                   "current_sharpe": sharpe(cur_s), "giveback_sharpe": sharpe(gb_s)})

# --- December OOS equity curves: current vs give-back vs profit-take, all chosen/fixed on 2021-23 ---
cur = window(paths_gb[("December", None)], OOS).cumsum()
gb = window(paths_gb[("December", dec_best)], OOS).cumsum()
tp = window(paths_tp["December"], OOS).cumsum()
for d, a, b, c in zip(cur.index, cur.values, gb.values, tp.values):
    curves.append({"date": d.strftime("%Y-%m-%d"), "current": round(float(a), 2), "giveback": round(float(b), 2),
                    "profit_take": round(float(c), 2)})

meta = [{"label": "giveback_oos.py", "value": Path(__file__).read_text(), "type": "code", "lang": "python"},
        {"label": "exit_backtest.py", "value": Path(lib).read_text(), "type": "code", "lang": "python"},
        {"label": "Input", "value": Path(src).name},
        {"label": "Roll days", "value": "Priced in the old contract through the roll-day close (Roll Prices sheet, roll-fix branch); no dates excluded"}]
for name, rows in [("gb_grid", grid), ("gb_picks", picks), ("gb_walkforward", wf), ("gb_curves", curves)]:
    (outdir / f"{name}.json").write_text(json.dumps({"rows": rows, "meta": meta}, separators=(",", ":")))

pd.set_option("display.width", 220, "display.max_rows", 200, "display.float_format", lambda v: f"{v:,.2f}")
print("December chosen on 2021-23:", label(dec_best), "| all-month chosen on 2021-23:", label(all_best))
G = pd.DataFrame(grid)
print(G[G.test == "dec"].pivot(index="level", columns="period", values="sharpe").reindex([label(g) for g in GB]))
print(G[G.test == "all"].pivot(index="level", columns="period", values="sharpe").reindex([label(g) for g in GB]))
print(pd.DataFrame(picks).to_string())
print(pd.DataFrame(wf).to_string())
