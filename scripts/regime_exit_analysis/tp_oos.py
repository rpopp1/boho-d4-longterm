# Out-of-sample test of a trade-level profit-take on the Knox Position Model tabs.
# Usage: python tp_oos.py <exit_backtest.py> <workbook.xlsx> <outdir>
# The profit-take level is chosen using only earlier data, then scored on later data:
#   split        - choose on 2021-2023, score on 2024-2026
#   walk-forward - for each year Y >= 2023, choose on all years < Y, trade year Y
# Two forms: a flat $ level (December, and the same level in all 8 months), and a multiple k of each
# month's stop (all 8 months, one k). tp_levels/tp_curves also score every flat level for the page's selector.
import json, runpy, sys
from pathlib import Path
import numpy as np
import pandas as pd

lib, src, outdir = sys.argv[1], sys.argv[2], Path(sys.argv[3])
sys.argv = [lib, src]
g = runpy.run_path(lib, run_name="lib")
load, simulate, SB, MONTHS = g["load"], g["simulate"], g["STOP_BASE"], g["MONTHS"]

LEVELS = [None] + list(range(50000, 300001, 25000))           # flat $ (None = current rule, no profit-take)
KS = [None, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0, 4.0]   # profit-take = k x |month's stop|
YEARS = [2021, 2022, 2023, 2024, 2025, 2026]


def sharpe(s):
    return float(s.mean() / s.std() * np.sqrt(252)) if s.std() > 0 else 0.0


def maxdd(s):
    c = s.cumsum()
    return float((c - c.cummax()).min())


data = {m: load(m) for m in MONTHS}
# daily effective PnL for every month x rule, simulated over the full history (exits are path-dependent)
paths = {}
for m in MONTHS:
    for lv in LEVELS:
        eff, _ = simulate(data[m], stop=SB[m], profit=lv)
        paths[(m, "usd", lv)] = pd.Series(eff, index=data[m].Date)
    for k in KS:
        eff, _ = simulate(data[m], stop=SB[m], profit=None if k is None else k * -SB[m])
        paths[(m, "k", k)] = pd.Series(eff, index=data[m].Date)


def window(s, years):
    return s[s.index.year.isin(years)]


def score(s):
    return {"total": float(s.sum()), "sharpe": sharpe(s), "max_dd": maxdd(s)}


def label_usd(lv):
    return "None" if lv is None else f"${lv // 1000}k"


def label_k(k):
    return "None" if k is None else f"{k:g}x stop"


def pooled(form, p, years):
    return pd.concat([window(paths[(m, form, p)], years) for m in MONTHS], axis=1).fillna(0).sum(axis=1)


grid, picks, wf, curves = [], [], [], []
IS, OOS = [2021, 2022, 2023], [2024, 2025, 2026]

# --- December, flat $ level ---
for lv in LEVELS:
    s = paths[("December", "usd", lv)]
    for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
        grid.append({"test": "dec_usd", "level": label_usd(lv), "level_num": lv or 0, "period": per, **score(window(s, yrs))})
is_best = max(LEVELS[1:], key=lambda lv: sharpe(window(paths[("December", "usd", lv)], IS)))

# --- all 8 months, one k ---
for k in KS:
    for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
        grid.append({"test": "all_k", "level": label_k(k), "level_num": k or 0, "period": per, **score(pooled("k", k, yrs))})
    for m in MONTHS:
        for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
            grid.append({"test": f"month_k:{m}", "level": label_k(k), "level_num": k or 0, "period": per,
                         **score(window(paths[(m, "k", k)], yrs))})
k_best = max(KS[1:], key=lambda k: sharpe(pooled("k", k, IS)))

# --- all 8 months, one flat $ level (same level in every month) ---
for lv in LEVELS:
    for per, yrs in [("2021-23", IS), ("2024-26", OOS)]:
        grid.append({"test": "all_usd", "level": label_usd(lv), "level_num": lv or 0, "period": per, **score(pooled("usd", lv, yrs))})

# --- every flat $ level scored on every window: December, the 8-month book and each month ---
# (feeds the page's level selector; period is 2021-23, 2024-26, 2021-26 or a single year)
level_rows = []
PERIODS = [("2021-23", IS), ("2024-26", OOS), ("2021-26", YEARS)] + [(str(y), [y]) for y in YEARS]
for lv in LEVELS:
    for per, yrs in PERIODS:
        level_rows.append({"test": "dec_usd", "level": label_usd(lv), "period": per, **score(window(paths[("December", "usd", lv)], yrs))})
        level_rows.append({"test": "all_usd", "level": label_usd(lv), "period": per, **score(pooled("usd", lv, yrs))})
        for m in MONTHS:
            level_rows.append({"test": f"month_usd:{m}", "level": label_usd(lv), "period": per, **score(window(paths[(m, "usd", lv)], yrs))})

for test, chosen, cur_s, tp_s in [
    ("dec_usd", label_usd(is_best), window(paths[("December", "usd", None)], OOS), window(paths[("December", "usd", is_best)], OOS)),
    ("all_k", label_k(k_best), pooled("k", None, OOS), pooled("k", k_best, OOS)),
]:
    picks.append({"test": test, "chosen": chosen, "rule": "Current", **score(cur_s)})
    picks.append({"test": test, "chosen": chosen, "rule": "Profit-take", **score(tp_s)})
for m in MONTHS:
    for rule, k in [("Current", None), ("Profit-take", k_best)]:
        picks.append({"test": f"month:{m}", "chosen": label_k(k_best), "rule": rule, **score(window(paths[(m, "k", k)], OOS))})

# --- walk-forward: each year traded with the level chosen on all prior years ---
for y in YEARS[2:]:
    prior = [x for x in YEARS if x < y]
    lv = max(LEVELS[1:], key=lambda lv: sharpe(window(paths[("December", "usd", lv)], prior)))
    k = max(KS[1:], key=lambda k: sharpe(pooled("k", k, prior)))
    for test, chosen, cur_s, tp_s in [
        ("dec_usd", label_usd(lv), window(paths[("December", "usd", None)], [y]), window(paths[("December", "usd", lv)], [y])),
        ("all_k", label_k(k), pooled("k", None, [y]), pooled("k", k, [y])),
    ]:
        wf.append({"test": test, "year": y, "chosen": chosen, "current": float(cur_s.sum()), "profit_take": float(tp_s.sum()),
                   "current_sharpe": sharpe(cur_s), "profit_take_sharpe": sharpe(tp_s)})

# --- December equity curves over the full history: current vs level chosen on 2021-23 ---
# period tags each day as the choosing window (2021-23) or the unseen window (2024-26).
# One column per flat $ level: "current" is no profit-take, "lv_<level>" each profit-take level.
cum = {("current" if lv is None else f"lv_{lv}"): window(paths[("December", "usd", lv)], YEARS).cumsum() for lv in LEVELS}
idx = cum["current"].index
for i, d in enumerate(idx):
    row = {"date": d.strftime("%Y-%m-%d"), "period": "2021-23" if d.year in IS else "2024-26"}
    row.update({k: round(float(v.iloc[i]), 2) for k, v in cum.items()})
    curves.append(row)

meta = [{"label": "tp_oos.py", "value": Path(__file__).read_text(), "type": "code", "lang": "python"},
        {"label": "exit_backtest.py", "value": Path(lib).read_text(), "type": "code", "lang": "python"},
        {"label": "Input", "value": Path(src).name},
        {"label": "Roll days", "value": "Priced in the old contract through the roll-day close (Roll Prices sheet, roll-fix branch); no dates excluded"}]
for name, rows in [("tp_grid", grid), ("tp_picks", picks), ("tp_walkforward", wf), ("tp_curves", curves), ("tp_levels", level_rows)]:
    (outdir / f"{name}.json").write_text(json.dumps({"rows": rows, "meta": meta}, separators=(",", ":")))

pd.set_option("display.width", 220, "display.float_format", lambda v: f"{v:,.2f}")
print("December level chosen on 2021-23:", label_usd(is_best), "| all-month k chosen on 2021-23:", label_k(k_best))
G = pd.DataFrame(grid)
print(G[G.test == "dec_usd"].pivot(index="level", columns="period", values="sharpe").reindex([label_usd(l) for l in LEVELS]))
print(G[G.test == "all_k"].pivot(index="level", columns="period", values="sharpe").reindex([label_k(k) for k in KS]))
print(pd.DataFrame(picks).to_string())
print(pd.DataFrame(wf).to_string())
