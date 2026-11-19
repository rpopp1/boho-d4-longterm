# Test 4 (month-start re-entry) and Test 5 (N-day time exit) on the Knox Position Model tabs.
# Usage: python reentry_maxdays.py <exit_backtest.py> <workbook.xlsx> <outdir>
# Writes all_months.json (the Re-Entry Rule Results dashboard's data: every month x 5 rules)
# and prints the Test 5 grid (8-month book and December, N-day cap with and without the
# $100k profit-take). Sharpe counts every day, as the Config sheet does.
import json, runpy, sys
from pathlib import Path
import numpy as np
import pandas as pd

lib, src, outdir = sys.argv[1], sys.argv[2], Path(sys.argv[3])
sys.argv = [lib, src]
g = runpy.run_path(lib, run_name="lib")
load, simulate, SB, MONTHS = g["load"], g["simulate"], g["STOP_BASE"], g["MONTHS"]
TP = 100000
N_DAYS = [None, 3, 5, 10, 15, 20, 30, 40]


def stats(eff, index):
    s = pd.Series(eff, index=index)
    c = s.cumsum()
    return {"total": float(s.sum()), "sharpe": float(s.mean() / s.std() * np.sqrt(252)),
            "maxDD": float((c - c.cummax()).min())}


data = {m: load(m) for m in MONTHS}

# --- Test 4: re-entry rules ---
RULES = [
    ("stop only", {}),
    ("stop + reentry_month", {"reentry_month": True}),
    ("stop + TP100k", {"profit": TP}),
    ("stop + TP100k + reentry_month (any reason)", {"profit": TP, "reentry_month": True}),
    ("stop + TP100k + reentry_month (profit only)", {"profit": TP, "reentry_month": "profit_only"}),
]
rows = []
for m in MONTHS:
    for name, kw in RULES:
        eff, tr = simulate(data[m], stop=SB[m], **kw)
        reentries = int(tr["reentry"].fillna(False).astype(bool).sum()) if "reentry" in tr else 0
        rows.append({"month": m, "rule": name, **{k: round(v, 2) for k, v in stats(eff, data[m].Date).items()},
                     "reentries": reentries})
r = pd.DataFrame(rows)
pd.set_option("display.width", 220, "display.max_rows", 200, "display.float_format", lambda v: f"{v:,.2f}")
print("=== Test 4: re-entry rules ===")
print(r.pivot(index="month", columns="rule", values="total").reindex(MONTHS).to_string())
print(r.pivot(index="month", columns="rule", values="sharpe").reindex(MONTHS).to_string())
print(r.pivot(index="month", columns="rule", values="maxDD").reindex(MONTHS).to_string())
print(r.pivot(index="month", columns="rule", values="reentries").reindex(MONTHS).to_string())
base, add = "stop + TP100k", "stop + TP100k + reentry_month (profit only)"
p = r.pivot(index="month", columns="rule", values=["total", "sharpe", "maxDD"])
print("profit-only gate vs plain TP: total better/equal in", int((p["total"][add] >= p["total"][base] - 0.5).sum()),
      "| Sharpe better/equal in", int((p["sharpe"][add] >= p["sharpe"][base] - 1e-9).sum()),
      "| max DD unchanged in", int((abs(p["maxDD"][add] - p["maxDD"][base]) < 0.5).sum()), "of 8")

meta = [{"label": "reentry_maxdays.py", "value": Path(__file__).read_text(), "type": "code", "lang": "python"},
        {"label": "exit_backtest.py", "value": Path(lib).read_text(), "type": "code", "lang": "python"},
        {"label": "Workbook", "value": Path(src).name},
        {"label": "Roll days", "value": "Priced in the old contract through the roll-day close (Roll Prices sheet, roll-fix branch); no dates excluded"},
        {"label": "Profit-take level", "value": "$100,000 flat trade-level exit"}]
(outdir / "all_months.json").write_text(json.dumps({"rows": rows, "meta": meta}, separators=(",", ":")))

# --- Test 5: N-day time exit ---
print("\n=== Test 5: N-day time exit (book = 8 months summed; Dec alone) ===")
out = []
for n in N_DAYS:
    for tp_on in (False, True):
        kw = {"max_days": n} if n else {}
        if tp_on:
            kw["profit"] = TP
        effs = {m: pd.Series(simulate(data[m], stop=SB[m], **kw)[0], index=data[m].Date) for m in MONTHS}
        book = pd.concat(effs.values(), axis=1).fillna(0).sum(axis=1)
        bs, ds = stats(book.values, book.index), stats(effs["December"].values, effs["December"].index)
        out.append({"N": n or "no cap", "profit_take": tp_on, "book_total": bs["total"], "book_sharpe": bs["sharpe"],
                    "book_maxDD": bs["maxDD"], "dec_total": ds["total"], "dec_sharpe": ds["sharpe"], "dec_maxDD": ds["maxDD"]})
print(pd.DataFrame(out).to_string(index=False))
