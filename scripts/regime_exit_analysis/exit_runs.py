import sys, runpy, io, contextlib
import pandas as pd, numpy as np
src = sys.argv[2]
sys.argv = [sys.argv[1], src]
g = runpy.run_path(sys.argv[0], run_name="lib")
load, simulate, metrics, STOP_BASE, MONTHS = g["load"], g["simulate"], g["metrics"], g["STOP_BASE"], g["MONTHS"]
pd.set_option("display.width", 260, "display.max_columns", 30, "display.max_rows", 200, "display.float_format", lambda v: f"{v:,.2f}")
cols = ["total", "sharpe", "maxDD", "calmar", "pnl_21_23", "pnl_24_26", "sh_21_23", "sh_24_26", "trades", "win", "avg_win", "avg_loss", "worst", "exits"]
dec = load("December")
print("EWMA/expanding-median ratio at trade entries:", dec.loc[dec.BN == 1, "EWMA"].div(dec.loc[dec.BN == 1, "EWMARef"]).describe().round(2).to_dict())

def run(name, **kw):
    eff, tr = simulate(dec, **kw)
    return pd.Series(metrics(dec, eff, tr), name=name), tr

R = {}
for nm, kw in [("No stop", {}), ("Current: fixed -65k", {"stop": -65000})]:
    R[nm] = run(nm, **kw)
print("\n=== TEST 1: vol-scaled stop (replaces fixed stop) ===")
t1 = [R["No stop"][0], R["Current: fixed -65k"][0]]
for base in [-50000, -65000, -80000, -100000]:
    for mode in ["entry", "daily"]:
        for lo, hi in [(0.5, 3.0), (1.0, 3.0)]:
            t1.append(run(f"vol {base/1000:.0f}k {mode} [{lo},{hi}]", vol_stop=(base, lo, hi, mode))[0])
print(pd.DataFrame(t1)[cols].to_string())

print("\n=== TEST 2: profit take (with current -65k stop) ===")
t2 = [R["Current: fixed -65k"][0]]
for p in [50000, 75000, 100000, 125000, 150000, 200000, 250000]:
    t2.append(run(f"TP {p/1000:.0f}k", stop=-65000, profit=p)[0])
print(pd.DataFrame(t2)[cols].to_string())

print("\n=== TEST 3: time stop in Extreme (with current -65k stop) ===")
t3 = [R["Current: fixed -65k"][0]]
for n in [1, 2, 3, 4, 5, 7, 10]:
    t3.append(run(f"Exit after {n} Extreme days held", stop=-65000, ext_days=n)[0])
for n in [2, 3, 5, 10]:
    t3.append(run(f"Exit on Extreme day once age>={n}", stop=-65000, ext_age=n)[0])
print(pd.DataFrame(t3)[cols].to_string())

# trade-level detail for the three headline variants
base_tr = R["Current: fixed -65k"][1].reset_index(drop=True)
for nm, kw in [("TP 100k", dict(stop=-65000, profit=100000)), ("Time 3 Extreme days", dict(stop=-65000, ext_days=3)),
               ("Vol stop -65k entry [0.5,3]", dict(vol_stop=(-65000, 0.5, 3.0, "entry")))]:
    _, tr = simulate(dec, **kw); tr = tr.reset_index(drop=True)
    d = pd.DataFrame({"start": base_tr.start.dt.strftime("%Y-%m-%d"), "entry_regime": base_tr.entry_regime,
                      "ratio": base_tr.ratio.round(2), "base_pnl": base_tr.pnl, "base_exit": base_tr.reason,
                      "new_pnl": tr.pnl, "new_exit": tr.reason})
    d["delta"] = d.new_pnl - d.base_pnl
    ch = d[d.delta.abs() > 1]
    print(f"\n--- {nm}: {len(ch)} trades changed, net {ch.delta.sum():,.0f} ---")
    print(ch.to_string())

print("\n=== Robustness: same rules on all 8 months (each month's own Config stop) ===")
rows = []
for m in MONTHS:
    df = load(m); sb = STOP_BASE[m]
    for nm, kw in [("Current fixed", dict(stop=sb)), ("Vol stop entry [0.5,3]", dict(vol_stop=(sb, 0.5, 3.0, "entry"))),
                   ("Vol stop entry [1,3]", dict(vol_stop=(sb, 1.0, 3.0, "entry"))),
                   ("TP 100k", dict(stop=sb, profit=100000)), ("TP 150k", dict(stop=sb, profit=150000)),
                   ("Time 3 Extreme days", dict(stop=sb, ext_days=3)), ("Time 5 Extreme days", dict(stop=sb, ext_days=5))]:
        eff, tr = simulate(df, **kw); mt = metrics(df, eff, tr)
        rows.append({"month": m, "rule": nm, "total": mt["total"], "sharpe": mt["sharpe"], "maxDD": mt["maxDD"]})
r = pd.DataFrame(rows)
for f in ["total", "sharpe", "maxDD"]:
    p = r.pivot(index="month", columns="rule", values=f).reindex(MONTHS)
    base = p["Current fixed"]
    print(f"\n{f}:"); print(p.to_string())
    print("months better than current:", {c: int(((p[c] > base) if f != "maxDD" else (p[c] > base)).sum()) for c in p.columns if c != "Current fixed"})
