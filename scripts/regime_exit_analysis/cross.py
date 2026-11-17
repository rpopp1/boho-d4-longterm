import sys, runpy, io, contextlib
import numpy as np, pandas as pd
months = ["January","March","May","July","August","September","October","December"]
sys.argv = [sys.argv[1], sys.argv[2], "December"]
with contextlib.redirect_stdout(io.StringIO()):
    g = runpy.run_path(sys.argv[0])
load = g["load"]
rows = []; allp = []
for m in months:
    df = load(m)
    for lag, col in [("same", "Regime"), ("prev", "RegimePrev")]:
        for reg, x in df[df.RA != 0].groupby(col):
            rows.append(dict(month=m, lag=lag, regime=reg, n=len(x), sum=x.RA.sum(), mean=x.RA.mean(),
                             std=x.RA.std(), sharpe=x.RA.mean()/x.RA.std()*np.sqrt(252), hit=(x.RA>0).mean(),
                             p05=x.RA.quantile(.05), perVol=(x.RA/x.EWMA).mean()))
    allp.append(df.assign(month=m))
r = pd.DataFrame(rows)
pd.set_option("display.width", 250, "display.float_format", lambda v: f"{v:,.2f}")
for lag in ["same", "prev"]:
    print(f"\n=== {lag}-day regime: annualised Sharpe of nonzero RA days ===")
    print(r[r.lag==lag].pivot(index="month", columns="regime", values="sharpe").reindex(months)[["Normal","High","Extreme"]])
    print(f"--- std of daily RA PnL ---")
    print(r[r.lag==lag].pivot(index="month", columns="regime", values="std").reindex(months)[["Normal","High","Extreme"]])
    print(f"--- total RA PnL ---")
    print(r[r.lag==lag].pivot(index="month", columns="regime", values="sum").reindex(months)[["Normal","High","Extreme"]])
    print(f"--- 5th pct daily RA PnL ---")
    print(r[r.lag==lag].pivot(index="month", columns="regime", values="p05").reindex(months)[["Normal","High","Extreme"]])
a = pd.concat(allp); a = a[a.RA != 0]
print("\n=== Pooled all 8 months (nonzero RA days) ===")
for col in ["Regime", "RegimePrev"]:
    t = a.groupby(col).RA.agg(["count","sum","mean","std"]).reindex(["Normal","High","Extreme"])
    t["sharpe"] = t["mean"]/t["std"]*np.sqrt(252); t["hit"] = a.groupby(col).RA.apply(lambda s:(s>0).mean())
    print(col); print(t)
# portfolio-level: sum across months per date, regime = December's label
port = pd.concat(allp).groupby("Date").RA.sum().rename("Port").to_frame().join(load("December").set_index("Date")[["Regime","RegimePrev"]])
port = port[port.Port != 0]
print("\n=== Portfolio daily RA PnL (sum of 8 months) by December regime ===")
for col in ["Regime","RegimePrev"]:
    t = port.groupby(col).Port.agg(["count","sum","mean","std"]).reindex(["Normal","High","Extreme"]); t["sharpe"]=t["mean"]/t["std"]*np.sqrt(252); print(col); print(t)
