import sys, runpy, io, contextlib
import numpy as np, pandas as pd
months = ["January","March","May","July","August","September","October","December"]
sys.argv = [sys.argv[1], sys.argv[2], "December"]
with contextlib.redirect_stdout(io.StringIO()):
    g = runpy.run_path(sys.argv[0])
pd.set_option("display.width", 250, "display.float_format", lambda v: f"{v:,.0f}")
out = []; trs = []
for m in months:
    df = g["load"](m)
    run = (df.Regime != df.Regime.shift()).cumsum()
    df["DayInRun"] = df.groupby(run).cumcount() + 1
    ex = df[(df.Regime == "Extreme") & (df.RA != 0)]
    out.append(ex.assign(month=m, b=pd.cut(ex.DayInRun, [0, 10, 1000], labels=["d1-10", "d11+"])))
    df["TradeId"] = (df.NewTrade == 1).cumsum()
    t = df[(df.Active == 1) | (df.RA != 0)].groupby("TradeId").agg(pnl=("RA", "sum"), stop=("StopToday", "max"),
        ext=("Regime", lambda s: (s == "Extreme").any()), entry=("Regime", "first"))
    trs.append(t[t.index > 0].assign(month=m))
e = pd.concat(out)
print(e.pivot_table(index="month", columns="b", values="RA", aggfunc="sum", observed=True).reindex(months))
print(e.groupby("b", observed=True).RA.agg(["count", "sum", "mean"]))
t = pd.concat(trs)
print("\nAll-month trades, touched Extreme during life:")
print(t.groupby("ext").agg(n=("pnl", "size"), total=("pnl", "sum"), avg=("pnl", "mean"), win=("pnl", lambda s: (s > 0).mean()), stop=("stop", "mean")).to_string(float_format=lambda v: f"{v:,.2f}"))
print("\nAll-month trades by entry regime:")
print(t.groupby("entry").agg(n=("pnl", "size"), total=("pnl", "sum"), avg=("pnl", "mean"), win=("pnl", lambda s: (s > 0).mean()), stop=("stop", "mean")).to_string(float_format=lambda v: f"{v:,.2f}"))
