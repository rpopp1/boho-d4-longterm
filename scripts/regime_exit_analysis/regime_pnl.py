import sys
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as ci

path = sys.argv[1]
months = sys.argv[2].split(",") if len(sys.argv) > 2 else ["December"]
COLS = {"Date": "A", "BOHO": "D", "Z": "L", "Pos": "V", "PnL": "BL", "Active": "BM", "NewTrade": "BN",
        "StopToday": "BP", "StopLatched": "BQ", "EffActive": "BS", "RA": "BT", "dBOHO": "CV", "Vol20": "CW",
        "EWMA": "CY", "ShockRatio": "CZ", "Shock": "DA", "Pctl": "DB", "Regime": "DC"}

wb = openpyxl.load_workbook(path, read_only=True, data_only=True)


def load(month):
    ws = wb[f"Position Model - {month}"]
    idx = {k: ci(v) - 1 for k, v in COLS.items()}
    recs = []
    for r in ws.iter_rows(min_row=2, max_col=max(idx.values()) + 1, values_only=True):
        if r[0] is None:
            continue
        recs.append({k: r[i] for k, i in idx.items()})
    df = pd.DataFrame(recs)
    for c in df.columns:
        if c not in ("Date", "Regime"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df["Regime"].isin(["Normal", "High", "Extreme"])].copy()
    df["Date"] = pd.to_datetime(df["Date"])
    df = df[df["Date"] < "2026-11-16"]  # suspect HO settle on 11/16 (4.26 -> 3.27)
    df["RegimePrev"] = df["Regime"].shift(1)
    df["PctlPrev"] = df["Pctl"].shift(1)
    df["EWMAPrev"] = df["EWMA"].shift(1)
    df["Year"] = df["Date"].dt.year
    return df.reset_index(drop=True)


ORDER = ["Normal", "High", "Extreme"]


def stats(g, col="RA"):
    x = g[col]
    nz = x[x != 0]
    return pd.Series({
        "days": len(x), "pnl_days": len(nz), "sum": x.sum(), "mean_on_pnl_days": nz.mean(),
        "std_on_pnl_days": nz.std(), "sharpe_ann": nz.mean() / nz.std() * np.sqrt(252) if len(nz) > 2 else np.nan,
        "hit": (nz > 0).mean(), "p05": nz.quantile(.05) if len(nz) else np.nan,
        "p95": nz.quantile(.95) if len(nz) else np.nan, "worst": nz.min(), "best": nz.max(),
        "mean_abs": nz.abs().mean(),
    })


pd.set_option("display.width", 250, "display.max_columns", 30, "display.float_format", lambda v: f"{v:,.3f}")

for m in months:
    df = load(m)
    print(f"\n================ {m}  ({df.Date.min():%Y-%m-%d} -> {df.Date.max():%Y-%m-%d}, {len(df)} days) ================")
    print("\nRegime day counts by year:")
    print(pd.crosstab(df.Year, df.Regime)[ORDER])
    print("\nRegime counts on days in an effective trade (BS=1):")
    print(df[df.EffActive == 1].Regime.value_counts().reindex(ORDER))

    print("\n[A] Same-day regime vs RA daily PnL (BT):")
    print(df.groupby("Regime").apply(stats).reindex(ORDER))
    print("\n[B] PRIOR-day regime (tradable signal) vs RA daily PnL:")
    print(df.groupby("RegimePrev").apply(stats).reindex(ORDER))

    act = df[df.RA != 0]
    print("\n[C] Correlations on days with nonzero RA PnL (n=%d):" % len(act))
    for a in ["Pctl", "EWMA", "PctlPrev", "EWMAPrev", "ShockRatio"]:
        print(f"  {a:10s}  corr(PnL)={act[a].corr(act.RA): .3f}  corr(|PnL|)={act[a].corr(act.RA.abs()): .3f}  "
              f"spearman(PnL)={act[a].rank().corr(act.RA.rank()): .3f}")

    print("\n[D] EWMA percentile quintiles (same day), nonzero-PnL days:")
    act = act.copy()
    act["PctlBin"] = pd.cut(act.Pctl, [-.01, .2, .4, .6, .7, .8, .9, 1.0])
    print(act.groupby("PctlBin", observed=True).apply(stats))

    print("\n[E] Shock flag (EWMA >= 1.25x 20d vol) vs RA PnL:")
    print(df.groupby("Shock").apply(stats))
    print(pd.crosstab(df.Regime, df.Shock).reindex(ORDER))

    # Dollar-vol normalised PnL: does PnL scale with vol, or does edge per unit vol change?
    act["PnL_per_vol"] = act.RA / act.EWMA
    print("\n[F] PnL per unit EWMA vol by regime (mean, nonzero days):")
    print(act.groupby("Regime").PnL_per_vol.agg(["mean", "std", "count"]).reindex(ORDER))

    # Trade-level: regime at entry
    df["TradeId"] = (df.NewTrade == 1).cumsum()
    tr = df[(df.Active == 1) | (df.RA != 0)].groupby("TradeId").agg(
        start=("Date", "first"), end=("Date", "last"), entry_regime=("Regime", "first"),
        entry_pctl=("Pctl", "first"), max_regime=("Pctl", "max"), ra_pnl=("RA", "sum"),
        stopped=("StopToday", "max"), days=("Date", "count"),
        extreme_days=("Regime", lambda s: (s == "Extreme").sum()))
    tr = tr[tr.index > 0]
    print(f"\n[G] Trade-level by ENTRY regime ({len(tr)} trades):")
    print(tr.groupby("entry_regime").agg(n=("ra_pnl", "size"), total=("ra_pnl", "sum"), avg=("ra_pnl", "mean"),
                                          med=("ra_pnl", "median"), win=("ra_pnl", lambda s: (s > 0).mean()),
                                          stop_rate=("stopped", "mean"), avg_days=("days", "mean")).reindex(ORDER))
    tr["touched_extreme"] = tr.extreme_days > 0
    print("\n    Trades that experienced any Extreme day during life:")
    print(tr.groupby("touched_extreme").agg(n=("ra_pnl", "size"), total=("ra_pnl", "sum"), avg=("ra_pnl", "mean"),
                                             win=("ra_pnl", lambda s: (s > 0).mean()), stop_rate=("stopped", "mean")))
    print("\n    Stop-outs: regime on stop day")
    print(df[df.StopToday == 1].Regime.value_counts().reindex(ORDER))
    print("\n    Trade list:")
    print(tr.to_string())

    print("\n[H] Regime x Year: RA PnL sum")
    print(pd.pivot_table(df, index="Year", columns="Regime", values="RA", aggfunc="sum").reindex(columns=ORDER))

    # Regime transitions: PnL in the 5 days after entering Extreme vs rest of Extreme
    df["EnterExt"] = (df.Regime == "Extreme") & (df.RegimePrev != "Extreme")
    run = (df.Regime != df.Regime.shift()).cumsum()
    df["DayInRun"] = df.groupby(run).cumcount() + 1
    ex = df[df.Regime == "Extreme"]
    print("\n[I] Extreme regime: PnL by day-in-run bucket")
    print(ex.groupby(pd.cut(ex.DayInRun, [0, 3, 10, 30, 1000])).apply(stats)[["days", "pnl_days", "sum", "mean_on_pnl_days", "hit"]])

    out = df[["Date", "Z", "Pos", "RA", "EWMA", "Vol20", "Pctl", "Regime", "Shock", "EffActive", "StopToday"]]
    out.to_csv(sys.argv[0].replace("regime_pnl.py", f"regime_{m}.csv"), index=False)
