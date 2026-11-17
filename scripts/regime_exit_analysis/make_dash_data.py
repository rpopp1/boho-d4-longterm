# Builds regime_daily.json and regime_trades.json from the Knox multi-month
# workbook's Position Model tabs (cached values). Usage: python make_dash_data.py <workbook.xlsx> <outdir>
import json, sys
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as ci

src, outdir = Path(sys.argv[1]), Path(sys.argv[2])
MONTHS = ["January", "March", "May", "July", "August", "September", "October", "December"]
COLS = {"Date": "A", "Active": "BM", "NewTrade": "BN", "StopToday": "BP", "RA": "BT",
        "EWMA": "CY", "Shock": "DA", "Pctl": "DB", "Regime": "DC"}
EXCLUDE = {"December": ["2026-11-16"]}  # Dec HO settle 4.2585 -> 3.2674 on 11/16: suspect print

wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
code = Path(__file__).read_text()
daily, trades = [], []
for m in MONTHS:
    ws = wb[f"Position Model - {m}"]
    idx = {k: ci(v) - 1 for k, v in COLS.items()}
    recs = [{k: r[i] for k, i in idx.items()}
            for r in ws.iter_rows(min_row=2, max_col=max(idx.values()) + 1, values_only=True) if r[0] is not None]
    df = pd.DataFrame(recs)
    for c in df.columns:
        if c not in ("Date", "Regime"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df.Regime.isin(["Normal", "High", "Extreme"])].copy()
    df["Date"] = pd.to_datetime(df.Date)
    df = df[~df.Date.dt.strftime("%Y-%m-%d").isin(EXCLUDE.get(m, []))].reset_index(drop=True)
    df["RegimePrev"] = df.Regime.shift(1)
    run = (df.Regime != df.Regime.shift()).cumsum()
    df["DayInRun"] = df.groupby(run).cumcount() + 1
    df["TradeId"] = (df.NewTrade == 1).cumsum()
    for r in df.itertuples():
        daily.append({"month": m, "date": r.Date.strftime("%Y-%m-%d"), "ra_pnl": round(float(r.RA), 2) if r.RA == r.RA else 0.0,  # BT warm-up rows are #N/A
                      "ewma_vol": round(float(r.EWMA), 5), "ewma_pctl": round(float(r.Pctl), 4),
                      "regime": r.Regime, "regime_prev": r.RegimePrev if isinstance(r.RegimePrev, str) else None,
                      "day_in_run": int(r.DayInRun), "shock": int(r.Shock) if r.Shock == r.Shock else 0})
    t = df[(df.Active == 1) | (df.RA != 0)].groupby("TradeId").agg(
        start=("Date", "first"), end=("Date", "last"), entry_regime=("Regime", "first"),
        ra_pnl=("RA", "sum"), stopped=("StopToday", "max"), days=("Date", "count"),
        extreme_days=("Regime", lambda s: int((s == "Extreme").sum())))
    for tid, r in t[t.index > 0].iterrows():
        trades.append({"month": m, "trade": f"{m[:3]}-{int(tid)}", "start": r.start.strftime("%Y-%m-%d"),
                       "end": r.end.strftime("%Y-%m-%d"), "entry_regime": r.entry_regime,
                       "ra_pnl": round(float(r.ra_pnl), 2) if r.ra_pnl == r.ra_pnl else 0.0, "stopped": int(r.stopped or 0), "days": int(r.days),
                       "extreme_days": int(r.extreme_days),
                       "touched_extreme": "Yes" if r.extreme_days > 0 else "No"})

meta = [{"label": "make_dash_data.py", "value": code, "type": "code", "lang": "python"},
        {"label": "Input", "value": src.name},
        {"label": "Excluded", "value": "December 2026-11-16 (suspect HO settle 4.2585 -> 3.2674)"}]
(outdir / "regime_daily.json").write_text(json.dumps({"rows": daily, "meta": meta}, separators=(",", ":")))
(outdir / "regime_trades.json").write_text(json.dumps({"rows": trades, "meta": meta}, separators=(",", ":")))
print(len(daily), "daily rows;", len(trades), "trades")
