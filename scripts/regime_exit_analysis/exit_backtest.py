# Exit-rule overlay backtest on the Knox Position Model tabs.
# Mirrors the workbook's stop overlay (BM/BN/BO/BP/BQ/BS/BT) on the gross daily PnL stream (BL)
# and adds alternative exit rules. Positions/signals are untouched: an exit just goes flat
# until the next New Trade Start flag (BN=1), exactly like the workbook's latched stop.
import sys
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as ci

SRC = sys.argv[1]
MONTHS = ["January", "March", "May", "July", "August", "September", "October", "December"]
STOP_BASE = {"January": -70000, "March": -100000, "May": -100000, "July": -70000, "August": -80000,
             "September": -50000, "October": -80000, "December": -65000}  # Config row 27
EXCLUDE = {"December": ["2026-11-16"]}
COLS = {"Date": "A", "BL": "BL", "BN": "BN", "BT": "BT", "EWMA": "CY", "Regime": "DC"}

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)


def load(month):
    ws = wb[f"Position Model - {month}"]
    idx = {k: ci(v) - 1 for k, v in COLS.items()}
    recs = [{k: r[i] for k, i in idx.items()}
            for r in ws.iter_rows(min_row=2, max_col=max(idx.values()) + 1, values_only=True) if r[0] is not None]
    df = pd.DataFrame(recs)
    for c in ("BL", "BN", "BT", "EWMA"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Date"] = pd.to_datetime(df.Date)
    df = df[~df.Date.dt.strftime("%Y-%m-%d").isin(EXCLUDE.get(month, []))]
    df = df[df.Date.notna()].reset_index(drop=True)
    df[["BL", "BN", "BT"]] = df[["BL", "BN", "BT"]].fillna(0)
    # vol reference: expanding median of EWMA vol known up to that day (no look-ahead)
    df["EWMARef"] = df.EWMA.expanding(min_periods=20).median()
    return df


def simulate(df, stop=None, vol_stop=None, profit=None, ext_days=None, ext_age=None):
    """Return daily effective PnL and a per-trade log.
    stop: fixed $ stop (negative). vol_stop: (base, lo, hi, mode) with mode 'entry' or 'daily'.
    profit: take-profit $ level. ext_days: exit once N Extreme-regime days have been held in the trade.
    ext_age: exit on any Extreme day once the trade is >= N days old."""
    BL, BN, ew, ref, reg = df.BL.values, df.BN.values, df.EWMA.values, df.EWMARef.values, df.Regime.values
    n = len(df)
    eff = np.zeros(n)
    trades, active, latched, prev_bo, bo = [], False, False, 0.0, 0.0
    cur = None
    for t in range(n):
        if BN[t] == 1:
            active, latched, bo, prev_bo = True, False, 0.0, 0.0
            ratio = ew[t] / ref[t] if ref[t] == ref[t] and ref[t] > 0 else 1.0
            cur = {"start": df.Date[t], "pnl": 0.0, "reason": "signal", "age": 0, "ext": 0, "ratio": ratio,
                   "entry_regime": reg[t]}
            trades.append(cur)
        elif active and BL[t] == 0:
            active = False  # trade closed by the signal (BM -> 0)
        if not active:
            continue
        bo = bo + BL[t]
        cur["age"] += 1
        cur["ext"] += reg[t] == "Extreme"
        if latched:
            continue
        eff[t] = BL[t]
        cur["pnl"] += BL[t]
        cur["end"] = df.Date[t]
        hit = None
        if stop is not None and bo <= stop and prev_bo > stop:  # workbook BP: first close at/below the stop
            hit = "stop"
        if vol_stop is not None:
            base, lo, hi, mode = vol_stop
            r = cur["ratio"] if mode == "entry" else (ew[t] / ref[t] if ref[t] == ref[t] and ref[t] > 0 else 1.0)
            lvl = base * min(max(r, lo), hi)
            if bo <= lvl:
                hit = hit or "vol stop"
        if profit is not None and bo >= profit:
            hit = hit or "profit"
        if ext_days is not None and reg[t] == "Extreme" and cur["ext"] >= ext_days:
            hit = hit or "time"
        if ext_age is not None and reg[t] == "Extreme" and cur["age"] >= ext_age:
            hit = hit or "time"
        if hit and BN[t] != 1:  # workbook BQ is forced to 0 on a New Trade Start day: no exit on day 1
            latched = True
            cur["reason"] = hit
        prev_bo = bo
    return eff, pd.DataFrame(trades)


def metrics(df, eff, tr):
    s = pd.Series(eff, index=df.Date)
    cum = s.cumsum()
    dd = (cum - cum.cummax()).min()
    sh = s.mean() / s.std() * np.sqrt(252)
    a, b = s[s.index.year <= 2023], s[s.index.year >= 2024]
    w, l = tr.pnl[tr.pnl > 0], tr.pnl[tr.pnl <= 0]
    return {"total": s.sum(), "sharpe": sh, "maxDD": dd, "calmar": s.sum() / -dd if dd < 0 else np.nan,
            "pnl_21_23": a.sum(), "pnl_24_26": b.sum(),
            "sh_21_23": a.mean() / a.std() * np.sqrt(252), "sh_24_26": b.mean() / b.std() * np.sqrt(252),
            "trades": len(tr), "win": (tr.pnl > 0).mean(), "avg_win": w.mean(), "avg_loss": l.mean(),
            "worst": tr.pnl.min(), "best": tr.pnl.max(),
            "exits": int((tr.reason != "signal").sum())}


if __name__ == "__main__":
    pd.set_option("display.width", 250, "display.max_columns", 30, "display.max_rows", 200,
                  "display.float_format", lambda v: f"{v:,.2f}")
    dec = load("December")
    eff, tr = simulate(dec, stop=STOP_BASE["December"])
    print("Validation vs workbook BT: max abs diff =", np.abs(eff - dec.BT.values).max(),
          " sums:", eff.sum(), dec.BT.sum())
