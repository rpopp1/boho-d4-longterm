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
EXCLUDE = {}  # roll days are priced correctly by the roll fix; nothing to drop
COLS = {"Date": "A", "BL": "BL", "BN": "BN", "BT": "BT", "EWMA": "CY", "Regime": "DC", "ZScore": "L"}

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)


def load(month):
    ws = wb[f"Position Model - {month}"]
    idx = {k: ci(v) - 1 for k, v in COLS.items()}
    recs = [{k: r[i] for k, i in idx.items()}
            for r in ws.iter_rows(min_row=2, max_col=max(idx.values()) + 1, values_only=True) if r[0] is not None]
    df = pd.DataFrame(recs)
    for c in ("BL", "BN", "BT", "EWMA", "ZScore"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Date"] = pd.to_datetime(df.Date)
    df = df[~df.Date.dt.strftime("%Y-%m-%d").isin(EXCLUDE.get(month, []))]
    df = df[df.Date.notna()].reset_index(drop=True)
    df[["BL", "BN", "BT"]] = df[["BL", "BN", "BT"]].fillna(0)
    # vol reference: expanding median of EWMA vol known up to that day (no look-ahead)
    df["EWMARef"] = df.EWMA.expanding(min_periods=20).median()
    # trailing realized $ vol of the raw daily PnL stream, known as of yesterday's close (no look-ahead).
    # EWMA (column CY) is a normalized vol ratio (~O(1)), not a $ figure, so it can't size a $ swing
    # threshold directly; this gives a $-denominated "how big is a typical day lately" reference instead.
    df["BLVol"] = df.BL.rolling(20, min_periods=10).std().shift(1)
    # first trading day of a new calendar month (for the weight-refresh re-entry test)
    df["MonthStart"] = df.Date.dt.month != df.Date.dt.month.shift(1)
    return df


def simulate(df, stop=None, vol_stop=None, profit=None, ext_days=None, ext_age=None, giveback=None,
             profit_swing=None, reentry=None, reentry_month=False, max_days=None, first_day_stop=None):
    """Return daily effective PnL and a per-trade log.
    stop: fixed $ stop (negative). vol_stop: (base, lo, hi, mode) with mode 'entry' or 'daily'.
    profit: take-profit $ level. ext_days: exit once N Extreme-regime days have been held in the trade.
    ext_age: exit on any Extreme day once the trade is >= N days old.
    giveback: exit once trade PnL retraces this fraction of its peak gain (peak must be > 0).
    Pass a (fraction, floor) tuple to require the peak to clear $floor before the rule arms.
    profit_swing: (level, threshold, mode, direction) - once trade PnL reaches level, exit on the
    first day its single-day gross PnL move (BL) clears `threshold` (mode 'usd') or `threshold` x
    the trailing 20-day realized $ std of BL as of yesterday (mode 'vol', adapts to how choppy the
    month has been lately). direction 'up' only arms on a big up day (BL >= threshold); 'abs' arms
    on a big move either way. A grinding trade above level with no qualifying day keeps running.
    reentry: if not None, an integer day-count. Once a trade is latched flat by any exit above,
    if the underlying signal is still active (BL != 0, i.e. BM hasn't gone flat) for `reentry`
    full days after the exit, a fresh sub-trade re-opens (its own stop/profit/peak tracking from
    zero) instead of staying flat until the next New Trade Start flag. 0 = re-open the next day.
    reentry_month: if True, re-opens only on the first trading day of a new calendar month (a
    weight-refresh event), and only if the signal is still active AND that day's Z-Score is at
    least as large in magnitude, same sign, as the Z-Score on the day of the original exit -
    i.e. only re-enter on genuinely fresh/stronger evidence, not just time passing.
    Pass 'profit_only' instead of True to additionally require the original exit that latched
    the trade to have been a profit-take (cur["reason"] == "profit"), not a stop/time/giveback
    exit - i.e. only chase a position that was cut short on a win, never one proven wrong.
    max_days: time exit - go flat once the trade has been open this many trading days (day 1 = the
    New Trade Start day), latched until the next New Trade Start like every other exit.
    first_day_stop: None (workbook rule: nothing can exit on a New Trade Start day, so a trade
    whose first close is already past the stop never gets its crossing day), 'all' (the stop may
    fire on day 1 too) or 'fresh' (only on day 1 of a trade entered from flat, not on a flip day,
    whose PnL includes the old position's overnight move)."""
    gb_frac, gb_floor = giveback if isinstance(giveback, tuple) else (giveback, 0.0)
    ps_level, ps_thr, ps_mode, ps_dir = profit_swing if profit_swing is not None else (None, None, None, "abs")
    BL, BN, ew, ref, reg = df.BL.values, df.BN.values, df.EWMA.values, df.EWMARef.values, df.Regime.values
    blvol, z, mstart = df.BLVol.values, df.ZScore.values, df.MonthStart.values
    n = len(df)
    eff = np.zeros(n)
    trades, active, latched, prev_bo, bo, latch_day = [], False, False, 0.0, 0.0, 0
    cur = None
    for t in range(n):
        if BN[t] == 1:
            active, latched, bo, prev_bo, latch_day = True, False, 0.0, 0.0, 0
            ratio = ew[t] / ref[t] if ref[t] == ref[t] and ref[t] > 0 else 1.0
            cur = {"start": df.Date[t], "pnl": 0.0, "reason": "signal", "age": 0, "ext": 0, "ratio": ratio,
                   "entry_regime": reg[t], "peak": 0.0}
            trades.append(cur)
        elif reentry is not None and active and latched and BL[t] != 0 and latch_day >= reentry:
            latched, bo, prev_bo = False, 0.0, 0.0
            ratio = ew[t] / ref[t] if ref[t] == ref[t] and ref[t] > 0 else 1.0
            cur = {"start": df.Date[t], "pnl": 0.0, "reason": "signal", "age": 0, "ext": 0, "ratio": ratio,
                   "entry_regime": reg[t], "peak": 0.0, "reentry": True}
            trades.append(cur)
        elif (reentry_month and active and latched and BL[t] != 0 and mstart[t]
              and z[t] == z[t] and cur.get("exit_z", None) is not None
              and np.sign(z[t]) == np.sign(cur["exit_z"]) and abs(z[t]) >= abs(cur["exit_z"])
              and (reentry_month != "profit_only" or cur.get("reason") == "profit")):
            latched, bo, prev_bo = False, 0.0, 0.0
            ratio = ew[t] / ref[t] if ref[t] == ref[t] and ref[t] > 0 else 1.0
            cur = {"start": df.Date[t], "pnl": 0.0, "reason": "signal", "age": 0, "ext": 0, "ratio": ratio,
                   "entry_regime": reg[t], "peak": 0.0, "reentry": True, "reentry_z": z[t]}
            trades.append(cur)
        elif active and BL[t] == 0:
            active = False  # trade closed by the signal (BM -> 0)
        if not active:
            continue
        bo = bo + BL[t]
        cur["age"] += 1
        cur["ext"] += reg[t] == "Extreme"
        cur["peak"] = max(cur["peak"], bo)
        if latched:
            latch_day += 1
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
        if giveback is not None and cur["peak"] >= gb_floor and cur["peak"] > 0 and bo <= cur["peak"] * (1 - gb_frac):
            hit = hit or "giveback"
        if max_days is not None and cur["age"] >= max_days:
            hit = hit or "max_days"
        if ps_level is not None and bo >= ps_level:
            thr = ps_thr * blvol[t] if ps_mode == "vol" else ps_thr
            move = BL[t] if ps_dir == "up" else abs(BL[t])
            if thr == thr and move >= thr:  # NaN-safe: no gate until BLVol has 10 days of history
                hit = hit or "swing"
        day1_ok = (hit == "stop" and BN[t] == 1 and
                   (first_day_stop == "all" or (first_day_stop == "fresh" and (t == 0 or BL[t - 1] == 0))))
        if hit and (BN[t] != 1 or day1_ok):  # workbook BQ is forced to 0 on a New Trade Start day: no exit on day 1
            latched = True
            latch_day = 0
            cur["reason"] = hit
            cur["exit_z"] = z[t] if z[t] == z[t] else None
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
