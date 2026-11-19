"""Trade-level profit-take and profit-only month-start re-entry for the
Position Model's stop overlay (columns BL:BU), plus five helper columns EK:EO.

Mirrors exit_backtest.simulate(stop, profit, reentry_month="profit_only"):

- Profit-take: once a trade's running PnL (BO) reaches Config row 30, the
  trade exits at that day's close and stays flat until the next New Trade
  Start (BN), exactly like the stop. Neither can fire on a trade's first day.
- Re-entry: a trade that was exited by the profit-take re-opens on the first
  trading day of a new calendar month if the signal is still on (the trade is
  still economically active and that day's gross PnL is non-zero) and that
  day's Z-Score has the same sign as, and at least the magnitude of, the
  Z-Score on the exit day. The re-opened position is a fresh sub-trade: its
  running PnL, stop and profit-take start from zero, and it can exit on its
  first day. Switched on per month by Config row 31 = 1.

With Config row 30 blank (or 0) and row 31 not 1, every rewritten cell
computes exactly what the original formula did.
"""

from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS

FIRST_ROW = 21  # first steady-state row of the stop overlay (BOOT_END for BL:BY)

_ROW = {name: row for name, row, _default in CONFIG_FIELDS}

# Original formula shapes (row r, previous row p) the overlay replaces. BP's
# stop cell is written either relative or absolute in live workbooks; {stop}
# stands for either spelling.
ORIGINAL_SHAPES = {
    "BO": "IF(BM{r}=0,0,IF(BN{r}=1,BL{r},IFERROR(BO{p},0)+BL{r}))",
    "BP": "IF(AND(BM{r}=1,BO{r}<={stop},BO{p}>{stop}),1,0)",
    "BQ": "IF(BN{r}=1,0,IF(IFERROR(BQ{p},0)=1,1,BP{r}))",
    "BT": "IF(BS{r}=1,BL{r},IF(BQ{r}<>BQ{p},BL{r},0))",
    "BU": "IF(BS{r}=0,0,IF(BN{r}=1,BT{r},IFERROR(BU{p},0)+BT{r}))",
}

REWRITTEN_COLS = tuple(ORIGINAL_SHAPES)
HELPER_COLS = ("EK", "EL", "EM", "EN", "EO")

HEADERS = {
    "BP": "Exit Triggered Today (Stop or Profit-Take)",
    "BQ": "Exit State (Latched)",
    "BR": "Total Exited (Stop or Profit-Take)",
    "EK": "Month Start",
    "EL": "Exit Today",
    "EM": "Exit Z-Score",
    "EN": "Exit Was Profit-Take",
    "EO": "Re-Entry Today",
}


def config_refs(month: str) -> dict:
    """Absolute Config cells for `month`'s stop, profit-take and re-entry switch."""
    col = CONFIG_MONTH_COLUMNS[month]
    return {key: f"Config!${col}${_ROW[field]}" for key, field in
            (("stop", "StopLossUSD_Base"), ("tp", "ProfitTakeUSD"), ("reentry", "ReentryMonthStart"))}


def original_formula(col: str, month: str, row: int, absolute_stop: bool = True) -> str:
    """The pre-overlay formula for `col` on `row` (used to check a workbook
    before patching it)."""
    stop = config_refs(month)["stop"]
    if not absolute_stop:
        stop = stop.replace("$", "")
    return "=" + ORIGINAL_SHAPES[col].format(r=row, p=row - 1, stop=stop)


def overlay_formulas(month: str, row: int) -> dict:
    """{column: formula} for the overlay on `row` (row >= FIRST_ROW)."""
    r, p = row, row - 1
    c = config_refs(month)
    stop, tp, reentry = c["stop"], c["tp"], c["reentry"]
    return {
        "EK": f"=IF(MONTH(A{r})<>MONTH(A{p}),1,0)",
        "EO": (f"=IFERROR(IF(AND({reentry}=1,BN{r}<>1,IFERROR(BQ{p},0)=1,IFERROR(BM{p},0)=1,BL{r}<>0,"
               f"EK{r}=1,IFERROR(EN{p},0)=1,ISNUMBER(L{r}),ISNUMBER(EM{p}),"
               f"SIGN(L{r})=SIGN(EM{p}),ABS(L{r})>=ABS(EM{p})),1,0),0)"),
        "BO": f"=IF(BM{r}=0,0,IF(OR(BN{r}=1,EO{r}=1),BL{r},IFERROR(BO{p},0)+BL{r}))",
        "BP": (f"=IF(AND(BM{r}=1,OR(AND(BO{r}<={stop},IF(EO{r}=1,0,BO{p})>{stop}),"
               f"AND(N({tp})>0,BO{r}>={tp}))),1,0)"),
        "BQ": f"=IF(BN{r}=1,0,IF(EO{r}=1,BP{r},IF(IFERROR(BQ{p},0)=1,1,BP{r})))",
        "BT": f"=IF(BS{r}=1,BL{r},IF(OR(BQ{r}<>BQ{p},EO{r}=1),BL{r},0))",
        "BU": f"=IF(BS{r}=0,0,IF(OR(BN{r}=1,EO{r}=1),BT{r},IFERROR(BU{p},0)+BT{r}))",
        # Helper bookkeeping is wrapped so the warm-up rows' #N/A/#DIV/0! in
        # BL:BQ don't spread into EL:EN (the rewritten columns keep theirs).
        "EL": f"=IFERROR(IF(AND(BQ{r}=1,OR(IFERROR(BQ{p},0)=0,EO{r}=1)),1,0),0)",
        "EM": f'=IF(EL{r}=1,IF(ISNUMBER(L{r}),L{r},""),IF(IFERROR(BN{r},0)=1,"",IFERROR(EM{p},"")))',
        "EN": (f"=IFERROR(IF(EL{r}=1,IF(AND(N({tp})>0,BO{r}>={tp}),1,0),"
               f"IF(BN{r}=1,0,IFERROR(EN{p},0))),0)"),
    }
