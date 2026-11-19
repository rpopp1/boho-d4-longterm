"""The Dashboard's "Today" section: one row per contract month showing the
position the strategy actually holds (signal gated by the stop/profit-take
overlay), its Z-score, the next signal and how far away it is, and the
volatility regime.

Each month's latest values are pulled into a small block in Dashboard_Calc
(TODAY_FIELDS, rows 2-9 in MONTHS order); the Dashboard's text formulas read
that block, so they can be tested on plain input cells in real Excel."""

from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill

from config_map import CONFIG_MONTH_COLUMNS
from dashboard_metrics import CALC_SHEET, CURRENCY_FMT, PADDED_LAST_ROW, pm
from dashboard_style import (
    BOLD, DIVERGING_NEG, HEAT_MID, HEAT_NEG, HEAT_POS, SECTION_FONT, style_header_row,
)
from multi_month import MONTHS

TODAY_FIRST_CALC_ROW = 2

# field -> Dashboard_Calc column
TODAY_FIELDS = {
    "asof": "CA", "row": "CB", "pos": "CC", "active": "CD", "z": "CE",
    "pt_exit": "CF", "exit_z": "CG", "regime": "CH", "pctl": "CI", "vol": "CJ",
    "shock": "CK", "trade_pnl": "CL",
    "s1": "CM", "s2": "CN", "s3": "CO", "l3": "CP", "l2": "CQ", "l1": "CR",
    "reentry": "CS", "t1": "CT", "t2": "CU", "t3": "CV",
    "exit_short": "CW", "exit_long": "CX",
}

# field -> Position Model column read on the month's last dated row
_PM_COLUMNS = {
    "pos": "V", "active": "BS", "z": "L", "pt_exit": "EN", "exit_z": "EM",
    "regime": "DC", "pctl": "DB", "vol": "CY", "shock": "DA", "trade_pnl": "BU",
    "s1": "M", "s2": "N", "s3": "O", "l3": "P", "l2": "Q", "l1": "R",
}
# field -> Config row
_CONFIG_ROWS = {"reentry": 31, "t1": 14, "t2": 15, "t3": 16, "exit_short": 17}

HEADERS = ["Month", "Position", "Signal", "Z-Score", "Regime", "Vol Pctl", "EWMA Vol",
           "Shock", "Open Trade PnL", "Next Signal"]


def build_today_block(ws) -> None:
    for field, col in TODAY_FIELDS.items():
        ws[f"{col}1"] = f"Today: {field}"
    for i, month in enumerate(MONTHS):
        r = TODAY_FIRST_CALC_ROW + i
        dates = f"{pm(month)}!$A$2:$A${PADDED_LAST_ROW}"
        ws[f"CA{r}"] = f"=LOOKUP(2,1/({dates}>DATE(2015,1,1)),{dates})"
        ws[f"CB{r}"] = f"=MATCH($CA{r},{dates},0)"
        for field, src in _PM_COLUMNS.items():
            ws[f"{TODAY_FIELDS[field]}{r}"] = (
                f"=INDEX({pm(month)}!${src}$2:${src}${PADDED_LAST_ROW},$CB{r})")
        for field, cfg_row in _CONFIG_ROWS.items():
            ws[f"{TODAY_FIELDS[field]}{r}"] = f"=Config!${CONFIG_MONTH_COLUMNS[month]}${cfg_row}"
        ws[f"CX{r}"] = f"={pm(month)}!$DV$7"


def position_formula(refs: dict) -> str:
    pos, active = refs["pos"], refs["active"]
    return f'=IF({pos}=0,"Flat",IF({active}=1,IF({pos}>0,"Long","Short"),"Flat (exited)"))'


def next_signal_formula(refs: dict) -> str:
    g = refs
    t = lambda x: f'TEXT({x},"0.00")'  # noqa: E731
    now = f'" (now "&{t(g["z"])}&")"'
    min_tier = f'MIN({g["t1"]},{g["t2"]},{g["t3"]})'
    short_next = f'MIN(IF({g["s1"]}=0,{g["t1"]},99),IF({g["s2"]}=0,{g["t2"]},99),IF({g["s3"]}=0,{g["t3"]},99))'
    long_next = f'MIN(IF({g["l1"]}=0,{g["t1"]},99),IF({g["l2"]}=0,{g["t2"]},99),IF({g["l3"]}=0,{g["t3"]},99))'
    short_add = f'IF({short_next}<99," · adds at Z ≥ "&{t(short_next)},"")'
    long_add = f'IF({long_next}<99," · adds at Z ≤ "&{t("-" + long_next)},"")'
    flat = f'IF({g["z"]}>=0,"Short at Z ≥ "&{t(min_tier)},"Long at Z ≤ "&{t("-" + min_tier)})&{now}'
    in_trade = (f'IF({g["pos"]}<0,"Exit at Z ≤ "&{t(g["exit_short"])}&{now}&{short_add},'
                f'"Exit at Z ≥ "&{t(g["exit_long"])}&{now}&{long_add})')
    reentry_date = f'TEXT(WORKDAY(EOMONTH({g["asof"]},0),1),"m/d")'
    reentry = (f'"Flat, profit-take exit. Re-entry "&{reentry_date}&" if Z "&'
               f'IF({g["exit_z"]}<0,"≤ ","≥ ")&{t(g["exit_z"])}&{now}')
    reset = (f'IF({g["pt_exit"]}=1,"Flat, profit-take exit. ","Flat, stopped out. ")&'
             f'"Waiting for signal reset at Z "&'
             f'IF({g["pos"]}<0,"≤ "&{t(g["exit_short"])},"≥ "&{t(g["exit_long"])})&{now}')
    return (f'=IF({g["pos"]}=0,{flat},IF({g["active"]}=1,{in_trade},'
            f'IF(AND({g["pt_exit"]}=1,{g["reentry"]}=1),{reentry},{reset})))')


def _calc_refs(i: int) -> dict:
    r = TODAY_FIRST_CALC_ROW + i
    return {k: f"{CALC_SHEET}!${col}${r}" for k, col in TODAY_FIELDS.items()}


def build_today_section(ws, first_row: int = 4) -> int:
    """Title row, header row, one row per month, Portfolio row. Returns the last row."""
    ws.cell(row=first_row, column=1, value="Today").font = SECTION_FONT
    header = first_row + 1
    for c, h in enumerate(HEADERS, start=1):
        ws.cell(row=header, column=c, value=h)
    style_header_row(ws, header, 1, len(HEADERS))
    first = header + 1
    for i, month in enumerate(MONTHS):
        r, g = first + i, _calc_refs(i)
        values = [
            (month, None),
            (position_formula(g), None),
            (f"={g['pos']}", "0.00;-0.00;0"),
            (f"={g['z']}", "0.00"),
            (f"={g['regime']}", None),
            (f"={g['pctl']}", "0%"),
            (f"={g['vol']}", "0.00"),
            (f'=IF({g["shock"]}=1,"⚡ shock","")', None),
            (f'=IF({g["active"]}=1,{g["trade_pnl"]},"")', CURRENCY_FMT),
            (next_signal_formula(g), None),
        ]
        for c, (v, fmt) in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if fmt:
                cell.number_format = fmt
            if c in (2, 5, 8):
                cell.alignment = Alignment(horizontal="center")
    last = first + len(MONTHS) - 1
    port = last + 1
    pos_rng, reg_rng = f"$B${first}:$B${last}", f"$E${first}:$E${last}"
    ws.cell(row=port, column=1, value="Portfolio").font = BOLD
    ws.cell(row=port, column=2, value=(
        f'=COUNTIF({pos_rng},"Long")&" long · "&COUNTIF({pos_rng},"Short")&" short · "&'
        f'({len(MONTHS)}-COUNTIF({pos_rng},"Long")-COUNTIF({pos_rng},"Short"))&" flat"'))
    ws.cell(row=port, column=5, value=(
        f'=(COUNTIF({reg_rng},"High")+COUNTIF({reg_rng},"Extreme"))&" of {len(MONTHS)} High/Extreme"'))
    total = ws.cell(row=port, column=9, value=f"=SUM(I{first}:I{last})")
    total.number_format = CURRENCY_FMT
    for c in range(1, len(HEADERS) + 1):
        ws.cell(row=port, column=c).font = BOLD

    ws.conditional_formatting.add(f"D{first}:D{last}", ColorScaleRule(
        start_type="num", start_value=-2.5, start_color=HEAT_NEG,
        mid_type="num", mid_value=0, mid_color=HEAT_MID,
        end_type="num", end_value=2.5, end_color=HEAT_POS))
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(
        operator="equal", formula=['"Extreme"'],
        fill=PatternFill("solid", fgColor=DIVERGING_NEG), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(
        operator="equal", formula=['"High"'], fill=PatternFill("solid", fgColor="F4C7A8")))
    return port
