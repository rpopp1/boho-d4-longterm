"""The Dashboard's "Performance Review" section: Month / Year / Metric
dropdowns, headline tiles, one Year x Month heat table, the per-month detail
table and the (collapsed) gross tier attribution. Every figure is an INDEX
into the Dashboard_Calc metric matrix (dashboard_metrics)."""

from openpyxl.formatting.rule import ColorScaleRule, Rule
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.worksheet.datavalidation import DataValidation

from dashboard_metrics import (
    ALL_YEARS, CALC_SHEET, CURRENCY_FMT, FIRST_YEAR, LATEST_YEAR, MATRIX_RANGE, METRIC_CELL,
    METRICS, MONTH_CELL, PADDED_LAST_ROW, SERIES, TABLE_YEARS, YEAR_CELL, YEAR_OPTIONS_COL,
    matrix_lookup, matrix_row, pm,
)
from dashboard_style import (
    BOLD, HEAT_CURRENCY_FMT, HEAT_MID, HEAT_NEG, HEAT_POS, SECTION_FONT, TILE_LABEL_FONT,
    TILE_VALUE_FONT, TITLE_FONT, section_title, style_header_row,
)
from multi_month import MONTHS

REVIEW_FIRST_ROW = 17
REVIEW_ANCHORS = {"controls": 18, "tiles": 20, "charts": 23, "year_table": 71, "detail": 86,
                  "tiers": 98}
TILE_METRICS = ["pnl", "sharpe", "sortino", "maxdd", "calmar", "winrate", "pf", "trades"]

MONTH_IDX = f"{CALC_SHEET}!$Z$1"
PERIOD_IDX = f"{CALC_SHEET}!$Z$2"
SEL_YEAR = f"{CALC_SHEET}!$Z$3"
METRIC_IDX = f"{CALC_SHEET}!$Z$6"
MIDPOINT = f"{CALC_SHEET}!$Z$7"

_KEYS = [k for k, _, _ in METRICS]
_LABEL = {k: label for k, label, _ in METRICS}
_FMT = {k: fmt for k, _, fmt in METRICS}
_TILE_FILL = PatternFill("solid", fgColor="F2F2F2")
_OUTLINE = Side(style="medium", color="1F3864")
TILE_MILLIONS_FMT = '$#,##0.00,,"M";($#,##0.00,,"M")'
BAR_AXIS_FMT = '[>=1000]#,##0;[<=-1000]-#,##0;0.00'


def _abs(cell: str) -> str:
    col = "".join(ch for ch in cell if ch.isalpha())
    return f"${col}${cell[len(col):]}"


def _heat_fmt(fmt: str) -> str:
    return HEAT_CURRENCY_FMT if fmt == CURRENCY_FMT else fmt


def _controls(ws, row: int) -> None:
    year_list = f"{CALC_SHEET}!${YEAR_OPTIONS_COL}$2:${YEAR_OPTIONS_COL}$14"
    specs = [
        ("A", MONTH_CELL, "Month", "Portfolio", '"' + ",".join(SERIES) + '"'),
        ("C", YEAR_CELL, "Year", ALL_YEARS, year_list),
        ("E", METRIC_CELL, "Metric", _LABEL["pnl"], '"' + ",".join(lb for _, lb, _ in METRICS) + '"'),
    ]
    for label_col, cell, label, default, formula1 in specs:
        ws[f"{label_col}{row}"] = label
        ws[f"{label_col}{row}"].font = BOLD
        ws[f"{label_col}{row}"].alignment = Alignment(horizontal="right")
        ws[cell] = default
        ws[cell].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[cell].font = BOLD
        ws[cell].alignment = Alignment(horizontal="center")
        dv = DataValidation(type="list", formula1=formula1, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[cell])
    ws[f"H{row}"] = f'={_abs(MONTH_CELL)}&" · "&{_abs(YEAR_CELL)}'
    ws[f"H{row}"].font = TILE_LABEL_FONT


def _tiles(ws, row: int) -> None:
    for c, key in enumerate(TILE_METRICS, start=1):
        k = _KEYS.index(key) + 1
        label = ws.cell(row=row, column=c, value=_LABEL[key])
        value = ws.cell(row=row + 1, column=c, value="=" + matrix_lookup(str(k), MONTH_IDX, PERIOD_IDX))
        label.font, value.font = TILE_LABEL_FONT, TILE_VALUE_FONT
        value.number_format = TILE_MILLIONS_FMT if _FMT[key] == CURRENCY_FMT else _FMT[key]
        for cell in (label, value):
            cell.fill = _TILE_FILL
            cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row + 1].height = 24


def _year_table(ws, row: int) -> int:
    ws.cell(row=row, column=1,
            value=f'="Calendar-Year "&{_abs(METRIC_CELL)}&" by Contract Month (Risk-Adjusted Basis)"'
            ).font = SECTION_FONT
    header = row + 1
    names = MONTHS + ["Portfolio"]
    ws.cell(row=header, column=1, value="Year")
    for c, name in enumerate(names, start=2):
        ws.cell(row=header, column=c, value=name)
    style_header_row(ws, header, 1, 1 + len(names))
    first = header + 1
    for j, year in enumerate(TABLE_YEARS):
        r = first + j
        ws.cell(row=r, column=1, value=f'=IF({year}<={LATEST_YEAR},{year},"")')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        for c, name in enumerate(names, start=2):
            s = SERIES.index(name) + 1
            ws.cell(row=r, column=c, value=(
                f'=IF($A{r}="","",INDEX({MATRIX_RANGE},{matrix_row(METRIC_IDX, str(s))},$A{r}-{FIRST_YEAR}+2))'))
    last = first + len(TABLE_YEARS) - 1
    body = f"B{first}:J{last}"

    ws.conditional_formatting.add(body, ColorScaleRule(
        start_type="min", start_color=HEAT_NEG,
        mid_type="formula", mid_value=MIDPOINT, mid_color=HEAT_MID,
        end_type="max", end_color=HEAT_POS))
    # Number format follows the Metric dropdown.
    for i, (_key, label, fmt) in enumerate(METRICS):
        dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=170 + i, formatCode=_heat_fmt(fmt)))
        ws.conditional_formatting.add(body, Rule(
            type="expression", formula=[f'{_abs(METRIC_CELL)}="{label}"'], dxf=dxf))
    # Outline the selected month column (header text = Month dropdown) and year row.
    border = DifferentialStyle(border=Border(left=_OUTLINE, right=_OUTLINE, top=_OUTLINE, bottom=_OUTLINE))
    ws.conditional_formatting.add(body, Rule(
        type="expression", formula=[f'AND($A{first}<>"",B${header}={_abs(MONTH_CELL)})'], dxf=border))
    ws.conditional_formatting.add(f"A{first}:J{last}", Rule(
        type="expression", formula=[f"$A{first}={_abs(YEAR_CELL)}"], dxf=border))
    return last


def _midpoint(calc_ws, body: str) -> None:
    """Heatmap midpoint per metric: 0 for PnL/ratios, 0.5 win rate, 1 profit
    factor, the table's median for the rest."""
    med = f"MEDIAN(Dashboard!{body})"
    mids = {"pnl": "0", "sharpe": "0", "sortino": "0", "calmar": "0", "maxdd": med,
            "winrate": "0.5", "avgwin": med, "avgloss": med, "pf": "1", "trades": med}
    calc_ws["Z7"] = f"=IFERROR(CHOOSE($Z$6,{','.join(mids[k] for k in _KEYS)}),0)"


def _detail_table(ws, row: int) -> int:
    ws.cell(row=row, column=1,
            value=f'="By Contract Month — "&IF({SEL_YEAR}=0,"All years",{SEL_YEAR})').font = SECTION_FONT
    header = row + 1
    ws.cell(row=header, column=1, value="Series")
    for c, (_, label, _) in enumerate(METRICS, start=2):
        ws.cell(row=header, column=c, value=label)
    style_header_row(ws, header, 1, 1 + len(METRICS))
    for i, series in enumerate(SERIES):
        r = header + 1 + i
        name = ws.cell(row=r, column=1, value=series)
        if series == "Portfolio":
            name.font = BOLD
        for c, (_key, _, fmt) in enumerate(METRICS, start=2):
            cell = ws.cell(row=r, column=c, value="=" + matrix_lookup(str(c - 1), str(i + 1), PERIOD_IDX))
            cell.number_format = _heat_fmt(fmt)
            cell.alignment = Alignment(horizontal="right")
    return header + len(SERIES)


def _tier_table(ws, row: int) -> int:
    section_title(ws, row, "Gross PnL Attribution by Position Tier (since inception, not stop-loss-adjusted)")
    header = row + 1
    for c, h in enumerate(["Month", "Tier 1 PnL", "Tier 2 PnL", "Tier 3 PnL", "Gross Total PnL"], start=1):
        ws.cell(row=header, column=c, value=h)
    style_header_row(ws, header, 1, 5)
    for i, month in enumerate(MONTHS):
        r = header + 1 + i
        ws.cell(row=r, column=1, value=month)
        for col, letter in ((2, "AZ"), (3, "BE"), (4, "BJ")):
            rng = f"{pm(month)}!${letter}$2:${letter}${PADDED_LAST_ROW}"
            ws.cell(row=r, column=col, value=f"=LOOKUP(2,1/ISNUMBER({rng}),{rng})").number_format = HEAT_CURRENCY_FMT
        ws.cell(row=r, column=5, value=f"=B{r}+C{r}+D{r}").number_format = HEAT_CURRENCY_FMT
    last = header + len(MONTHS)
    for r in range(header, last + 1):  # collapsed under the title; expand with the outline [+]
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].hidden = True
    return last


def build_review_section(ws, calc_ws) -> dict:
    a = REVIEW_ANCHORS
    ws.cell(row=REVIEW_FIRST_ROW, column=1, value="Performance Review").font = TITLE_FONT
    for c in range(1, 11):
        ws.cell(row=REVIEW_FIRST_ROW, column=c).border = Border(bottom=Side(style="thin", color="1F3864"))
    _controls(ws, a["controls"])
    _tiles(ws, a["tiles"])
    year_last = _year_table(ws, a["year_table"])
    first = a["year_table"] + 2
    _midpoint(calc_ws, f"$B${first}:$J${year_last}")
    _detail_table(ws, a["detail"])
    end = _tier_table(ws, a["tiers"])
    ws.print_area = f"A{REVIEW_FIRST_ROW}:J{year_last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return {**a, "year_last": year_last, "end": end}


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, Series  # noqa: E402
from openpyxl.chart.data_source import AxDataSource, NumRef  # noqa: E402
from openpyxl.chart.marker import Marker  # noqa: E402

FULL_WIDTH_CM = 30
BAR_HELPER_ROWS = range(2, 2 + len(MONTHS))  # Dashboard_Calc EW:EY rows 2-9
YEAR_LINE_GREYS = ["E0E0E0", "C8C8C8", "B0B0B0", "989898", "808080", "686868"]


def _named_series(ref, name: str, color: str, width: int) -> Series:
    s = Series(ref, title=None)
    s.val.numRef.f = f"{CALC_SHEET}!{name}"
    s.val.numRef.numCache = None
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = width
    s.smooth = False
    return s


def _date_axis(chart) -> None:
    for s in chart.series:
        s.cat = AxDataSource(numRef=NumRef(f=f"{CALC_SHEET}!Chart_SelDates"))
    chart.x_axis.number_format = "mmm yy"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend = None


def _build_bar_helpers(calc_ws) -> None:
    calc_ws["EW1"], calc_ws["EX1"], calc_ws["EY1"] = "Bar Month", "Other Months", "Selected Month"
    for i, month in enumerate(MONTHS):
        r = BAR_HELPER_ROWS[i]
        v = matrix_lookup("$Z$6", str(i + 2), "$Z$2")
        calc_ws[f"EW{r}"] = month
        calc_ws[f"EX{r}"] = f"=IF($Z$1={i + 2},NA(),{v})"
        calc_ws[f"EY{r}"] = f"=IF($Z$1={i + 2},{v},NA())"


def build_review_charts(ws, wb, anchor_row: int) -> None:
    calc_ws = wb[CALC_SHEET]
    placeholder = Reference(calc_ws, min_col=56, min_row=2, max_row=3)  # BD

    equity = LineChart()
    equity.title = "Equity curve (risk-adjusted cumulative PnL)"
    equity.y_axis.number_format = "$#,##0"
    equity.y_axis.majorGridlines.spPr = None
    eq = _named_series(placeholder, "Chart_SelEquity", "000000", 22000)
    eq.marker = Marker(symbol="none")
    equity.series.append(eq)
    _date_axis(equity)
    equity.width, equity.height = FULL_WIDTH_CM, 9
    ws.add_chart(equity, f"A{anchor_row}")

    dd = AreaChart()
    dd.title = "Drawdown from peak"
    dd.y_axis.number_format = "$#,##0"
    dds = _named_series(placeholder, "Chart_SelDD", "D55E00", 9525)
    dds.graphicalProperties.solidFill = "F4B183"
    dd.series.append(dds)
    _date_axis(dd)
    dd.x_axis.delete = True  # dates are on the equity chart directly above
    dd.width, dd.height = FULL_WIDTH_CM, 5
    ws.add_chart(dd, f"A{anchor_row + 18}")

    overlay = LineChart()
    overlay.title = "PnL by year (trading-day overlay)"
    overlay.y_axis.number_format = "$#,##0"
    overlay.display_blanks = "gap"
    overlay.x_axis.tickLblSkip = 20
    overlay.x_axis.tickMarkSkip = 20
    overlay.x_axis.delete = False
    overlay.y_axis.delete = False
    for j, color in enumerate(YEAR_LINE_GREYS):
        s = Series(Reference(calc_ws, min_col=18 + j, min_row=1, max_row=262), title_from_data=True)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 12700
        s.marker = Marker(symbol="none")
        s.smooth = False
        overlay.series.append(s)
    sel = Series(Reference(calc_ws, min_col=59, min_row=1, max_row=262), title_from_data=True)  # BG
    sel.graphicalProperties.line.solidFill = "000000"
    sel.graphicalProperties.line.width = 31750
    sel.marker = Marker(symbol="none")
    sel.smooth = False
    overlay.series.append(sel)
    overlay.set_categories(Reference(calc_ws, min_col=17, min_row=2, max_row=262))  # Q
    overlay.legend.position = "b"
    overlay.width, overlay.height = 15, 9
    ws.add_chart(overlay, f"A{anchor_row + 28}")

    _build_bar_helpers(calc_ws)
    bars = BarChart()
    bars.type = "col"
    bars.grouping = "clustered"
    bars.overlap = 100
    bars.gapWidth = 60
    bars.title = "Selected metric by contract month"
    for col, color in ((154, "BFBFBF"), (155, "0072B2")):  # EX, EY
        s = Series(Reference(calc_ws, min_col=col, min_row=1, max_row=1 + len(MONTHS)), title_from_data=True)
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.noFill = True
        bars.series.append(s)
    bars.set_categories(Reference(calc_ws, min_col=153, min_row=2, max_row=1 + len(MONTHS)))  # EW
    bars.x_axis.delete = False
    bars.y_axis.delete = False
    bars.legend = None
    bars.y_axis.number_format = BAR_AXIS_FMT
    bars.width, bars.height = 15, 9
    ws.add_chart(bars, f"F{anchor_row + 28}")
