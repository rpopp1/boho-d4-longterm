"""Palette, number formats and small styling helpers shared by the Dashboard builders."""

from openpyxl.styles import Alignment, Font, PatternFill


# Okabe-Ito colorblind-safe categorical set (7 hues) plus one extra
# well-separated hue for the 8th month; black is reserved for the Portfolio
# aggregate line so it reads as "the reference line," not just another month.
MONTH_COLORS = {
    "January": "E69F00",
    "March": "56B4E9",
    "May": "009E73",
    "July": "F0E442",
    "August": "0072B2",
    "September": "D55E00",
    "October": "CC79A7",
    "December": "6F4E37",
}
PORTFOLIO_COLOR = "000000"


# One diverging pair (blue=positive / vermillion=negative / white midpoint),
# reused for every polarity encoding on this tab (Z-score color scale,
# position data bars, PnL heatmap) for a consistent mental model.
DIVERGING_POS = "0072B2"
DIVERGING_NEG = "D55E00"
DIVERGING_MID = "FFFFFF"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FONT = Font(bold=True, size=12)
BOLD = Font(bold=True)

CURRENCY_FMT = '$#,##0;[RED]($#,##0)'
RATIO_FMT = '0.00'
PCT_FMT = '0.0%'

# Heatmap tints: light enough that black text stays readable on every cell.
HEAT_CURRENCY_FMT = '$#,##0;($#,##0)'
HEAT_POS = "9DC3E6"
HEAT_NEG = "F4B183"
HEAT_MID = "FFFFFF"

TITLE_FONT = Font(bold=True, size=18)
TILE_LABEL_FONT = Font(size=9, color="595959")
TILE_VALUE_FONT = Font(bold=True, size=16)


def style_header_row(ws, row: int, first_col: int, last_col: int) -> None:
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def section_title(ws, row: int, text: str) -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT

