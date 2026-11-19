"""Roll-day pricing for the Position Model's price-change columns (Z:AF).

On a block's roll day the Data row already carries the NEW contract (and D4
vintage), while the previous row carries the OLD one, so a plain day-over-day
difference books the contract gap as PnL. A position is held through the
roll-day close in the old contract, so on that day every price change must be
taken entirely in the old contract: its settle and open come from the 'Roll
Prices' sheet (one row per roll day). On every other day the held price falls
back to the row's own Data price, so non-roll days compute exactly as before.

Only Z:AF change. Data, DiffLag, the Johansen sheets, Monthly Beta and the
spread/Z-score/position columns are untouched, so the cointegration test,
weights and signals are identical to the unfixed model.
"""

from multi_month import (
    MONTHS, CONTRACT_CODES, ROLL_PRICES_SHEET, ROLL_PRICES_HEADERS, ROLL_PRICES_FIRST_ROW,
    ROLL_PRICES_LAST_ROW, ROLL_PRICES_YEARS, roll_prices_col,
)

# Position Model columns rewritten by the fix, with the original formula
# shape each replaces (row r, previous row p) for reference.
PRICE_CHANGE_COLS = {
    "Z": "B{r}-B{p}",    # $ΔSBO Settle
    "AA": "C{r}-C{p}",   # $ΔHO Settle
    "AB": "B{r}-F{r}",   # $ΔSBO Same day Settle-Open
    "AC": "C{r}-G{r}",   # $ΔHO Same day Settle-Open
    "AD": "F{r}-B{p}",   # $ΔSBO Open-Prev. Settle
    "AE": "G{r}-C{p}",   # $ΔHO Open-Prev. Settle
    "AF": "E{r}-E{p}",   # $ΔD4 Settle
}


def _held(month: str, header: str, row: int, fallback_col: str, scale: str, file_syntax: bool) -> str:
    """Old-contract `header` price on a roll day for `month`, else the row's
    own `fallback_col` price. `scale` converts Roll Prices units to the
    Position Model's (SBO is stored in cents, the Position Model uses $/lb)."""
    sheet = f"'{ROLL_PRICES_SHEET}'"
    rng = lambda h: (f"{sheet}!${roll_prices_col(month, h)}${ROLL_PRICES_FIRST_ROW}:"
                     f"${roll_prices_col(month, h)}${ROLL_PRICES_LAST_ROW}")
    lookup = f"INDEX({rng(header)},MATCH($A{row},{rng('Roll Date')},0))"
    # Saved-file XML spells LET and its names with Excel's future-function
    # prefixes (as the source workbook's own LET formulas do); typed into
    # Excel through COM they're plain.
    let, held = ("_xlfn.LET", "_xlpm.held") if file_syntax else ("LET", "held")
    return f"{let}({held},IFERROR({lookup},0),IF(N({held})>0,{held}{scale},{fallback_col}{row}))"


def write_roll_prices_sheet(wb) -> None:
    """Empty 'Roll Prices' sheet for a freshly built template (openpyxl):
    labels, headers and roll years only. Its prices are pasted from the
    pricing workbook's sheet of the same name and layout; while blank, every
    Position Model row falls back to its plain Data prices."""
    ws = wb[ROLL_PRICES_SHEET] if ROLL_PRICES_SHEET in wb.sheetnames else wb.create_sheet(ROLL_PRICES_SHEET)
    for month in MONTHS:
        ws[f"{roll_prices_col(month, 'Year')}1"] = f"{month} ({CONTRACT_CODES[month]}): old contract on its roll day"
        for h in ROLL_PRICES_HEADERS:
            ws[f"{roll_prices_col(month, h)}2"] = h
        for i, year in enumerate(ROLL_PRICES_YEARS):
            ws[f"{roll_prices_col(month, 'Year')}{ROLL_PRICES_FIRST_ROW + i}"] = year


def price_change_formulas(month: str, row: int, file_syntax: bool = False) -> dict:
    """{column: formula} for the roll-aware Z:AF cells of `month`'s Position
    Model on `row` (row >= 3; row 2 has no previous day and is left as is).
    `file_syntax=True` for openpyxl-written workbooks, False for Excel COM."""
    p = row - 1
    sbo_set = _held(month, "SBO Settle", row, "B", "/100", file_syntax)
    ho_set = _held(month, "HO Settle", row, "C", "", file_syntax)
    sbo_open = _held(month, "SBO Open", row, "F", "/100", file_syntax)
    ho_open = _held(month, "HO Open", row, "G", "", file_syntax)
    d4 = _held(month, "D4", row, "E", "", file_syntax)
    return {
        "Z": f"={sbo_set}-B{p}",
        "AA": f"={ho_set}-C{p}",
        "AB": f"={sbo_set}-{sbo_open}",
        "AC": f"={ho_set}-{ho_open}",
        "AD": f"={sbo_open}-B{p}",
        "AE": f"={ho_open}-C{p}",
        "AF": f"={d4}-E{p}",
    }
