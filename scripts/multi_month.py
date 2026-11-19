from openpyxl.utils import get_column_letter, column_index_from_string

MONTHS = ["January", "March", "May", "July", "August", "September", "October", "December"]

GAP = 2

DATA_BLOCK_WIDTH = 20
DIFFLAG_BLOCK_WIDTH = 8
JOHANSEN_MONTHLY_BLOCK_WIDTH = 11
MONTHLY_BETA_BLOCK_WIDTH = 8

MATRICES_BLOCK_WIDTH = 8
JOHANSEN_BLOCK_WIDTH = 6
RESULTS_BLOCK_WIDTH = 2
MJR_SUB_GAP = 2
MJR_MONTH_GAP = 4
MJR_BLOCK_WIDTH = (
    MATRICES_BLOCK_WIDTH + MJR_SUB_GAP + JOHANSEN_BLOCK_WIDTH + MJR_SUB_GAP + RESULTS_BLOCK_WIDTH
)

# Futures month code of each block's contract (Config's ZL/HOContractMonth).
CONTRACT_CODES = {"January": "F", "March": "H", "May": "K", "July": "N",
                  "August": "Q", "September": "U", "October": "V", "December": "Z"}

# 'Roll Prices' sheet, identical layout in the pricing workbook (formulas) and
# the Knox workbook (pasted values): one block per month, one row per roll year,
# holding the OLD contract's (and old D4 vintage's) prices on that roll day.
ROLL_PRICES_SHEET = "Roll Prices"
ROLL_PRICES_HEADERS = ["Year", "Roll Date", "Old Contract", "SBO Settle", "HO Settle", "D4", "SBO Open", "HO Open"]
ROLL_PRICES_BLOCK_WIDTH = len(ROLL_PRICES_HEADERS)
ROLL_PRICES_GAP = 1
ROLL_PRICES_FIRST_ROW = 3
ROLL_PRICES_YEARS = list(range(2021, 2036))
ROLL_PRICES_LAST_ROW = ROLL_PRICES_FIRST_ROW + len(ROLL_PRICES_YEARS) - 1


def roll_prices_col(month: str, header: str) -> str:
    """Column letter of `header` in `month`'s Roll Prices block."""
    offset = column_offset(month, ROLL_PRICES_BLOCK_WIDTH, ROLL_PRICES_GAP)
    return get_column_letter(ROLL_PRICES_HEADERS.index(header) + 1 + offset)


def month_index(month: str) -> int:
    return MONTHS.index(month)


def column_offset(month: str, block_width: int, gap: int = GAP) -> int:
    """How many columns to add to a block's original (single-month-template-
    style, starting at column A) column indices so that `month`'s copy of
    the block lands at the correct position, given every month's block is
    `block_width` columns wide with `gap` blank columns between consecutive
    months' blocks."""
    return month_index(month) * (block_width + gap)


def mjr_offsets(month: str) -> dict:
    """Column offsets (see column_offset) for each of the three sub-blocks
    making up `month`'s mega-block on the combined Matrices_Johansen_Results
    tab. Each offset is relative to that sub-block's own original
    column-A-start layout (Matrices A-H, Johansen A-F, Results A-B)."""
    mega_offset = column_offset(month, MJR_BLOCK_WIDTH, MJR_MONTH_GAP)
    matrices_offset = mega_offset
    johansen_offset = mega_offset + MATRICES_BLOCK_WIDTH + MJR_SUB_GAP
    results_offset = johansen_offset + JOHANSEN_BLOCK_WIDTH + MJR_SUB_GAP
    return {"Matrices": matrices_offset, "Johansen": johansen_offset, "Results": results_offset}


def shift_col(col_letter: str, delta: int) -> str:
    return get_column_letter(column_index_from_string(col_letter) + delta)
