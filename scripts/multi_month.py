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
