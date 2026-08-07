from multi_month import (
    MONTHS, GAP, month_index, column_offset, mjr_offsets, shift_col,
    DATA_BLOCK_WIDTH, DIFFLAG_BLOCK_WIDTH,
)


def test_months_in_calendar_order():
    assert MONTHS == [
        "January", "March", "May", "July", "August", "September", "October", "December",
    ]


def test_month_index():
    assert month_index("January") == 0
    assert month_index("December") == 7


def test_column_offset_january_is_zero_for_every_block_width():
    # January must land exactly where today's single-month template already
    # puts things -- a load-bearing sanity property for every later task.
    assert column_offset("January", DIFFLAG_BLOCK_WIDTH, GAP) == 0
    assert column_offset("January", DATA_BLOCK_WIDTH, GAP) == 0


def test_column_offset_march_diff_lag():
    assert column_offset("March", DIFFLAG_BLOCK_WIDTH, GAP) == 10  # 1 * (8 + 2)


def test_column_offset_december_data():
    assert column_offset("December", DATA_BLOCK_WIDTH, GAP) == 154  # 7 * (20 + 2)


def test_mjr_offsets_january():
    assert mjr_offsets("January") == {"Matrices": 0, "Johansen": 10, "Results": 18}


def test_mjr_offsets_march():
    # March's mega-block starts at 1 * (20 + 4) = 24
    assert mjr_offsets("March") == {"Matrices": 24, "Johansen": 34, "Results": 42}


def test_shift_col():
    assert shift_col("B", 10) == "L"
    assert shift_col("E", 5) == "J"
