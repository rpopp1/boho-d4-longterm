import re

import openpyxl
from openpyxl.utils import column_index_from_string

from sheet_cloner import fill_sheet_from_pattern
from config_map import CONFIG_FIELDS, config_cell, constant_map_for_month
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP
from roll_fix import price_change_formulas
from profit_take import FIRST_ROW as OVERLAY_FIRST_ROW, HEADERS as OVERLAY_HEADERS, overlay_formulas
from formula_tools import (
    extract_row_formulas,
    rewrite_constants,
    shift_formula_columns,
    shift_formula_rows,
    unwrap_scalar_array_formula,
)

# Config-sheet ROW per field, derived from the single source of truth rather
# than hardcoded here, so a row change in config_map can never silently point
# _apply_targeted_fix_round_2_overrides's BN6/DQ10 at the wrong field.
_CONFIG_ROW = {name: row for name, row, _default in CONFIG_FIELDS}

# ---------------------------------------------------------------------------
# Independent verification (see task-12-report.md for full methodology and
# output) found that Position Model's row-3 fill pattern is NOT uniformly
# representative across all ~130 columns, beyond the two failure categories
# already named in the task brief (spill anchors -- confirmed zero present in
# rows 1-5; fixed unanchored large-row references like DO10's C1266).
#
# A systematic column-by-column audit (comparing what naive row-shifting from
# row 3 predicts against the LIVE workbook's actual formulas at rows
# 300/600/900/1200/1380, then re-verified across every populated row) found
# three distinct column categories:
#
#   1. ROW3_OK (~45 columns): row 3's formula genuinely shifts correctly for
#      every later row. No special handling needed -- these are exactly what
#      the brief's naive fill_sheet_from_pattern(pattern_row=3) call handles.
#
#   2. BOOTSTRAP columns (~41 columns, e.g. J, K, AG-AL, AP-AU, AV/AW/BA/BB/
#      BF/BG, BC/BD/BH/BI, BL-BY, CD, CW-DC, BZ/CA/CB/CC): these implement
#      rolling-window statistics (EWMA vol, drawdown tracking, expanding
#      averages) whose formula literally changes SHAPE during a short
#      "bootstrap" phase (an expanding window anchored near the top of the
#      data) before settling into the true steady-state sliding-window
#      formula used for the rest of the sheet's life. Row 3 sits inside this
#      bootstrap phase for these columns, so naively shifting it down for
#      3000 rows would silently corrupt the EWMA vol regime detection and
#      drawdown tracking with an ever-expanding (instead of rolling) window.
#      BC specifically also contains the exact "fixed unanchored large-row
#      reference" failure pattern called out in the brief (row 3 = "=IF(
#      $S1268=0,0,-$S1268*$AF1269*$DF$5)", referencing a fixed calibration
#      row 1268/1269 with no $ on the row -- structurally identical to the
#      DO10/C1266 case). Copying these bootstrap rows literally (unshifted)
#      is the correct fix for both problems at once: it reproduces the fixed
#      reference verbatim (correct, since it's not meant to move) AND avoids
#      treating the non-representative early-window formula as the steady
#      state pattern.
#
#   3. PARAM_BLOCK columns (~39 columns, e.g. CE-CJ, CQ-CU, DD-EE): these are
#      not per-day data at all -- they are a static side panel of one-off
#      calibration/summary cells (matching the row-1-11 parameter block
#      already called out in the brief) that only occupies the first ~30
#      rows and is blank for the rest of the sheet's ~1380 live rows. These
#      are copied literally, never shifted/extended down 3000 rows.
#
# BOOT_END[col] = the first row at which column `col`'s formula becomes the
# true, indefinitely-shiftable steady-state pattern (verified by shifting
# that row's formula and comparing -- exactly, whitespace-normalized -- to
# every later populated row in the live workbook up to row 1386).
#
# NOTE (fix round 1): DZ/EA/EB/EC/ED are DELIBERATELY EXCLUDED from
# PARAM_BLOCK_COLS below, even though the live source's parameter-block area
# includes them. Those five columns hold the Dec-specific hardcoded per-year
# breakdown table (DZ3:DZ9 = literal years 2021-2026 + "Total"; EA4:EA8 =
# literal fixed-row-range COUNTs like "=COUNT($AH$250:$AH$494)" for 2022,
# etc.) -- exactly the hardcoded-row-bound pattern this task must eliminate.
# A first version of this file put them in PARAM_BLOCK_COLS, which literally
# copied that whole stale table into the built sheet; _write_dynamic_year_
# breakdown only overwrote rows 2-3 afterward, silently leaving rows 4-9 of
# DZ/EA/EB/EC/ED as leftover Dec-specific literals. Excluding them here means
# the baseline fill_sheet_from_pattern call (ROW3_OK default) writes into
# them instead, and _write_dynamic_year_breakdown explicitly clears
# rows 4..last_target_row after writing its dynamic row 3, so nothing
# Dec-specific can survive in these five columns anywhere in the built sheet.
# ---------------------------------------------------------------------------
BOOT_END = {
    "J": 41, "K": 40, "AG": 20, "AH": 20, "AI": 21, "AJ": 21, "AK": 21, "AL": 21,
    "AP": 21, "AQ": 21, "AR": 21, "AS": 21, "AT": 21, "AU": 21, "AX": 20,
    "BC": 20, "BD": 20, "BH": 20, "BI": 20,
    "BL": 21, "BM": 21, "BN": 21, "BO": 21, "BP": 21, "BQ": 21, "BR": 21,
    "BS": 21, "BT": 21, "BU": 21, "BV": 21, "BW": 21, "BX": 21, "BY": 21,
    "CD": 21, "CW": 21, "CX": 22, "CY": 22, "CZ": 22, "DA": 22, "DB": 22, "DC": 22,
    "AV": 21, "AW": 21, "BA": 21, "BB": 21, "BF": 21, "BG": 21,
    # NOTE (targeted-fix-round-2, finding 1): BZ/CA/CB/CC used to be here as
    # {"BZ": 741, ...} -- Dec's own live file is genuinely blank for these
    # columns until row 740 (where "740"/"741" is simply wherever calendar
    # year 2024 happens to start in Dec's specific historical data; see the
    # "24-26 Cumulative PnL" family of headers). That is a Dec-specific
    # calendar boundary, not a property of the strategy shared by every book
    # (unlike the genuine ~20-row rolling-window warm-up every other
    # BOOTSTRAP column above needs), so it does not belong in a general
    # template's BOOT_END table. These four are now handled by
    # _write_generalized_trailing_window_accum_cols() below instead, which
    # reuses the exact formula SHAPE Dec's row 741 shows (confirmed
    # byte-identical, modulo column letters, to BV/BW/BX/BY's all-time
    # Risk-Adjusted equivalent) but anchors it to row 2/3 like every other
    # ROW3_OK accumulator, instead of Dec's row 740/741.
    #
    # CX ("EWMA Variance", =$DQ$7*CX{r-1}+(1-$DQ$7)*(CV{r}^2)) IS also a
    # self-referencing recursion and stays out of ACCUM_SELF_REF_COLS
    # deliberately -- not an oversight. It differs from the PnL/drawdown
    # accumulators below in a way that makes IFERROR(prev,0) actively wrong
    # rather than merely unnecessary: those accumulators are literal running
    # totals, where "the world owes you nothing extra during an error row" is
    # exactly captured by re-seeding to 0. An EWMA *variance* seeded to 0
    # means "zero volatility", which downstream z-score/vol-based sizing
    # logic (ZThreshold_Tier*, EWMA_Lambda-driven sizing) could read as a
    # spurious signal (e.g. dividing by a near-zero vol estimate) rather than
    # a harmless "no PnL yet". CV (=D{r}-D{r-1}, a plain price/spread
    # difference) is in practice never itself an Excel error once D exists,
    # so CX's self-reference has negligible real-world risk of ever needing
    # the self-heal Y/AZ/etc. rely on -- left unwrapped on purpose.
}

PARAM_BLOCK_COLS = {
    "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO",
    "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY",
}

# MAJOR 4: CE-CJ ("RA 2021 Cumulative PnL" ... "RA 2026 Cumulative PnL") and
# CQ-CU ("2021 Cumulative PnL" ... "2025 Cumulative PnL") used to be in
# PARAM_BLOCK_COLS and literally copied from Dec. Inspection of the live
# workbook showed these are NOT one-off calibration cells at all -- each
# column is Dec's real per-calendar-year PnL total, keyed by a literal fixed
# row offset into CP/CD (e.g. CR3='=CP250', CS3='=CP495' -- 250/495 are simply
# where 2022/2023 happen to start in Dec's specific data; CF3='=CD487' is the
# RA equivalent). This is exactly the "hardcoded per-year columns requiring a
# new column every January" defect this project exists to eliminate -- one
# column per year, with Dec-specific row numbers baked into the formula, and
# no room for a 7th year (2026 has no RA column, 2026 CQ-column doesn't
# exist) without editing the sheet by hand.
#
# These are now decommissioned entirely and replaced by genuine SUMIFS-based
# per-year totals in the DZ:EF dynamic year-breakdown table below (EE/EF),
# keyed on this sheet's own Year column (CO) against the real per-day PnL
# columns (BK = "Grand Total PnL", BT = "Effective Daily PnL" i.e. the RA
# equivalent) -- no hardcoded row bounds, works for any number of years.
DECOMMISSIONED_PER_YEAR_COLS = ("CE", "CF", "CG", "CH", "CI", "CJ", "CQ", "CR", "CS", "CT", "CU")

# Columns holding the Dec-specific hardcoded per-year breakdown table that
# _write_dynamic_year_breakdown replaces. See the NOTE above.
# targeted-fix-round-2, finding 2: EG/EH/EI (Peak/Trough/Range) added --
# the genuine per-year equivalent of Dec's decommissioned DJ224:DN227
# MAX/MIN-over-CQ-CU cells. See _apply_targeted_fix_round_2_overrides() and
# _write_dynamic_year_breakdown() for the full story.
YEAR_BREAKDOWN_COLS = ("DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI")

# How many calendar-year rows the dynamic year-breakdown table enumerates
# before its Total row (generous headroom, matching this project's general
# "well over a decade" philosophy for row capacity elsewhere).
NUM_YEAR_BREAKDOWN_ROWS = 15

# CRITICAL 1: self-referencing accumulator/drawdown columns (e.g.
# Y{row} = Y{row-1}+X{row}). During the strategy's ~20-row warm-up window the
# per-row contribution (X, etc.) is naturally #N/A/#DIV/0! (not enough history
# yet for the rolling z-score). Dec's live file has a literal hardcoded 0 at
# specific warm-up rows (e.g. Y20) that breaks the #N/A chain once; the
# general formula this template writes everywhere has no such literal, so an
# error in the per-row contribution would otherwise propagate through the
# self-reference forever. Wrapping the self-reference in IFERROR(prev,0) makes
# the running total self-heal as soon as the per-row contribution is a real
# number again, matching Dec's own literal-0 patch behavior without hardcoding
# any Dec-specific row numbers.
#
# CM (=CK{row}-CL{row}) and BX (=BV{row}-BW{row}) are deliberately excluded:
# they are not self-referencing accumulators, they derive directly from other
# columns in the same row, so once CK/CL/BV/BW are hardened here they resolve
# correctly with no wrap of their own.
#
# Follow-on catch via CRITICAL 2's widened validation harness (which is
# exactly what it exists to do): CP/CD ("Calendar Year Cumulative PnL"/"RA ..."),
# and BM/BO/BQ/BU (an interlocking trade-state-machine cluster alongside
# BL/BN/BP/BR/BS/BT/BU -- "Trade Active", "Trade-Level PnL", "Stop State
# (Latched)", "Effective Trade Cumulative PnL") are ALSO self-referencing in
# exactly the same failure shape (e.g. BO{row}=IF(...,BO{row-1}+BL{row}),
# CP{row}=IF(CO{row}=CO{row-1},CP{row-1}+BK{row},0)) and were not in the
# original ~12-column enumeration, but suffer the identical permanent-error
# propagation once a #N/A/#DIV/0! enters the chain. BL/BN/BP/BR/BS/BT are NOT
# self-referencing themselves (they derive from BM/BO/BQ/BU/BL in the same
# row) and resolve automatically once those four are hardened.
#
# targeted-fix-round-2, finding 1: BZ/CA/CC ("24-26 Cumulative PnL"/"24-26
# Running Equity Peak"/"24-26 Max Drawdown", now generalized to an
# unconditional since-inception accumulator by
# _write_generalized_trailing_window_accum_cols() below) are the same
# self-referencing shape and were missed entirely by the original ~12-column
# enumeration -- they leaked Dec's un-wrapped literal formulas unmodified.
# CB (=BZ{row}-CA{row}) is deliberately excluded, same reasoning as CM/BX: it
# derives directly from BZ/CA in the same row, not from itself.
ACCUM_SELF_REF_COLS = (
    "Y", "AZ", "BE", "BJ", "CK", "CL", "CN", "BV", "BW", "BY",
    "CP", "CD", "BM", "BO", "BQ", "BU",
    "BZ", "CA", "CC",
)

# targeted-fix-round-2, finding 1: the four columns written by
# _write_generalized_trailing_window_accum_cols() below. Shared as a module-
# level constant so validate_against_dec.py's EXCLUDED_COMPARE_COLUMNS can
# import it directly (rather than re-typing the same four letters) -- these
# are now structurally redefined (since-inception, not Dec's row-740 trailing
# window), so they can never be expected to match Dec's real replayed values.
GENERALIZED_TRAILING_WINDOW_ACCUM_COLS = ("BZ", "CA", "CB", "CC")


def write_position_model_sheet(wb, source_path: str, month: str, num_rows: int = 3000, max_col: int = None) -> None:
    sheet_name = f"Position Model - {month}"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    constant_map = constant_map_for_month(month)

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Position Model"]
        if max_col is None:
            max_col = src.max_column

        # Header row, copied as-is (labels only, no formulas to rewrite).
        for cell in next(src.iter_rows(min_row=1, max_row=1, max_col=max_col)):
            if cell.value is not None:
                ws.cell(row=1, column=cell.column, value=cell.value)

        last_target_row = num_rows + 1  # row 2 (1 row) + rows 3..last_target_row (num_rows-1 rows)

        # Baseline per-day fill: generic row-3 shift across every column. This
        # is correct for the ~45 ROW3_OK columns and gets overwritten below
        # for the BOOTSTRAP/PARAM_BLOCK columns where row 3 is not
        # representative of the steady-state pattern.
        if num_rows > 1:
            fill_sheet_from_pattern(
                target_ws=ws,
                source_path=source_path,
                source_sheet="Position Model",
                pattern_row=3,
                target_start_row=3,
                num_rows=num_rows - 1,
                max_col=max_col,
                constant_map=constant_map,
            )

        # Row 2 filled from its own literal Dec pattern (boundary row, not
        # shifted from row 3).
        row2 = extract_row_formulas(source_path, "Position Model", row=2, max_col=max_col)
        for col_letter, value in row2.items():
            if isinstance(value, str) and value.startswith("="):
                value = rewrite_constants(value, constant_map)
            ws.cell(row=2, column=column_index_from_string(col_letter), value=value)

        # Fix up BOOTSTRAP columns: rows 3..(boot_end-1) copied literally
        # (unshifted) from the source, rows boot_end..last_target_row filled
        # by shifting the source's boot_end-row formula.
        for col_letter, boot_end in BOOT_END.items():
            col_idx = column_index_from_string(col_letter)
            if col_idx > max_col:
                continue
            literal_end = min(boot_end - 1, last_target_row, src.max_row)
            for row in range(3, literal_end + 1):
                cell = src.cell(row=row, column=col_idx)
                # MINOR 6: this literal-copy loop has no dynamic-array/spill
                # handling of its own -- unwrap any single-cell (legacy CSE)
                # ArrayFormula to plain formula text before it can be written
                # into the new sheet as an ArrayFormula object.
                value = unwrap_scalar_array_formula(cell.value, cell.coordinate)
                if isinstance(value, str) and value.startswith("="):
                    value = rewrite_constants(value, constant_map)
                # CRITICAL 3-class fix: ws.cell(row, column, value=None) is a
                # no-op in openpyxl (only assigns when value is not None) --
                # must set .value directly, or wherever the source's
                # bootstrap-window cell is genuinely blank (e.g. BC6:BC19),
                # whatever the earlier baseline fill_sheet_from_pattern
                # shift-fill pass already wrote there (a nonsensical shifted
                # version of BC's fixed-row-1268/1269 pattern) survives
                # underneath instead of being cleared. Found via CRITICAL 2's
                # widened validation harness.
                ws.cell(row=row, column=col_idx).value = value

            if boot_end <= last_target_row:
                pattern_cell = src.cell(row=boot_end, column=col_idx)
                # Same ArrayFormula unwrap as above -- this specific spot is
                # where Position Model!DC22 leaked in practice: DC's boot_end
                # row was an ArrayFormula, so `isinstance(pattern_value, str)`
                # was False and it fell into the else branch below, which
                # wrote the raw ArrayFormula object at boot_end only.
                pattern_value = unwrap_scalar_array_formula(pattern_cell.value, pattern_cell.coordinate)
                if isinstance(pattern_value, str) and pattern_value.startswith("="):
                    for row in range(boot_end, last_target_row + 1):
                        delta = row - boot_end
                        shifted = shift_formula_rows(pattern_value, delta)
                        rewritten = rewrite_constants(shifted, constant_map)
                        ws.cell(row=row, column=col_idx, value=rewritten)
                else:
                    ws.cell(row=boot_end, column=col_idx, value=pattern_value)

        # Fix up PARAM_BLOCK columns: static side-panel cells, copied literally
        # from the source wherever it has content, never shifted or extended.
        #
        # CRITICAL 3 fix: this loop used to `continue` (skip) wherever the
        # source cell was None, which meant the target cell retained whatever
        # the earlier baseline fill_sheet_from_pattern shift-fill pass had
        # already written there -- these columns are only genuinely populated
        # in the source's first ~30 rows (or, for a couple, up to ~230 rows),
        # so every row beyond that leaked thousands of shifted-formula/literal
        # cells per column. Setting .value directly (not skipping on None)
        # clears that leftover baseline-fill garbage, and iterating the full
        # row range up to last_target_row (not capped at src.max_row) ensures
        # rows beyond the source's own extent are cleared too.
        for col_letter in PARAM_BLOCK_COLS:
            col_idx = column_index_from_string(col_letter)
            if col_idx > max_col:
                continue
            for row in range(3, last_target_row + 1):
                cell = src.cell(row=row, column=col_idx)
                # MINOR 6: unwrap any single-cell ArrayFormula before writing
                # (this is where Position Model!DE43/DF43/DF46 leaked in
                # practice -- all three are legacy CSE array formulas stored
                # in DE/DF, both PARAM_BLOCK_COLS columns).
                value = unwrap_scalar_array_formula(cell.value, cell.coordinate)
                if isinstance(value, str) and value.startswith("="):
                    value = rewrite_constants(value, constant_map)
                ws.cell(row=row, column=col_idx).value = value

        # MAJOR 4 / CRITICAL 3: fully decommission the hardcoded per-year PnL
        # columns (see DECOMMISSIONED_PER_YEAR_COLS above) -- clear every row,
        # including the header row's Dec-specific year labels, so nothing
        # about them survives anywhere in the built sheet. Their real numbers
        # now live in the DZ:EF dynamic year-breakdown table instead.
        for col_letter in DECOMMISSIONED_PER_YEAR_COLS:
            col_idx = column_index_from_string(col_letter)
            if col_idx > max_col:
                continue
            for row in range(1, last_target_row + 1):
                ws.cell(row=row, column=col_idx).value = None
    finally:
        source_wb.close()

    _write_generalized_trailing_window_accum_cols(ws, last_target_row, max_col)
    _apply_targeted_fix_round_2_overrides(ws, last_target_row, max_col, month)
    _wrap_accumulator_self_references(ws, last_target_row)
    _write_dynamic_year_breakdown(ws, last_target_row)

    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
    jm_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
    _relocate_cross_sheet_references(
        ws, last_target_row, max_col,
        sheet_map={"Data": (data_offset, "Data"), "Johansen_Monthly": (jm_offset, "Johansen_Monthly")},
    )
    _write_roll_aware_price_changes(ws, month, last_target_row)
    _write_profit_take_overlay(ws, month, last_target_row)


def _write_profit_take_overlay(ws, month: str, last_target_row: int) -> None:
    """Profit-take and month-start re-entry on top of the stop overlay (see
    profit_take.py): rewrites BO/BP/BQ/BT/BU and adds helper columns EK:EO
    from the overlay's first steady-state row. Off (identical to the plain
    stop overlay) until this month's Config rows 30/31 are set."""
    for col, header in OVERLAY_HEADERS.items():
        ws[f"{col}1"] = header
    for row in range(OVERLAY_FIRST_ROW, last_target_row + 1):
        for col, formula in overlay_formulas(month, row).items():
            ws[f"{col}{row}"] = formula


def _write_roll_aware_price_changes(ws, month: str, last_target_row: int) -> None:
    """Replace the plain day-over-day price changes in Z:AF with roll-aware
    ones (see roll_fix.py), so a contract roll day is priced entirely in the
    old contract. Written after the cross-sheet relocation: these formulas
    reference only this sheet's own price columns and 'Roll Prices', whose
    per-month blocks are already addressed absolutely."""
    for row in range(3, last_target_row + 1):
        for col, formula in price_change_formulas(month, row, file_syntax=True).items():
            ws[f"{col}{row}"] = formula


def _write_generalized_trailing_window_accum_cols(ws, last_target_row: int, max_col: int) -> None:
    """targeted-fix-round-2, finding 1: BZ/CA/CB/CC ("24-26 Cumulative
    PnL"/"24-26 Running Equity Peak"/"24-26 Daily Drawdown"/"24-26 Max
    Drawdown") used to be a BOOT_END entry keyed to Dec's real row 741 --
    Dec's own live file is genuinely blank for these columns until row 740
    (wherever calendar year 2024 happens to start in Dec's specific
    historical data), then holds a literal 0 boundary at row 740 and the real
    formula from row 741 on. That row boundary is Dec-specific (a trailing
    "last ~2-3 calendar years" window with no general definition), not a
    property of the strategy shared by every book, so it cannot be
    reproduced as a bootstrap window in a general template.

    Inspection of the live workbook confirmed the formula SHAPE at row 741+
    is byte-identical (modulo column letters) to BV/BW/BX/BY's all-time
    Risk-Adjusted equivalent, both driven by BT ("Effective Daily PnL"):
        BZ{r} = BZ{r-1}+BT{r}          (cf. BV{r} = BV{r-1}+BT{r})
        CA{r} = MAX(BZ{r},CA{r-1})     (cf. BW{r} = MAX(BV{r},BW{r-1}))
        CB{r} = BZ{r}-CA{r}            (cf. BX{r} = BV{r}-BW{r})
        CC{r} = MIN(CB{r},CC{r-1})     (cf. BY{r} = MIN(BX{r},BY{r-1}))
    So these four are generalized to ordinary ROW3_OK self-referencing
    accumulators instead: row 2 holds a literal 0 boundary (the same
    convention Y2/AZ2/BE2/BJ2/CK2 already use), row 3 onward holds the real
    formula, anchored to row 3 instead of Dec's row 741. BZ/CA/CC get the
    IFERROR self-heal wrap via ACCUM_SELF_REF_COLS below, same as every
    other accumulator in this family; CB does not (it derives from BZ/CA in
    the same row, same reasoning as CM/BX).

    Deliberate consequence: BZ/CA/CB/CC now numerically equal BV/BW/BX/BY
    row-for-row (both are "since inception" now) -- this is intentional, not
    a bug, and is exactly why validate_against_dec.py's
    EXCLUDED_COMPARE_COLUMNS now also excludes these four: Dec's own real
    values for them (window starting row 740) no longer describe what this
    template's same-named columns compute by design.
    """
    if any(column_index_from_string(c) > max_col for c in GENERALIZED_TRAILING_WINDOW_ACCUM_COLS):
        return
    ws["BZ2"] = 0
    ws["CA2"] = 0
    ws["CB2"] = 0
    ws["CC2"] = 0
    for row in range(3, last_target_row + 1):
        prev = row - 1
        ws[f"BZ{row}"] = f"=BZ{prev}+BT{row}"
        ws[f"CA{row}"] = f"=MAX(BZ{row},CA{prev})"
        ws[f"CB{row}"] = f"=BZ{row}-CA{row}"
        ws[f"CC{row}"] = f"=MIN(CB{row},CC{prev})"


def _apply_targeted_fix_round_2_overrides(ws, last_target_row: int, max_col: int, month: str) -> None:
    """A handful of one-off cell-level fixes from the targeted-fix-round-2
    review, applied as explicit overrides after the generic baseline/
    BOOT_END/PARAM_BLOCK_COLS fill passes above have already written
    something (Dec's own literal text) into these exact cells.

    - DE43/DF43/DF46 (pre-existing: '=LOOKUP(2,1/($BV:$BV<>""),$BV:$BV)' and
      siblings over $BW:$BW/$BY:$BY) and DG49/DH49 (were hardcoded
      '=$CC$1386'/'=$CA$1386'): all five are "last real value in a per-day
      accumulator column" lookups. Two distinct problems, both traced while
      fixing finding 1/4 of targeted-fix-round-2:
        (a) DG49/DH49 hardcoded Dec's own real last data row (1386). A
            template of arbitrary num_rows (or a real month with fewer/more
            actual pasted rows than the sheet's row capacity) has no reason
            to have real data end at row 1386.
        (b) Using "1/(col<>"")" against the SAME self-referencing accumulator
            column being read (as Johansen_Monthly!K3's idiom does, safely,
            against its own manually-maintained, genuinely-blank-past-the-end
            column) is unsafe here: Position Model's per-day columns are
            filled with real FORMULAS all the way to the sheet's row
            capacity regardless of how many rows of real data were actually
            pasted, and a formula cell referencing blank upstream Data is
            NOT itself blank (Excel returns 0 for a formula reference to a
            genuinely empty cell, and 0<>"" is TRUE) -- so testing a
            self-referencing accumulator column's own blankness finds
            whatever row LOOKS non-blank last, which is contaminated by
            padding rows far past the real data, not Dec's true last row.
            Fixed by testing blankness against Data!$A (the actual pasted
            Date column) instead: a genuinely-unpasted cell is truly empty
            (no formula at all), and Excel treats a blank cell as equal to
            "" for comparison purposes, so Data!$A correctly goes "blank" the
            moment real data runs out, regardless of what any per-day FORMULA
            column downstream does. The return vector stays the target
            column (BV/BW/BY/CC/CA); only the TEST vector changes.
      DG49/DH49 additionally can never be expected to match Dec's real value
      any more (see validate_against_dec.py's EXCLUDED_COMPARE_COLUMNS/
      PARAM_BLOCK_STRUCTURALLY_REDEFINED_CELLS) -- CC/CA no longer describe
      Dec's real "24-26 window", they describe the all-time metric (same as
      CL/CN and BW/BY now, since finding 1 redefined them), so the genuine
      all-time max-drawdown/peak-equity this now correctly computes is a
      different, and equally valid, number than Dec's narrower-window one.
    - BN6 (='=BN4*(7.5/100)-BN5') and DQ10 (='=-65000*DQ11'): scattered
      literal numeric constants that rewrite_constants()/CONSTANT_MAP could
      never have caught, because that mechanism only replaces ADDRESS TOKENS
      (e.g. the literal text "$DK$3") that appear as a REFERENCE inside some
      OTHER cell's formula -- it does a blind string substitution, not
      numeric-literal detection. Here the raw numbers 7.5 and -65000 are
      typed directly into the formula text, not expressed as a reference to
      the parameter cell that already holds that value elsewhere
      (ZLMultiplier = 7.5, StopLossUSD_Base = -65000), so no address token
      ever appears for CONSTANT_MAP to find and replace. Rewritten directly
      here to reference the parameter cell instead.

      multi-month Task 11, review fix round 1: these two were originally
      rewritten to the bare workbook-scoped DEFINED NAMES "ZLMultiplier" and
      "StopLossUSD_Base". That was correct for the single-month template,
      whose write_config_sheet created those names -- but the multi-month
      write_config_sheet creates ZERO defined names (see
      test_config_sheet_no_defined_names), one Config COLUMN per contract
      month instead, so both cells would have evaluated to #NAME? on all
      eight Position Model tabs. They now use this month's own Config cell
      via config_cell(month, row), the same idiom as the rest of this
      migration. Note these are bare NAMES in the formula text, not
      "$XX$N"-style address tokens, so constant_map_for_month's substitution
      path could never have reached them -- hence the direct fix here.
      Rows confirmed against config_map.CONFIG_FIELDS: ZLMultiplier = row 8,
      StopLossUSD_Base = row 27.
    - DJ224:DN227 ("Peak"/"Trough"/"Range" MAX/MIN over the now-decommissioned
      CQ-CU per-year columns): see the DJ224_DN227 clearing block below and
      the new EG/EH/EI columns in _write_dynamic_year_breakdown.
    """
    blank_test = f'Data!$A$3:$A${last_target_row}<>""'
    if column_index_from_string("DF") <= max_col:
        ws["DE43"] = f'=LOOKUP(2,1/({blank_test}),$BV$3:$BV${last_target_row})'
        ws["DF43"] = f'=LOOKUP(2,1/({blank_test}),$BW$3:$BW${last_target_row})'
        ws["DF46"] = f'=LOOKUP(2,1/({blank_test}),$BY$3:$BY${last_target_row})'
    if column_index_from_string("DH") <= max_col:
        ws["DG49"] = f'=LOOKUP(2,1/({blank_test}),$CC$3:$CC${last_target_row})'
        ws["DH49"] = f'=LOOKUP(2,1/({blank_test}),$CA$3:$CA${last_target_row})'

    if column_index_from_string("BN") <= max_col:
        ws["BN6"] = f"=BN4*(Config!{config_cell(month, _CONFIG_ROW['ZLMultiplier'])}/100)-BN5"

    if column_index_from_string("DQ") <= max_col:
        ws["DQ10"] = f"=Config!{config_cell(month, _CONFIG_ROW['StopLossUSD_Base'])}*DQ11"

    # targeted-fix-round-2, finding 2: DJ224:DN227 used MAX/MIN(CQ2:CQ253)-
    # style formulas over the now-decommissioned, fixed-row-window per-year
    # columns (CQ-CU, see DECOMMISSIONED_PER_YEAR_COLS). With CQ-CU cleared
    # to blank, MAX/MIN of an all-blank range silently evaluates to 0 instead
    # of loudly failing -- exactly the kind of defect this project exists to
    # eliminate. There is no dynamic-length equivalent of "one column per
    # calendar year positioned at a specific column letter" (DJ=year1,
    # DK=year2, ... only ever fits Dec's own fixed 5-year layout), so rather
    # than reproduce that same defect in a new shape, these specific cells
    # are cleared entirely. The genuine general equivalent -- per-year Peak/
    # Trough/Range, keyed dynamically on Year via MAXIFS/MINIFS over CP
    # ("Calendar Year Cumulative PnL"), consistent with EE/EF's own per-year
    # SUMIFS design -- now lives in the dynamic year-breakdown table's new
    # EG/EH/EI columns (see _write_dynamic_year_breakdown).
    if column_index_from_string("DN") <= max_col:
        for col_letter in ("DJ", "DK", "DL", "DM", "DN"):
            for row in (224, 225, 227):
                ws[f"{col_letter}{row}"] = None


def _wrap_accumulator_self_references(ws, last_target_row: int) -> None:
    """CRITICAL 1 fix: for each column in ACCUM_SELF_REF_COLS, wrap the
    self-reference to the immediately preceding row (e.g. the "Y2" in
    "=Y2+X3") in IFERROR(...,0). Applied as a post-processing pass over
    whatever formula text is already sitting in each cell (written by the
    baseline fill_sheet_from_pattern pass for ROW3_OK columns, or by the
    BOOT_END steady-state shift for BOOTSTRAP columns), so it transparently
    covers both categories without needing to know which one a given column
    is in.
    """
    for col_letter in ACCUM_SELF_REF_COLS:
        col_idx = column_index_from_string(col_letter)
        if col_idx > ws.max_column:
            continue
        for row in range(3, last_target_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            value = cell.value
            if not (isinstance(value, str) and value.startswith("=")):
                continue
            prev_row = row - 1
            pattern = re.compile(rf"(?<![A-Za-z0-9$]){re.escape(col_letter)}{prev_row}(?![A-Za-z0-9])")
            if pattern.search(value):
                cell.value = pattern.sub(f"IFERROR({col_letter}{prev_row},0)", value, count=1)


def _write_dynamic_year_breakdown(ws, last_target_row: int = 3001) -> None:
    """Replaces Dec's hardcoded per-year row-bound COUNT/PnL columns (e.g.
    AH$2:AH$249 or CR3='=CP250', both specific to Dec's own historical row
    numbers) with a genuine, formula-driven multi-year table keyed on this
    sheet's own Year column (CO), so activating a new calendar year -- or a
    brand-new book with completely different dates -- needs no manual edit
    and no new column.

    MAJOR 4 fix: the previous design wrote exactly one data row (DZ3/
    EA3:ED3) -- it never actually enumerated multiple years or produced a
    total, despite the column header claiming a "Total" row existed. This
    now writes NUM_YEAR_BREAKDOWN_ROWS year-rows (each year = the previous
    row's year + 1, starting from this sheet's own first data row via
    YEAR(A3)) followed by a genuine Total row, then clears everything past
    it -- capped so the table never extends past last_target_row (relevant
    only for very small num_rows, e.g. in tests; the production 3,000-row
    template has ample headroom).

    EA/EB/EC/ED reproduce Dec's per-year trade-count columns (Pos1/Pos2/Pos3/
    Total), unchanged in formula shape from the original fix. EE/EF are new:
    they replace the decommissioned CQ-CU ("<year> Cumulative PnL") and CE-CJ
    ("RA <year> Cumulative PnL") columns with genuine per-year sums via
    SUMIFS against this sheet's own real per-day PnL columns (BK = "Grand
    Total PnL", BT = "Effective Daily PnL", the risk-adjusted equivalent).

    targeted-fix-round-2, finding 2: EG/EH/EI (Peak/Trough/Range) are new --
    the genuine per-year equivalent of Dec's decommissioned DJ224:DN227
    (MAX/MIN over the now-gone CQ-CU fixed-row-window columns), via MAXIFS/
    MINIFS over CP ("Calendar Year Cumulative PnL", which already resets to 0
    at the start of each calendar year using this sheet's own Year column,
    CO -- exactly what CQ-CU used to be a fixed-row window INTO). No "Total"
    row is written for these three: a sum of yearly peaks/troughs is not a
    meaningful "all-time" peak/trough, so the total row is deliberately left
    blank for them (EA-EF's SUM-based Total remains valid and unaffected).

    targeted-fix-round-2, finding 5: every SUMIFS/COUNTIFS/MAXIFS/MINIFS
    range below is bounded by this call's own last_target_row parameter
    (the sheet's actual last row) instead of a hardcoded "3000" -- the
    old hardcode silently excluded row 3001 whenever num_rows was 3000
    (last_target_row = num_rows+1 = 3001), and would have been wrong for
    any other num_rows too.
    """
    ws["DZ2"] = "Year"
    ws["EA2"] = "Pos1"
    ws["EB2"] = "Pos2"
    ws["EC2"] = "Pos3"
    ws["ED2"] = "Total"
    ws["EE2"] = "Cumulative PnL"
    ws["EF2"] = "RA Cumulative PnL"
    ws["EG2"] = "Peak"
    ws["EH2"] = "Trough"
    ws["EI2"] = "Range"

    num_year_rows = min(NUM_YEAR_BREAKDOWN_ROWS, max(1, last_target_row - 3))
    first_year_row = 3
    last_year_row = first_year_row + num_year_rows - 1
    total_row = last_year_row + 1
    last_row = last_target_row

    for row in range(first_year_row, last_year_row + 1):
        ws[f"DZ{row}"] = "=YEAR(A3)" if row == first_year_row else f"=DZ{row - 1}+1"
        ws[f"EA{row}"] = f'=COUNTIFS($CO$3:$CO${last_row},DZ{row},$AH$3:$AH${last_row},"<>")'
        ws[f"EB{row}"] = f'=COUNTIFS($CO$3:$CO${last_row},DZ{row},$AJ$3:$AJ${last_row},"<>")'
        ws[f"EC{row}"] = f'=COUNTIFS($CO$3:$CO${last_row},DZ{row},$AL$3:$AL${last_row},"<>")'
        ws[f"ED{row}"] = f"=SUM(EA{row}:EC{row})"
        ws[f"EE{row}"] = f"=SUMIFS($BK$3:$BK${last_row},$CO$3:$CO${last_row},DZ{row})"
        ws[f"EF{row}"] = f"=SUMIFS($BT$3:$BT${last_row},$CO$3:$CO${last_row},DZ{row})"
        ws[f"EG{row}"] = f"=MAXIFS($CP$3:$CP${last_row},$CO$3:$CO${last_row},DZ{row})"
        ws[f"EH{row}"] = f"=MINIFS($CP$3:$CP${last_row},$CO$3:$CO${last_row},DZ{row})"
        ws[f"EI{row}"] = f"=ABS(EG{row}-EH{row})"

    if total_row <= last_target_row:
        ws[f"DZ{total_row}"] = "Total"
        for col_letter in ("EA", "EB", "EC", "ED", "EE", "EF"):
            ws[f"{col_letter}{total_row}"] = f"=SUM({col_letter}{first_year_row}:{col_letter}{last_year_row})"

    for col_letter in YEAR_BREAKDOWN_COLS:
        col_idx = column_index_from_string(col_letter)
        for row in range(total_row + 1, last_target_row + 1):
            # NOTE: ws.cell(row, column, value=None) is a no-op in openpyxl
            # (it only assigns when value is not None) -- must set .value
            # directly to actually clear a previously-written cell.
            ws.cell(row=row, column=col_idx).value = None


def _relocate_cross_sheet_references(ws, last_target_row: int, max_col: int, sheet_map: dict) -> None:
    """Position Model tabs are never column-shifted themselves (each is a
    full standalone sheet starting at column A, exactly like today's
    single-month template) -- only their CROSS-sheet references need
    relocating, to whichever column-block this month's Data/Johansen_Monthly
    live at. Read-only inspection of the live source workbook confirmed
    Position Model references no other sheet directly (not DiffLag, not
    Matrices/Johansen/Results, not Monthly Beta -- nothing references Monthly
    Beta from anywhere in the workbook), so sheet_map only ever needs "Data"
    and "Johansen_Monthly" keys, and deliberately has no None key: Position
    Model's own internal self-references (the accumulator columns, the
    dynamic year breakdown, etc.) must never be touched by this pass.

    A blanket full-sheet scan, not a targeted one, on purpose -- consistent
    with this project's established preference for an opt-out validation
    surface over an opt-in one (see docs/superpowers/HANDOFF.md's "hard-won
    gotchas"): safer to check every cell than to guess which columns might
    contain a cross-sheet reference.
    """
    for row in ws.iter_rows(min_row=1, max_row=max(last_target_row, ws.max_row), max_col=max_col):
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                cell.value = shift_formula_columns(value, sheet_map)
