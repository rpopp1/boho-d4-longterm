import re

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP
from formula_tools import extract_row_formulas, rewrite_constants, shift_formula_rows

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
    "BZ": 741, "CA": 741, "CB": 741, "CC": 741,
}

PARAM_BLOCK_COLS = {
    "CE", "CF", "CG", "CH", "CI", "CJ", "CQ", "CR", "CS", "CT", "CU",
    "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO",
    "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY",
    "EE",
}

# Columns holding the Dec-specific hardcoded per-year breakdown table that
# _write_dynamic_year_breakdown replaces. See the NOTE above.
YEAR_BREAKDOWN_COLS = ("DZ", "EA", "EB", "EC", "ED", "EE", "EF")

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
ACCUM_SELF_REF_COLS = ("Y", "AZ", "BE", "BJ", "CK", "CL", "CN", "BV", "BW", "BY")


def write_position_model_sheet(wb, source_path: str, num_rows: int = 3000, max_col: int = None) -> None:
    if "Position Model" in wb.sheetnames:
        del wb["Position Model"]
    ws = wb.create_sheet("Position Model")

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
                constant_map=CONSTANT_MAP,
            )

        # Row 2 filled from its own literal Dec pattern (boundary row, not
        # shifted from row 3).
        row2 = extract_row_formulas(source_path, "Position Model", row=2, max_col=max_col)
        for col_letter, value in row2.items():
            if isinstance(value, str) and value.startswith("="):
                value = rewrite_constants(value, CONSTANT_MAP)
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
                value = src.cell(row=row, column=col_idx).value
                if isinstance(value, str) and value.startswith("="):
                    value = rewrite_constants(value, CONSTANT_MAP)
                ws.cell(row=row, column=col_idx, value=value)

            if boot_end <= last_target_row:
                pattern_value = src.cell(row=boot_end, column=col_idx).value
                if isinstance(pattern_value, str) and pattern_value.startswith("="):
                    for row in range(boot_end, last_target_row + 1):
                        delta = row - boot_end
                        shifted = shift_formula_rows(pattern_value, delta)
                        rewritten = rewrite_constants(shifted, CONSTANT_MAP)
                        ws.cell(row=row, column=col_idx, value=rewritten)
                else:
                    ws.cell(row=boot_end, column=col_idx, value=pattern_value)

        # Fix up PARAM_BLOCK columns: static side-panel cells, copied literally
        # from the source wherever it has content, never shifted or extended.
        for col_letter in PARAM_BLOCK_COLS:
            col_idx = column_index_from_string(col_letter)
            if col_idx > max_col:
                continue
            copy_end = min(last_target_row, src.max_row)
            for row in range(3, copy_end + 1):
                value = src.cell(row=row, column=col_idx).value
                if value is None:
                    continue
                if isinstance(value, str) and value.startswith("="):
                    value = rewrite_constants(value, CONSTANT_MAP)
                ws.cell(row=row, column=col_idx, value=value)
    finally:
        source_wb.close()

    _wrap_accumulator_self_references(ws, last_target_row)
    _write_dynamic_year_breakdown(ws, last_target_row)


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
    """Replaces Dec's hardcoded per-year row-bound COUNT columns (e.g. AH$2:AH$249
    for 2021 specifically) with COUNTIFS keyed on this sheet's own Year column (CO),
    so a new calendar year needs no manual edit.

    Fix round 1: also explicitly clears rows 4..last_target_row of
    DZ/EA/EB/EC/ED. The baseline fill_sheet_from_pattern call (these columns
    are no longer in PARAM_BLOCK_COLS, see the module-level NOTE) writes a
    row-3-shifted formula into every one of those rows first; because the
    source row-3 formulas in EA/EB/EC are fully $-anchored (e.g.
    "=COUNT($AH$2:$AH$249)"), shifting is a no-op and the same Dec-specific
    literal would otherwise be duplicated down every row. Clearing them here
    guarantees no Dec-specific literal survives anywhere in these columns.
    """
    ws["EA2"] = "Pos1"
    ws["EB2"] = "Pos2"
    ws["EC2"] = "Pos3"
    ws["ED2"] = "Total"
    ws["DZ3"] = "=YEAR(A3)"
    ws["EA3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AH$3:$AH$3000,"<>")'
    ws["EB3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AJ$3:$AJ$3000,"<>")'
    ws["EC3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AL$3:$AL$3000,"<>")'
    ws["ED3"] = "=SUM(EA3:EC3)"

    for col_letter in YEAR_BREAKDOWN_COLS:
        col_idx = column_index_from_string(col_letter)
        for row in range(4, last_target_row + 1):
            # NOTE: ws.cell(row, column, value=None) is a no-op in openpyxl
            # (it only assigns when value is not None) -- must set .value
            # directly to actually clear a previously-written cell.
            ws.cell(row=row, column=col_idx).value = None
