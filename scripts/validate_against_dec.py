import csv
import shutil
import openpyxl
import win32com.client
from openpyxl.utils import get_column_letter
from config_map import CONFIG_LAYOUT
from build_position_model import PARAM_BLOCK_COLS, DECOMMISSIONED_PER_YEAR_COLS, YEAR_BREAKDOWN_COLS

# Dec's live Config-equivalent values, read from the same cells Task 4's
# CONSTANT_MAP was built from, used to populate the throwaway template copy
# so it trades on identical parameters to Dec.
DEC_LIVE_VALUES = {
    "HOGalPerContract": 42000,
    "ZLMultiplier": 7.5,
    "JohansenWindowLen": 63,
    "ZThreshold_Tier1": 2.25,
    "ZThreshold_Tier2": 2.1,
    "ZThreshold_Tier3": 1.75,
    "ZExitThreshold": 0.1,
    "SizingWeight_Tier1": 0.25,
    "SizingWeight_Tier2": 0.5,
    "SizingWeight_Tier3": 0.25,
    "EWMA_Lambda": 0.94,
    "ShockMultiplierThreshold": 1.25,
    "StopLossUSD_Base": -65000,
    "D4Multiplier": 1.5,
    "RiskLevelInput": 1,
    "ZLCentsPerLB": 60000,
}
# Field names above must match scripts/config_map.py's CONFIG_LAYOUT exactly
# (reconciled against the live workbook in Task 4) — "ShockMultiplierThreshold"
# and "D4Multiplier" are the corrected names; earlier drafts of this plan used
# "ShockMultiplierThresh" and "RiskLevelScalar", which no longer exist.

# CRITICAL 2 fix: the old COMPARE_COLUMNS = ["I","L","S","T","U","V","X"] was a
# narrow opt-in list covering only 7 of ~137 Position Model columns, and none
# of them were self-referencing accumulators -- so it could not have caught
# CRITICAL 1 (the accumulator #N/A-propagation defect), no matter how many
# times it was run. This is now an opt-OUT list: every column is compared
# except a small, documented set that is genuinely not comparable row-for-row:
#
#   - PARAM_BLOCK_COLS: static one-off calibration/summary cells (a side
#     panel, not per-day data) -- literally copied from Dec, so they'd trivially
#     match, but comparing them row-for-row provides no real validation signal.
#   - DECOMMISSIONED_PER_YEAR_COLS (CE-CJ / CQ-CU): Dec has real hardcoded
#     per-year data here; the template deliberately clears these columns
#     entirely (MAJOR 4) since the same numbers now live in the dynamic
#     year-breakdown table instead. A byte-for-byte compare would always
#     "mismatch" by design -- that is the intended behavior, not a defect.
#   - YEAR_BREAKDOWN_COLS (DZ-EF): the dynamic year-breakdown table. It is a
#     small per-YEAR summary table, not a per-DAY time series, and its row
#     layout does not correspond 1:1 with Dec's row-3-only original design, so
#     comparing it via this per-(day-)row harness is structurally meaningless.
#
# Any column not in this exclusion set -- including every accumulator/
# drawdown column CRITICAL 1 fixed -- is compared at every row.
EXCLUDED_COMPARE_COLUMNS = set(PARAM_BLOCK_COLS) | set(DECOMMISSIONED_PER_YEAR_COLS) | set(YEAR_BREAKDOWN_COLS)

# Excel error-value sentinels, as returned by openpyxl when a workbook is
# loaded with data_only=True and the cell's cached value is an error.
EXCEL_ERROR_VALUES = ("#N/A", "#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!")


def _is_excel_error(value) -> bool:
    return isinstance(value, str) and value.startswith(EXCEL_ERROR_VALUES)


def compare_columns_for(max_col: int) -> list:
    """Every column letter from A through max_col, excluding
    EXCLUDED_COMPARE_COLUMNS. Computed from the actual built sheet's column
    count rather than hardcoded, so it stays correct if Position Model's
    width ever changes."""
    return [
        get_column_letter(c)
        for c in range(1, max_col + 1)
        if get_column_letter(c) not in EXCLUDED_COMPARE_COLUMNS
    ]


def _populate_throwaway_copy(template_path: str, dec_source_path: str, throwaway_path: str) -> None:
    shutil.copy(template_path, throwaway_path)
    wb = openpyxl.load_workbook(throwaway_path, data_only=False)
    config_ws = wb["Config"]
    layout_by_name = {name: cell for name, cell, _default in CONFIG_LAYOUT}
    for name, value in DEC_LIVE_VALUES.items():
        cell = layout_by_name[name]
        config_ws[cell] = value
    dec_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    dec_data = dec_wb["Data"]
    data_ws = wb["Data"]
    for row in dec_data.iter_rows(min_row=2, max_row=dec_data.max_row, max_col=20):
        for cell in row:
            if cell.value is not None:
                data_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    # Seed the Johansen_Monthly walk-forward log (F:I) with Dec's real
    # historical values. As of the Task 11 v2 redesign, F/G/H/I are all
    # manually-maintained columns in the generated template (no formula
    # derives them) — Position Model's LOOKUP($A-1, Johansen_Monthly!$G:$G,
    # ...!$H:$H / ...!$I:$I) formulas have nothing to find without this,
    # regardless of any other fix, since this dependency predates and is
    # independent of the spill-anchor defect that was fixed above it.
    dec_jm = dec_wb["Johansen_Monthly"]
    jm_ws = wb["Johansen_Monthly"]
    for row in dec_jm.iter_rows(min_row=2, max_row=dec_jm.max_row, min_col=6, max_col=9):  # F:I
        for cell in row:
            if cell.value is not None:
                jm_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    wb.save(throwaway_path)


def _recalculate_with_excel(path: str, max_retries: int = 3) -> None:
    # Hardened after a real failure: excel.CalculateFullRebuild() can raise
    # "Call was rejected by callee" if fired while Excel is still busy with
    # its own implicit on-open recalculation — observed with the heavier
    # formula load from the Johansen_Monthly redesign. A short settle delay
    # plus a retry-with-fresh-Excel-instance loop resolves this; it is COM
    # automation flakiness, not a formula defect.
    import time

    last_error = None
    for attempt in range(max_retries):
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(path)
            time.sleep(2)  # let Excel's own on-open recalculation settle
            excel.CalculateFullRebuild()
            wb.Save()
            wb.Close()
            excel.Quit()
            return
        except Exception as exc:
            last_error = exc
            try:
                excel.Quit()
            except Exception:
                pass
            if attempt < max_retries - 1:
                time.sleep(3)
    raise last_error


def run_validation(template_path: str, dec_source_path: str, output_diff_path: str) -> list:
    # Global Constraints forbid writing into the live source file, so Dec is never
    # opened in-place by Excel here — it is only ever read by openpyxl (read-only),
    # and a separate throwaway copy of Dec itself is what gets recalculated.
    template_throwaway_path = template_path.replace(".xlsx", "_validation_copy.xlsx")
    dec_throwaway_path = dec_source_path.replace(".xlsx", "_dec_validation_copy.xlsx")
    shutil.copy(dec_source_path, dec_throwaway_path)

    _populate_throwaway_copy(template_path, dec_source_path, template_throwaway_path)
    _recalculate_with_excel(template_throwaway_path)
    _recalculate_with_excel(dec_throwaway_path)  # recalculate the COPY, never the original

    template_wb = openpyxl.load_workbook(template_throwaway_path, data_only=True)
    dec_wb = openpyxl.load_workbook(dec_throwaway_path, data_only=True)
    template_pm = template_wb["Position Model"]
    dec_pm = dec_wb["Position Model"]

    max_col = min(dec_pm.max_column, template_pm.max_column)
    compare_columns = compare_columns_for(max_col)

    mismatches = []
    max_row = min(dec_pm.max_row, template_pm.max_row)
    for row in range(3, max_row + 1):
        for col in compare_columns:
            dec_val = dec_pm[f"{col}{row}"].value
            tmpl_val = template_pm[f"{col}{row}"].value

            # CRITICAL 2's second, explicit check: the template must never
            # show an error where Dec's real recalculated value is a plain
            # number. This is exactly the shape of defect CRITICAL 1 was --
            # a narrow value-equality check alone can be right for the wrong
            # reason (e.g. both sides coincidentally blank), so this is
            # tagged and reported distinctly from a generic value mismatch.
            if isinstance(dec_val, (int, float)) and _is_excel_error(tmpl_val):
                mismatches.append({
                    "row": row, "col": col, "dec": dec_val, "template": tmpl_val,
                    "kind": "error_parity",
                })
                continue

            if isinstance(dec_val, (int, float)) and isinstance(tmpl_val, (int, float)):
                if abs(dec_val - tmpl_val) > 1e-6:
                    mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val, "kind": "value"})
            elif dec_val != tmpl_val:
                mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val, "kind": "value"})

    if mismatches:
        with open(output_diff_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["row", "col", "dec", "template", "kind"])
            writer.writeheader()
            writer.writerows(mismatches)

    return mismatches
