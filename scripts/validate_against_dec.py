import csv
import shutil
import openpyxl
import win32com.client
from config_map import CONFIG_LAYOUT

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

COMPARE_COLUMNS = ["I", "L", "S", "T", "U", "V", "X"]  # spread, z-score, positions 1-3, total, daily PnL


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


def _recalculate_with_excel(path: str) -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(path)
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()


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

    mismatches = []
    max_row = min(dec_pm.max_row, template_pm.max_row)
    for row in range(3, max_row + 1):
        for col in COMPARE_COLUMNS:
            dec_val = dec_pm[f"{col}{row}"].value
            tmpl_val = template_pm[f"{col}{row}"].value
            if isinstance(dec_val, (int, float)) and isinstance(tmpl_val, (int, float)):
                if abs(dec_val - tmpl_val) > 1e-6:
                    mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val})
            elif dec_val != tmpl_val:
                mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val})

    if mismatches:
        with open(output_diff_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["row", "col", "dec", "template"])
            writer.writeheader()
            writer.writerows(mismatches)

    return mismatches
