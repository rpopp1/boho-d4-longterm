# Reconciled against a live, read-only inspection of the Dec source workbook
# (Johansen_FM_BOHO_D4_Strategy.xlsx, 'Position Model' and 'Matrices' sheets).
# See .superpowers/sdd/2026-08-05-johansen-fixed-month-template/task-4-report.md
# for the exact inspection output and the corrections made relative to the
# plan's original draft.

# Defined name -> (Config sheet cell, default value or formula, source Dec address it replaces)
CONFIG_LAYOUT = [
    ("ContractLabel",             "B2",  "TEMPLATE"),
    ("ZLContractMonth",           "B3",  "TEMPLATE"),
    ("HOContractMonth",           "B4",  "TEMPLATE"),
    ("HOGalPerContract",          "B7",  42000),   # was Position Model!$DK$5 ("HO Gal")
    ("ZLMultiplier",              "B8",  7.5),     # was Position Model!$DK$3 ("ZL Multiplier")
    ("ZLCentsPerLB",              "B9",  60000),   # was Position Model!$DK$2 ("ZL cents/LBs"); B9/B10 fill the blank spacer rows before B11
    ("JohansenWindowLen",         "B11", 63),      # was Matrices!$B$3 ("Window Length (rows)")
    ("ZThreshold_Tier1",          "B14", 2.25),    # was Position Model!$DV$3
    ("ZThreshold_Tier2",          "B15", 2.1),     # was Position Model!$DV$4
    ("ZThreshold_Tier3",          "B16", 1.75),    # was Position Model!$DV$5
    ("ZExitThreshold",            "B17", 0.1),     # was Position Model!$DV$6
    ("SizingWeight_Tier1",        "B20", 0.25),    # was Position Model!$DL$7
    ("SizingWeight_Tier2",        "B21", 0.5),     # was Position Model!$DM$7
    ("SizingWeight_Tier3",        "B22", 0.25),    # was Position Model!$DN$7
    ("EWMA_Lambda",               "B25", 0.94),    # was Position Model!$DQ$7 ("EWMA Lambda")
    ("ShockMultiplierThreshold",  "B26", 1.25),    # was Position Model!$DQ$8 ("Shock Multiplier Threshold")
    ("StopLossUSD_Base",          "B27", -65000),  # was Position Model!$DQ$10 ("Trade Level Stop Loss", before *DQ11 scalar)
    ("D4Multiplier",              "B28", 1.5),     # was Position Model!$DK$11 ("D4 Multiplier")
    ("RiskLevelInput",            "B29", 1),       # was Position Model!$DQ$11 ("Risk Level")
]

# Dec absolute address -> defined name, used by the generic per-day/per-sheet
# formula rewriter (Tasks 5, 7, 9, 11, 12 — DiffLag, Monthly Beta,
# Johansen_Monthly, Position Model). Keys here must be long/unusual absolute
# addresses (high columns like DK-DX) with effectively zero collision risk
# against unrelated formulas on other sheets, since this dict is applied via
# blind string/regex substitution with no sheet-scoping. Do NOT add short,
# generic addresses (e.g. "$B$3") here — see MATRICES_ONLY_CONSTANT_MAP below.
# NOTE: addresses below are Position Model sheet cells unless noted otherwise.
CONSTANT_MAP = {
    "$DK$2": "ZLCentsPerLB",              # "ZL cents/LBs" label; NOT HO gallons (corrected, see report)
    "$DK$5": "HOGalPerContract",          # "HO Gal" label
    "$DK$3": "ZLMultiplier",
    "$DV$3": "ZThreshold_Tier1",
    "$DV$4": "ZThreshold_Tier2",
    "$DV$5": "ZThreshold_Tier3",
    "$DV$6": "ZExitThreshold",
    "$DL$7": "SizingWeight_Tier1",
    "$DM$7": "SizingWeight_Tier2",
    "$DN$7": "SizingWeight_Tier3",
    "$DQ$7": "EWMA_Lambda",
    "$DQ$8": "ShockMultiplierThreshold",
    "$DQ$10": "StopLossUSD_Base",
    "$DK$11": "D4Multiplier",
    "$DQ$11": "RiskLevelInput",
}

# Applied only within the Matrices sheet's own builder (Task 10) — never
# passed to the generic per-day cloner, since "$B$3" is too generic an
# address (column B, row 3) to safely blind-substitute across other sheets'
# formulas without risking a silent, uncaught collision.
MATRICES_ONLY_CONSTANT_MAP = {
    "$B$3": "JohansenWindowLen",          # Matrices!$B$3 ("Window Length (rows)"); corrected from $C$3, see report
}

from openpyxl.utils import get_column_letter
from multi_month import MONTHS

# Per-month Config layout: same fields/rows/defaults as CONFIG_LAYOUT above,
# but with every field independently valued per contract month (one column
# per month) instead of a single shared column + workbook-scoped defined
# name. CONFIG_LAYOUT/CONSTANT_MAP/MATRICES_ONLY_CONSTANT_MAP above are
# unused by the multi-month template once Tasks 5-11 finish migrating their
# consumers; they are removed as the last step of Task 11.
CONFIG_FIELDS = [(name, int("".join(ch for ch in cell if ch.isdigit())), default)
                  for name, cell, default in CONFIG_LAYOUT]

CONFIG_MONTH_COLUMNS = {month: get_column_letter(2 + i) for i, month in enumerate(MONTHS)}

_CONFIG_ROW_BY_FIELD = {name: row for name, row, _default in CONFIG_FIELDS}

# Dec absolute Position-Model-local-param-block address -> defined name,
# derived from CONSTANT_MAP above (same address->field mapping, just used
# internally by constant_map_for_month). Kept as a private lookup so
# constant_map_for_month below only needs to know "which Config row", not
# re-derive the address list.
_ADDRESS_TO_FIELD = CONSTANT_MAP


def config_cell(month: str, row: int) -> str:
    return f"{CONFIG_MONTH_COLUMNS[month]}{row}"


def constant_map_for_month(month: str) -> dict:
    """Dec absolute address -> that month's Config sheet cell reference
    (e.g. "Config!F17"), replacing CONSTANT_MAP's address -> defined-name
    mapping now that every Config field is independently valued per month
    instead of a single shared workbook-scoped name."""
    return {
        address: f"Config!{config_cell(month, _CONFIG_ROW_BY_FIELD[field])}"
        for address, field in _ADDRESS_TO_FIELD.items()
    }


def matrices_only_constant_map_for_month(month: str) -> dict:
    return {"$B$3": f"Config!{config_cell(month, _CONFIG_ROW_BY_FIELD['JohansenWindowLen'])}"}
