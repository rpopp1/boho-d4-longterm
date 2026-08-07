# Reconciled against a live, read-only inspection of the Dec source workbook
# (Johansen_FM_BOHO_D4_Strategy.xlsx, 'Position Model' and 'Matrices' sheets).
# See .superpowers/sdd/2026-08-05-johansen-fixed-month-template/task-4-report.md
# for the exact inspection output and the corrections made relative to the
# plan's original draft.

from openpyxl.utils import get_column_letter
from multi_month import MONTHS

# Per-month Config layout: every field is independently valued per contract
# month (one column per month), replacing the original single-month
# CONFIG_LAYOUT (a single shared column B + workbook-scoped defined names),
# which was removed in Task 11 of the multi-month consolidation once every
# builder had migrated to CONFIG_FIELDS/constant_map_for_month.
#
# (field name, Config sheet ROW, default value or formula)
CONFIG_FIELDS = [
    ("ContractLabel",             2,  "TEMPLATE"),
    ("ZLContractMonth",           3,  "TEMPLATE"),
    ("HOContractMonth",           4,  "TEMPLATE"),
    ("HOGalPerContract",          7,  42000),   # was Position Model!$DK$5 ("HO Gal")
    ("ZLMultiplier",              8,  7.5),     # was Position Model!$DK$3 ("ZL Multiplier")
    ("ZLCentsPerLB",              9,  60000),   # was Position Model!$DK$2 ("ZL cents/LBs"); rows 9/10 fill the blank spacer rows before row 11
    ("JohansenWindowLen",         11, 63),      # was Matrices!$B$3 ("Window Length (rows)")
    ("ZThreshold_Tier1",          14, 2.25),    # was Position Model!$DV$3
    ("ZThreshold_Tier2",          15, 2.1),     # was Position Model!$DV$4
    ("ZThreshold_Tier3",          16, 1.75),    # was Position Model!$DV$5
    ("ZExitThreshold",            17, 0.1),     # was Position Model!$DV$6
    ("SizingWeight_Tier1",        20, 0.25),    # was Position Model!$DL$7
    ("SizingWeight_Tier2",        21, 0.5),     # was Position Model!$DM$7
    ("SizingWeight_Tier3",        22, 0.25),    # was Position Model!$DN$7
    ("EWMA_Lambda",               25, 0.94),    # was Position Model!$DQ$7 ("EWMA Lambda")
    ("ShockMultiplierThreshold",  26, 1.25),    # was Position Model!$DQ$8 ("Shock Multiplier Threshold")
    ("StopLossUSD_Base",          27, -65000),  # was Position Model!$DQ$10 ("Trade Level Stop Loss", before *DQ11 scalar)
    ("D4Multiplier",              28, 1.5),     # was Position Model!$DK$11 ("D4 Multiplier")
    ("RiskLevelInput",            29, 1),       # was Position Model!$DQ$11 ("Risk Level")
]

CONFIG_MONTH_COLUMNS = {month: get_column_letter(2 + i) for i, month in enumerate(MONTHS)}

_CONFIG_ROW_BY_FIELD = {name: row for name, row, _default in CONFIG_FIELDS}

# Dec absolute Position-Model-local-param-block address -> Config field name.
# Keys here must be long/unusual absolute addresses (high columns like DK-DX)
# with effectively zero collision risk against unrelated formulas on other
# sheets, since this dict is applied via blind string substitution with no
# sheet-scoping (see formula_tools.rewrite_constants). Do NOT add short,
# generic addresses (e.g. "$B$3") here — see
# matrices_only_constant_map_for_month below.
# NOTE: addresses below are Position Model sheet cells unless noted otherwise.
_ADDRESS_TO_FIELD = {
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
