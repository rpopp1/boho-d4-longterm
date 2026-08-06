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
    ("ZLCentsPerLB",              "B9",  60000),   # was Position Model!$DK$2 ("ZL cents/LBs")
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

# Dec absolute address -> defined name, used by the formula rewriter.
# NOTE: addresses below are Position Model sheet cells unless noted otherwise.
CONSTANT_MAP = {
    "$DK$2": "ZLCentsPerLB",              # "ZL cents/LBs" label; NOT HO gallons (corrected, see report)
    "$DK$5": "HOGalPerContract",          # "HO Gal" label
    "$DK$3": "ZLMultiplier",
    "$B$3": "JohansenWindowLen",          # Matrices sheet address (corrected from $C$3, see report)
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
