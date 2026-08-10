from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
from config_map import config_cell


def write_notes_sheet(wb) -> None:
    if "Notes" in wb.sheetnames:
        del wb["Notes"]
    ws = wb.create_sheet("Notes", 0)
    lines = [
        "Johansen (2-variable) Cointegration Multi-Month Template",
        "",
        "This workbook covers all 8 BOHO contract months (January, March, May, July,",
        "August, September, October, December) at once. 'Data', 'DiffLag',",
        "'Matrices_Johansen_Results', 'Johansen_Monthly', and 'Monthly Beta' each hold",
        "8 side-by-side per-month column-blocks in month order. 'Position Model' has",
        "8 separate tabs, one per contract month.",
        "",
        "How to activate a contract month:",
        "1) Fill in that month's column on the 'Config' sheet: contract identity,",
        "   contract specs, Johansen window length, z-score thresholds, sizing",
        "   weights, and risk parameters. The 'Config' sheet has one column per",
        "   contract month, independently calibrated -- there are no shared/global",
        "   values.",
        "2) Go to 'Data' and paste that month's SBO/HO settlement and open pricing",
        "   into that month's block, starting at its own row 2. The BOHO/D4 columns",
        "   in each block are pasted values, matching the live December workbook's",
        "   convention.",
        "3) DiffLag, Matrices_Johansen_Results, Johansen_Monthly, Monthly Beta, and",
        "   that month's Position Model tab recompute automatically from Data +",
        "   Config; no formulas need editing.",
        "4) REQUIRED MANUAL STEP, every month, per contract month: on",
        "   Johansen_Monthly, in that month's block, append one new row to its",
        "   EndRow/Date/BOHO_Weight/D4_Weight log columns -- this walk-forward log",
        "   is manually maintained, NOT auto-computed, separately for each of the",
        "   8 months. That month's Position Model tab's spread construction looks",
        "   up the prior month's locked weights from this log -- if it isn't",
        "   advanced for a given month, that month's book will not compute new",
        "   positions.",
        "",
        "This template's engine formulas are generated from, and must remain",
        "behaviorally identical to, the live December BOHO/D4 workbook's engine.",
    ]
    for i, line in enumerate(lines, start=1):
        ws[f"A{i}"] = line


def write_data_sheet(wb, month: str) -> None:
    if "Data" not in wb.sheetnames:
        wb.create_sheet("Data")
    ws = wb["Data"]
    offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)

    def col(letter):
        return shift_col(letter, offset)

    ws[f"{col('A')}1"] = "Date"
    ws[f"{col('B')}1"] = "BOHO"
    ws[f"{col('C')}1"] = "D4"
    ws[f"{col('E')}1"] = "Settlement Pricing"
    ws[f"{col('K')}1"] = "Open Pricing"
    ws[f"{col('E')}2"] = "Date"
    ws[f"{col('F')}2"] = "SBO"
    ws[f"{col('G')}2"] = "HO"
    ws[f"{col('H')}2"] = "BOHO"
    ws[f"{col('I')}2"] = "D4"
    ws[f"{col('K')}2"] = "Date"
    ws[f"{col('L')}2"] = "SBO"
    ws[f"{col('M')}2"] = "HO"
    ws[f"{col('N')}2"] = "BOHO"
    ws[f"{col('S')}2"] = "ZL cents/LBs"
    ws[f"{col('T')}2"] = f"=Config!{config_cell(month, 9)}"   # ZLCentsPerLB, row 9
    ws[f"{col('S')}3"] = "ZL Multiplier"
    ws[f"{col('T')}3"] = f"=Config!{config_cell(month, 8)}"   # ZLMultiplier, row 8
    ws[f"{col('S')}4"] = "ZL Gal"
    ws[f"{col('T')}4"] = f"={col('T')}2/{col('T')}3"
    ws[f"{col('S')}5"] = "HO Gal"
    ws[f"{col('T')}5"] = f"=Config!{config_cell(month, 7)}"   # HOGalPerContract, row 7

    # Plain text label in the gap column right after this month's block, so
    # it's visually obvious where one month's block ends and the next
    # begins on the shared tab.
    ws[f"{shift_col('A', offset + DATA_BLOCK_WIDTH)}1"] = month
