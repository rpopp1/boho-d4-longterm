def write_notes_sheet(wb) -> None:
    if "Notes" in wb.sheetnames:
        del wb["Notes"]
    ws = wb.create_sheet("Notes", 0)
    lines = [
        "Johansen (2-variable) Cointegration Template Engine",
        "",
        "How to activate a new contract month:",
        "1) Fill in the 'Config' sheet: contract identity, contract specs, Johansen window",
        "   length, z-score thresholds, sizing weights, and risk parameters.",
        "2) Go to the 'Data' sheet and paste that month's SBO/HO settlement and open",
        "   pricing starting at row 2. Columns B (BOHO) and C (D4) are pasted values,",
        "   matching the live December workbook's convention.",
        "3) DiffLag, Matrices, Johansen, Monthly Beta, and Position Model recompute",
        "   automatically from Data + Config; no formulas need editing.",
        "4) REQUIRED MANUAL STEP, every month: on Johansen_Monthly, append one new row",
        "   to columns F:I (Date, EndRow, BOHO_Weight, D4_Weight) — this walk-forward log",
        "   is manually maintained, NOT auto-computed. Log that month's end-of-month date",
        "   in F, the corresponding DiffLag row number in G, and that month's locked",
        "   BOHO_Weight/D4_Weight (read off Johansen_Monthly!C2/D2 at month-end) in H/I.",
        "   Position Model's spread construction (column I) looks up the prior month's",
        "   weights from this log — if it isn't advanced, it has nothing to find and the",
        "   book will not compute new positions for that month.",
        "5) Results summarizes the Johansen cointegration test for the current window.",
        "",
        "This template's engine formulas are generated from, and must remain",
        "behaviorally identical to, the live December BOHO/D4 workbook's engine.",
    ]
    for i, line in enumerate(lines, start=1):
        ws[f"A{i}"] = line


def write_data_sheet(wb) -> None:
    if "Data" in wb.sheetnames:
        del wb["Data"]
    ws = wb.create_sheet("Data")
    ws["A1"] = "Date"
    ws["B1"] = "BOHO"
    ws["C1"] = "D4"
    ws["E1"] = "Settlement Pricing"
    ws["K1"] = "Open Pricing"
    ws["E2"] = "Date"
    ws["F2"] = "SBO"
    ws["G2"] = "HO"
    ws["H2"] = "BOHO"
    ws["I2"] = "D4"
    ws["K2"] = "Date"
    ws["L2"] = "SBO"
    ws["M2"] = "HO"
    ws["N2"] = "BOHO"
    ws["S2"] = "ZL cents/LBs"
    ws["T2"] = "=ZLCentsPerLB"
    ws["S3"] = "ZL Multiplier"
    ws["T3"] = "=ZLMultiplier"
    ws["S4"] = "ZL Gal"
    ws["T4"] = "=T2/T3"
    ws["S5"] = "HO Gal"
    ws["T5"] = "=HOGalPerContract"
