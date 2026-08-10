import openpyxl, pathlib
from build_matrices_johansen import write_matrices_johansen_results_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_january_matrices_sub_block_at_column_a():
    wb = openpyxl.Workbook()
    write_matrices_johansen_results_sheet(wb, SOURCE_PATH, "January")
    ws = wb["Matrices_Johansen_Results"]
    assert ws["A2"].value == "End Row (DiffLag)"
    assert ws["B3"].value == "=Config!B11"           # JohansenWindowLen, January column B
    assert ws["B2"].value == "=Johansen_Monthly!$K$3"  # January's own Johansen_Monthly offset is 0
    assert ws["C7"].value == (
        '=IF($B$5=0,"",SUMPRODUCT(INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2), '
        'INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2))/$B$5)'
    )


def test_january_johansen_sub_block_references_matrices_unqualified():
    wb = openpyxl.Workbook()
    write_matrices_johansen_results_sheet(wb, SOURCE_PATH, "January")
    ws = wb["Matrices_Johansen_Results"]
    # Johansen's sub-block offset for January is 10 (Matrices width 8 + sub-gap 2),
    # so its own B3 (source formula "=Matrices!C50") physically lands at
    # B(2)+10=12='L', row 3. The formula itself references Matrices' C50
    # (Matrices offset 0 for January, so the column is untouched) with the
    # "Matrices!" qualifier stripped, since Matrices now lives on this same sheet.
    assert ws["L3"].value == "=C50"


def test_january_results_sub_block_references_matrices_and_johansen_unqualified():
    wb = openpyxl.Workbook()
    write_matrices_johansen_results_sheet(wb, SOURCE_PATH, "January")
    ws = wb["Matrices_Johansen_Results"]
    # Results' sub-block offset for January is 18 (10 + Johansen width 6 + sub-gap 2),
    # so its own B2/B3 (source formulas "=Matrices!B2"/"=Johansen!B11")
    # physically land at B(2)+18=20='T', rows 2/3.
    # B2 = "=Matrices!B2" -> Matrices offset 0 -> stays "=B2"
    assert ws["T2"].value == "=B2"
    # B3 = "=Johansen!B11" -> Johansen offset 10 -> B(2)+10=12='L' -> "=L11"
    assert ws["T3"].value == "=L11"


def test_march_matrices_sub_block_relocated_and_references_marchs_diff_lag():
    wb = openpyxl.Workbook()
    write_matrices_johansen_results_sheet(wb, SOURCE_PATH, "March")
    ws = wb["Matrices_Johansen_Results"]
    # March's mega-block offset: 1 * (20 + 4) = 24. Matrices sub-block offset = 24
    # (self-references shift by this). DiffLag's offset for March is 10
    # (only the "DiffLag!" cross-references shift by this, separately).
    assert ws["Y2"].value == "End Row (DiffLag)"   # A(1)+24=25='Y'
    assert ws["AA7"].value == (                    # C(3)+24=27='AA'
        '=IF($Z$5=0,"",SUMPRODUCT(INDEX(DiffLag!O:O,$Z$4):INDEX(DiffLag!O:O,$Z$2), '  # B(2)+24=26='Z'; E(5)+10=15='O'
        'INDEX(DiffLag!O:O,$Z$4):INDEX(DiffLag!O:O,$Z$2))/$Z$5)'
    )


def test_january_month_label_in_gap_after_mega_block():
    from multi_month import column_offset, MJR_BLOCK_WIDTH, MJR_MONTH_GAP, shift_col
    wb = openpyxl.Workbook()
    write_matrices_johansen_results_sheet(wb, SOURCE_PATH, "January")
    label_col = shift_col("A", column_offset("January", MJR_BLOCK_WIDTH, MJR_MONTH_GAP) + MJR_BLOCK_WIDTH)
    assert wb["Matrices_Johansen_Results"][f"{label_col}1"].value == "January"
