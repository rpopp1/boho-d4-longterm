import openpyxl
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from build_position_model import write_position_model_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_position_model_fills_and_rewrites_constants():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Position Model"]
    m4 = ws["M4"].value
    assert "$DV$6" not in m4
    assert "ZExitThreshold" in m4
    i3 = ws["I3"].value
    assert "Johansen_Monthly!$G:$G" in i3


def test_year_breakdown_is_dynamic_not_hardcoded_per_year():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Position Model"]
    # The old Dec pattern hardcoded literal row bounds like AH$2:AH$249 per calendar year.
    # The template's equivalent must key off the sheet's own Year column instead.
    year_formula = ws["EA3"].value
    assert "COUNTIFS" in year_formula or "SUMIFS" in year_formula
    assert "AH$2:AH$249" not in year_formula


def test_no_array_formula_cells_in_rows_2_and_3():
    # Independent-verification sanity check (see task brief): earlier scan claimed
    # zero ArrayFormula/spill-anchor cells in Position Model rows 2-3. Confirm the
    # built sheet has none either (fill_sheet_from_pattern would have preserved any
    # as real ArrayFormula objects, not plain strings).
    from openpyxl.worksheet.formula import ArrayFormula

    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Position Model"]
    for row in (2, 3):
        for col in range(1, ws.max_column + 1):
            assert not isinstance(ws.cell(row=row, column=col).value, ArrayFormula)


def test_bootstrap_columns_use_literal_rows_before_steady_state():
    # BC3 in the live source is '=IF($S1268=0,0,-$S1268*$AF1269*$DF$5)' -- an
    # unanchored reference to a fixed, unusual row (1268/1269), not a genuine
    # per-day relative reference (same failure category as row-1-11's DO10/C1266).
    # BC's real per-day steady-state formula only begins at row 20 in the live
    # workbook. Row 3 must be copied literally (unshifted), not row-shifted.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=25)
    ws = wb["Position Model"]
    bc3 = ws["BC3"].value
    assert bc3 is not None
    assert "1268" in bc3 or "1269" in bc3
    # And it must NOT have been row-shifted (shifting would move 1268/1269 far
    # away from their true fixed anchors for a target row of 3, this is a no-op,
    # so the real assertion is that it is copied verbatim from the source row 3).
    assert "$S1268" in bc3 or "$T1268" in bc3


def test_year_breakdown_rows_beyond_3_have_no_leftover_dec_literals():
    # Fix round 1: an earlier version put DZ/EA/EB/EC/ED in PARAM_BLOCK_COLS,
    # which literally copied the Dec source's whole hardcoded per-year table
    # (DZ4:DZ9 = literal years 2022-2026 + "Total"; EA4:EA8 = literal
    # fixed-row-range COUNTs like '=COUNT($AH$250:$AH$494)') into the built
    # sheet -- _write_dynamic_year_breakdown only overwrote rows 2-3, so rows
    # 4-9 silently retained Dec-specific literals. Confirm none survive.
    #
    # MAJOR 4 superseded the single-row design: rows 4+ are no longer cleared
    # blank -- they now hold the genuine, formula-driven multi-year breakdown
    # (see test_dynamic_year_breakdown_enumerates_multiple_years_plus_total).
    # What must never reappear is Dec's own literal fixed-row-bound formulas.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Position Model"]
    for row in range(3, 12):
        for col in ("EA", "EB", "EC"):
            v = ws[f"{col}{row}"].value
            assert "$AH$2:$AH$249" not in (v or "")
            assert "$AJ$2:$AJ$249" not in (v or "")
            assert "$AL$2:$AL$249" not in (v or "")
            assert "$AH$250:$AH$494" not in (v or "")


def test_bootstrap_columns_reach_steady_state_row_shift():
    # From row 20 on, BC follows a genuine per-day relative pattern in the live
    # workbook: '=IF($T19=0,0,-$T19*$AF20*$DM$3)' at row 20, shifting by +1 row
    # at a time. Confirm the built sheet reproduces this shifted pattern.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=25)
    ws = wb["Position Model"]
    bc21 = ws["BC21"].value
    assert bc21 is not None
    assert "$T20" in bc21
    assert "$AF21" in bc21


def test_accumulator_self_references_are_iferror_wrapped():
    # CRITICAL 1: Y/AZ/BE/BJ/CK/CL/CN/BV/BW/BY are self-referencing accumulators
    # (e.g. Y{row} = Y{row-1}+X{row}). Dec's live file has a literal hardcoded 0
    # at specific warm-up rows (e.g. Y20) that breaks the #N/A chain once, but the
    # template writes the general formula everywhere, so any #N/A/#DIV/0! in a
    # per-row contribution propagates through the accumulator forever. Each
    # self-reference to the previous row must be wrapped in IFERROR(...,0) so the
    # running total self-heals once real numbers exist again.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=30)
    ws = wb["Position Model"]

    y3 = ws["Y3"].value
    assert y3 == "=IFERROR(Y2,0)+X3"

    az3 = ws["AZ3"].value
    assert az3 == "=IFERROR(AZ2,0)+AY3"

    be3 = ws["BE3"].value
    assert be3 == "=IFERROR(BE2,0)+BD3"

    bj3 = ws["BJ3"].value
    assert bj3 == "=IFERROR(BJ2,0)+BI3"

    ck3 = ws["CK3"].value
    assert ck3 == "=IFERROR(CK2,0)+BK3"

    cl3 = ws["CL3"].value
    assert cl3 == "=MAX(IFERROR(CL2,0),CK3)"

    cn3 = ws["CN3"].value
    assert cn3 == "=MIN(IFERROR(CN2,0),CM3)"

    # BV/BW/BY are BOOTSTRAP columns; their real per-day steady formula only
    # begins at row 21 in the live workbook.
    bv21 = ws["BV21"].value
    assert bv21 == "=IFERROR(BV20,0)+BT21"

    bw21 = ws["BW21"].value
    assert bw21 == "=MAX(BV21,IFERROR(BW20,0))"

    by21 = ws["BY21"].value
    assert by21 == "=MIN(BX21,IFERROR(BY20,0))"

    # A later row must also stay wrapped, and reference its own immediate
    # predecessor (not the row-3/row-21 boundary), so a mid-strategy #N/A
    # doesn't propagate past that one row either.
    y25 = ws["Y25"].value
    assert y25 == "=IFERROR(Y24,0)+X25"


def test_additional_self_referencing_columns_caught_by_widened_validation_are_wrapped():
    # CRITICAL 2's widened validation harness caught two more genuinely
    # self-referencing columns beyond the original ~12-column enumeration:
    # CP/CD ("Calendar Year Cumulative PnL"/RA equivalent) and the
    # BM/BO/BQ/BU trade-state-machine cluster (BL/BN/BP/BR/BS/BT are NOT
    # self-referencing themselves and resolve once these four are hardened).
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=30)
    ws = wb["Position Model"]

    assert ws["CP3"].value == '=IF(CO3=CO2,IFERROR(CP2,0)+BK3,0)'
    assert ws["CD21"].value == '=IF(CO21=CO20,IFERROR(CD20,0)+BT21,0)'

    assert ws["BM21"].value == '=IF(BN21=1,1,IF(IFERROR(BM20,0)=1,IF(BL21=0,0,1),0))'
    assert ws["BO21"].value == '=IF(BM21=0,0,IF(BN21=1,BL21,IFERROR(BO20,0)+BL21))'
    assert ws["BQ21"].value == '=IF(BN21=1,0,IF(IFERROR(BQ20,0)=1,1,BP21))'
    assert ws["BU21"].value == '=IF(BS21=0,0,IF(BN21=1,BT21,IFERROR(BU20,0)+BT21))'


def test_boot_end_literal_copy_clears_genuinely_blank_source_cells():
    # CRITICAL 3-class fix: the BOOT_END literal-copy loop (rows 3..boot_end-1)
    # used to write via ws.cell(row, column, value=value), which is a no-op
    # in openpyxl when value is None -- so wherever the source's bootstrap
    # window is genuinely blank (e.g. BC6:BC19 in the live workbook -- BC's
    # boot_end is 20, and only rows 3-5 have real fixed-row formulas, rows
    # 6-19 are blank), whatever the baseline fill_sheet_from_pattern pass had
    # already shifted into those cells (a nonsensical shifted version of
    # BC's fixed-row-1268/1269 pattern) survived underneath.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=25)
    ws = wb["Position Model"]
    for row in range(6, 20):
        assert ws[f"BC{row}"].value is None, f"BC{row} unexpectedly non-empty: {ws[f'BC{row}'].value!r}"


def test_param_block_cols_have_no_leftover_baseline_fill_beyond_source_extent():
    # CRITICAL 3: PARAM_BLOCK_COLS's literal-copy pass used to `continue` (skip)
    # wherever the source was None, silently leaving whatever the earlier
    # baseline fill_sheet_from_pattern shift-fill pass had written underneath.
    # DJ/DK/DL/DM/DN are only populated through row 227 in the live source;
    # nothing beyond that should survive in the built sheet at all.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=300)
    ws = wb["Position Model"]
    for col in ("DJ", "DK", "DL", "DM", "DN"):
        for row in (228, 250, 300, 301):
            assert ws[f"{col}{row}"].value is None, f"{col}{row} unexpectedly non-empty"


def test_param_block_cols_still_copy_their_real_populated_extent():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=300)
    ws = wb["Position Model"]
    # DK1 is a real literal constant (750000) in the live source's param block.
    assert ws["DK1"].value == 750000


def test_decommissioned_per_year_pnl_cols_are_fully_cleared():
    # MAJOR 4: CQ-CU/CE-CJ used to be literally copied from Dec (a column per
    # calendar year, with hardcoded row bounds like "=CP250" for 2022's start).
    # They're now handled generically by the dynamic year-breakdown table
    # instead, so nothing -- not even Dec's header labels -- should survive in
    # these columns anywhere in the built sheet.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=300)
    ws = wb["Position Model"]
    for col in ("CE", "CF", "CG", "CH", "CI", "CJ", "CQ", "CR", "CS", "CT", "CU"):
        for row in (1, 2, 3, 4, 250, 300):
            v = ws[f"{col}{row}"].value
            assert v is None, f"{col}{row} unexpectedly non-empty: {v!r}"


def test_dynamic_year_breakdown_enumerates_multiple_years_plus_total():
    # MAJOR 4: the old design wrote exactly one data row (DZ3/EA3:ED3). The
    # real fix must enumerate a genuine sequence of years, starting from the
    # sheet's own first data row, and end with a Total row -- all formula
    # driven, no Dec-specific hardcoded years or row bounds.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=300)
    ws = wb["Position Model"]
    assert ws["DZ3"].value == "=YEAR(A3)"
    assert ws["DZ4"].value == "=DZ3+1"
    assert ws["DZ5"].value == "=DZ4+1"

    total_row = None
    for row in range(3, 40):
        if ws[f"DZ{row}"].value == "Total":
            total_row = row
            break
    assert total_row is not None, "no Total row found in the dynamic year breakdown"
    assert total_row > 4, "breakdown must enumerate more than a single year before the Total row"

    ea_total = ws[f"EA{total_row}"].value
    assert ea_total == f"=SUM(EA3:EA{total_row - 1})"

    # Everything past the Total row must be cleared.
    assert ws[f"DZ{total_row + 1}"].value is None


def test_dynamic_year_breakdown_produces_genuine_per_year_pnl_sums():
    # MAJOR 4: the real per-year PnL totals (Dec's CQ-CU / CE-CJ) must now be
    # produced via SUMIFS keyed on this sheet's own Year column (CO), not by
    # literally copying Dec's per-year columns.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=300)
    ws = wb["Position Model"]
    ee3 = ws["EE3"].value
    assert ee3 == "=SUMIFS($BK$3:$BK$3000,$CO$3:$CO$3000,DZ3)"
    ef3 = ws["EF3"].value
    assert ef3 == "=SUMIFS($BT$3:$BT$3000,$CO$3:$CO$3000,DZ3)"


def test_non_self_referencing_drawdown_columns_are_untouched():
    # CM (=CK{row}-CL{row}) and BX (=BV{row}-BW{row}) are not self-referencing
    # accumulators -- they derive directly from other (now error-hardened)
    # columns, so they need no IFERROR wrap of their own.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=30)
    ws = wb["Position Model"]
    assert ws["CM3"].value == "=CK3-CL3"
    assert ws["BX21"].value == "=BV21-BW21"
