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
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Position Model"]
    for row in range(4, 10):
        for col in ("DZ", "EA", "EB", "EC", "ED"):
            v = ws[f"{col}{row}"].value
            assert v is None
    # And row 3 (row 2 header labels aside) must not contain any of the
    # Dec-specific literal per-year row bounds either.
    for col in ("EA", "EB", "EC"):
        v = ws[f"{col}3"].value
        assert "$AH$2:$AH$249" not in (v or "")
        assert "$AJ$2:$AJ$249" not in (v or "")
        assert "$AL$2:$AL$249" not in (v or "")


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


def test_non_self_referencing_drawdown_columns_are_untouched():
    # CM (=CK{row}-CL{row}) and BX (=BV{row}-BW{row}) are not self-referencing
    # accumulators -- they derive directly from other (now error-hardened)
    # columns, so they need no IFERROR wrap of their own.
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, num_rows=30)
    ws = wb["Position Model"]
    assert ws["CM3"].value == "=CK3-CL3"
    assert ws["BX21"].value == "=BV21-BW21"
