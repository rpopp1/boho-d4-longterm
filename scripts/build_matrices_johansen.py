import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from formula_tools import rewrite_constants, unwrap_scalar_array_formula, shift_formula_columns
from config_map import constant_map_for_month, matrices_only_constant_map_for_month
from multi_month import (
    mjr_offsets, column_offset, DIFFLAG_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col,
    MJR_BLOCK_WIDTH, MJR_MONTH_GAP,
)

COMBINED_SHEET_NAME = "Matrices_Johansen_Results"


def _copy_block(src, ws, max_row, max_col, self_offset, sheet_map, constant_map):
    for row in src.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value is None:
                continue
            value = unwrap_scalar_array_formula(cell.value, cell.coordinate)
            if isinstance(value, str) and value.startswith("="):
                value = rewrite_constants(value, constant_map)
                value = shift_formula_columns(value, sheet_map)
            source_col, _source_row = coordinate_from_string(cell.coordinate)
            target_col = shift_col(source_col, self_offset)
            ws[f"{target_col}{cell.row}"] = value


def write_matrices_johansen_results_sheet(wb, source_path: str, month: str) -> None:
    if COMBINED_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(COMBINED_SHEET_NAME)
    else:
        ws = wb[COMBINED_SHEET_NAME]

    offsets = mjr_offsets(month)
    diff_lag_offset = column_offset(month, DIFFLAG_BLOCK_WIDTH, GAP)
    jm_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
    month_constant_map = constant_map_for_month(month)
    matrices_map = {**month_constant_map, **matrices_only_constant_map_for_month(month)}

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        _copy_block(
            source_wb["Matrices"], ws, max_row=53, max_col=8,
            self_offset=offsets["Matrices"],
            sheet_map={None: (offsets["Matrices"], None), "DiffLag": (diff_lag_offset, "DiffLag"),
                       "Johansen_Monthly": (jm_offset, "Johansen_Monthly")},
            constant_map=matrices_map,
        )
        _copy_block(
            source_wb["Johansen"], ws, max_row=26, max_col=6,
            self_offset=offsets["Johansen"],
            sheet_map={None: (offsets["Johansen"], None), "Matrices": (offsets["Matrices"], None)},
            constant_map=month_constant_map,
        )
        _copy_block(
            source_wb["Results"], ws, max_row=11, max_col=2,
            self_offset=offsets["Results"],
            sheet_map={None: (offsets["Results"], None), "Matrices": (offsets["Matrices"], None),
                       "Johansen": (offsets["Johansen"], None)},
            constant_map=month_constant_map,
        )
    finally:
        source_wb.close()

    # Structural overrides, same reasoning as today's single-month template:
    # B3 = window length (was the literal 63 in Dec's own live file);
    # B2 = the dynamic Active End Row pointer on this month's Johansen_Monthly block.
    b_col = shift_col("B", offsets["Matrices"])
    k_col_jm = shift_col("K", jm_offset)
    ws[f"{b_col}3"] = f"={matrices_map['$B$3']}"
    ws[f"{b_col}2"] = f"=Johansen_Monthly!${k_col_jm}$3"

    # Plain text label in the gap right after this month's mega-block.
    mega_offset = column_offset(month, MJR_BLOCK_WIDTH, MJR_MONTH_GAP)
    ws[f"{shift_col('A', mega_offset + MJR_BLOCK_WIDTH)}1"] = month
