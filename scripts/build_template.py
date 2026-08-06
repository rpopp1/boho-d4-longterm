import openpyxl

from build_config_sheet import write_config_sheet
from build_static_sheets import write_notes_sheet, write_data_sheet
from build_diff_lag import write_diff_lag_sheet
from build_matrices_johansen import write_matrices_sheet, write_johansen_sheet, write_results_sheet
from build_johansen_monthly import write_johansen_monthly_sheet
from build_monthly_beta import write_monthly_beta_sheet
from build_position_model import write_position_model_sheet

SHEET_ORDER = [
    "Notes", "Config", "Data", "DiffLag", "Matrices",
    "Johansen", "Johansen_Monthly", "Monthly Beta",
    "Position Model", "Results",
]


def build_template(source_path: str, output_path: str, num_rows: int = 3000) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    write_notes_sheet(wb)
    write_config_sheet(wb)
    write_data_sheet(wb)
    write_diff_lag_sheet(wb, source_path, num_rows=num_rows)
    write_matrices_sheet(wb, source_path)
    write_johansen_sheet(wb, source_path)
    write_johansen_monthly_sheet(wb, source_path, num_rows=num_rows)
    write_monthly_beta_sheet(wb, source_path, num_rows=num_rows)
    write_position_model_sheet(wb, source_path, num_rows=num_rows)
    write_results_sheet(wb, source_path)
    wb.remove(default_sheet)
    wb._sheets.sort(key=lambda ws: SHEET_ORDER.index(ws.title))
    wb.save(output_path)


if __name__ == "__main__":
    import pathlib
    source = (pathlib.Path(__file__).parent / "DEC_SOURCE_PATH.txt").read_text().strip()
    out = pathlib.Path(__file__).parent.parent / "Johansen_Template_Engine.xlsx"
    build_template(source, str(out))
    print(f"Template written to {out}")
