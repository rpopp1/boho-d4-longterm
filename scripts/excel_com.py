"""Isolated Excel automation shared by the roll-fix scripts.

Same safety pattern as build_dashboard.refresh_via_excel_com: DispatchEx forces
a brand-new Excel process (never the user's own session), and the run aborts if
it somehow lands in a pre-existing one. `manual_calc=True` opens a blank
workbook first and switches to manual calculation before the target opens, so
nothing in the target recalculates unless asked -- required for the ICE-fed
pricing workbook, whose _xll.ICESeries formulas cannot evaluate without the ICE
add-in, which automation sessions don't load.
"""

import contextlib
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path

XL_CALC_MANUAL = -4135
XL_CELL_TYPE_FORMULAS = -4123
XL_ERRORS = 16


def retry(fn, attempts=8, delay=1.0):
    """Retry a COM call through transient 'Call was rejected by callee' and
    OneDrive 'disconnected from its clients' errors."""
    last_exc = None
    for _ in range(attempts):
        try:
            return fn()
        except Exception as exc:
            last_exc = exc
            time.sleep(delay)
    raise last_exc


def _running_excel_pids():
    out = subprocess.run(["tasklist"], capture_output=True, text=True).stdout
    return {int(line.split()[1]) for line in out.splitlines() if line.upper().startswith("EXCEL.EXE")}


@contextlib.contextmanager
def isolated_workbook(path: Path, manual_calc: bool = False):
    """Yield (excel, workbook) for `path` inside a fresh, invisible Excel
    process. The caller saves; the workbook is always closed and the process
    always quit on exit."""
    import win32com.client as win32
    import win32process

    pids_before = _running_excel_pids()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    my_pid = retry(lambda: win32process.GetWindowThreadProcessId(excel.Hwnd)[1])
    if my_pid in pids_before:
        excel.Quit()
        raise RuntimeError(f"Excel automation landed in pre-existing process {my_pid}; aborted.")

    blank = None
    wb = None
    try:
        if manual_calc:
            blank = retry(lambda: excel.Workbooks.Add())
            excel.Calculation = XL_CALC_MANUAL
            excel.CalculateBeforeSave = False

        def _open():
            candidate = excel.Workbooks.Open(str(path), UpdateLinks=0)
            try:
                candidate.Worksheets(1).Name  # detect a disconnected handle now
                # OneDrive-synced files open with AutoSave on, even under
                # automation: a failed run's half-applied edits would be saved.
                candidate.AutoSaveOn = False
            except Exception:
                try:
                    candidate.Close(SaveChanges=False)
                except Exception:
                    pass
                raise
            return candidate

        wb = retry(_open, attempts=6, delay=5.0)
        yield excel, wb
    finally:
        for book in (wb, blank):
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception:
                    pass
        excel.Quit()
        missing = pids_before - _running_excel_pids()
        if missing:
            print(f"WARNING: pre-existing Excel PID(s) no longer running: {sorted(missing)}")


def sheet_names(wb) -> list:
    """Worksheet names by index (COM sheet collections don't support Python iteration reliably)."""
    return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]


def replace_sheet(wb, name: str):
    """Delete sheet `name` if present and add a fresh empty one at the end."""
    if name in sheet_names(wb):
        wb.Worksheets(name).Delete()
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    return ws


def error_cell_count(ws) -> int:
    try:
        return ws.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS, XL_ERRORS).Count
    except Exception:  # SpecialCells raises when nothing matches
        return 0


def reset_calc_mode_to_automatic(path: Path) -> None:
    """Strip calcMode="manual" from a saved workbook's calcPr element, so a
    file saved from a manual-calculation automation session opens in the
    user's normal automatic mode. Rewrites only xl/workbook.xml."""
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)  # Windows can't move or delete a file whose handle is still open
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    text = data.decode("utf-8")
                    text = re.sub(r'\s+calcMode="manual"', "", text)
                    text = re.sub(r'\s+calcOnSave="0"', "", text)
                    data = text.encode("utf-8")
                dst.writestr(item, data)
        shutil.move(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
