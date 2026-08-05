# Johansen Fixed-Month BOHO/D4 Template Engine Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce `Johansen_Template_Engine.xlsx` — a Config-driven, script-generated clone of the live December BOHO/D4 Johansen engine, parameterized so a new fixed contract month can be activated by filling in Config + pasting Data, with zero hand-edited formulas.

**Architecture:** A Python build script (`scripts/build_template.py`) reads the actual formulas out of the live source workbook `Johansen_FM_BOHO_D4_Strategy.xlsx` (a byte-identical copy of the live Dec engine — see Task 0), generalizes them with a formula row-shift utility and a constant→defined-name rewrite utility, and writes a fresh 9-sheet template workbook filled to 3,000 rows. A separate validation script proves numeric equivalence against Dec's real output using Excel COM automation (openpyxl cannot calculate formulas).

**Tech Stack:** Python 3, `openpyxl` (already installed, v3.1.5), `pytest` (to install), `pywin32` (to install, for Excel COM automation in the validation task).

## Global Constraints

- Never modify, copy from-in-place, or write into any file under the live `Johansen_BOHO_D4_Strategy_v08.*` or `Johansen_FM_BOHO_D4_Strategy.xlsx` paths — read-only access to these for formula extraction and validation comparison only.
- Output template file: pure `.xlsx`, no VBA/macros.
- Row capacity: 3,000 rows on every per-day sheet (Data, DiffLag, Johansen_Monthly log, Monthly Beta, Position Model).
- No contract-roll logic, no "Active Trading Window", no legacy/prototype sheet.
- Every scattered literal constant from Dec's Position Model / Matrices sheets must be replaced by a workbook-scoped Excel defined name bound to a Config sheet cell.
- The Johansen_Monthly "Active End Row" pointer must be a self-updating last-row formula, not a hardcoded cell reference.
- The Position Model per-calendar-year PnL breakdown must derive years dynamically from the Date column, not use hardcoded per-year row ranges.
- This directory is not a git repository — skip all `git commit` steps; each task's "commit" step is replaced with "save and verify the file/script is in place."

---

### Task 0: Establish the read-only source snapshot

**Files:**
- Create: `scripts/DEC_SOURCE_PATH.txt` (single line, the absolute path to the source workbook)

**Interfaces:**
- Produces: a single canonical path string every later task reads instead of hardcoding the path inline.

- [ ] **Step 1: Confirm the source file and record its path**

Run:
```bash
py -c "import openpyxl; wb=openpyxl.load_workbook(r'C:\Users\Ryan Popp\OneDrive - Karbone Inc\Desktop\Renewable Fuels Analysis\RINs\BOHO Trading\Johansen_FM_BOHO_D4_Strategy.xlsx'); print(wb.sheetnames)"
```
Expected output: `['Notes', 'Data', 'DiffLag', 'Matrices', 'Johansen', 'Johansen_Monthly', 'Monthly Beta', 'Results', 'Spread', 'Position Model']`

- [ ] **Step 2: Write the path file**

Write `scripts/DEC_SOURCE_PATH.txt` containing exactly:
```
C:\Users\Ryan Popp\OneDrive - Karbone Inc\Desktop\Renewable Fuels Analysis\RINs\BOHO Trading\Johansen_FM_BOHO_D4_Strategy.xlsx
```

- [ ] **Step 3: Verify**

Run: `py -c "print(open(r'scripts\DEC_SOURCE_PATH.txt').read().strip())"` from the `BOHO Trading` directory — output must match the path above exactly, with no trailing newline issues.

- [ ] **Step 4: Save checkpoint** (no git in this repo — confirm the file exists via `ls scripts/`)

---

### Task 1: Environment setup

**Files:**
- Modify: none (environment only)

**Interfaces:**
- Produces: `pytest` and `pywin32` available on the `py` interpreter for all later tasks.

- [ ] **Step 1: Install pytest**

Run: `py -m pip install pytest`
Expected: `Successfully installed pytest-...`

- [ ] **Step 2: Install pywin32**

Run: `py -m pip install pywin32`
Expected: `Successfully installed pywin32-...`

- [ ] **Step 3: Verify both import cleanly**

Run: `py -c "import pytest, win32com.client; print('ok')"`
Expected: `ok`

- [ ] **Step 4: Save checkpoint**

---

### Task 2: Formula extraction utility

**Files:**
- Create: `scripts/formula_tools.py`
- Test: `scripts/tests/test_formula_tools.py`

**Interfaces:**
- Produces: `extract_row_formulas(path: str, sheet_name: str, row: int, max_col: int) -> dict[str, str]` — returns `{column_letter: formula_or_value_string}` for every non-empty cell in that row. Array formulas (`openpyxl.worksheet.formula.ArrayFormula`) are unwrapped to their `.text` string.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_formula_tools.py
import pathlib
from formula_tools import extract_row_formulas

SOURCE = pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt"
SOURCE_PATH = SOURCE.read_text().strip()

def test_extract_position_model_row2():
    formulas = extract_row_formulas(SOURCE_PATH, "Position Model", row=2, max_col=12)
    assert formulas["A"] == "=Data!A2"
    assert formulas["I"] == (
        "=(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D2)"
        "+(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$I:$I)*$E2)"
    )
    assert formulas["J"] == "=AVERAGE($I$2:I2)"
    assert formulas["K"] == "=_xlfn.STDEV.S($I$2:I2)"
    assert formulas["L"] == "=(I2-J2)/K2"

def test_extract_diff_lag_row2():
    formulas = extract_row_formulas(SOURCE_PATH, "DiffLag", row=2, max_col=8)
    assert formulas["A"] == 1
    assert formulas["B"] == '=IF(Data!B2="","",Data!A2)'
    assert formulas["C"] == "=Data!B2"
    assert formulas["D"] == "=Data!C2"
    assert formulas["E"] == '=IF(OR(C2="",C1=""),"",C2-C1)'

def test_extract_array_formula_unwrapped():
    formulas = extract_row_formulas(SOURCE_PATH, "Johansen_Monthly", row=2, max_col=1)
    assert formulas["A"] == "=_xlfn._xlws.FILTER(DiffLag!$B$2:$B$1048576,DiffLag!$B$2:$B$1048576<>\"\")"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'formula_tools'`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/formula_tools.py
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter


def extract_row_formulas(path: str, sheet_name: str, row: int, max_col: int) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb[sheet_name]
    result = {}
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        value = cell.value
        if value is None:
            continue
        if isinstance(value, ArrayFormula):
            value = value.text
        result[get_column_letter(col)] = value
    return result
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v`
Expected: 3 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 3: Formula row-shift utility

This is the core "fill down" engine: given a formula extracted from one row, produce the equivalent formula for a different row, exactly matching what Excel does when you drag-fill a cell down — relative row references shift, `$`-anchored (absolute) rows and defined names do not.

**Files:**
- Modify: `scripts/formula_tools.py`
- Test: `scripts/tests/test_formula_tools.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `shift_formula_rows(formula: str, row_delta: int) -> str`.

- [ ] **Step 1: Write the failing tests**

```python
# append to scripts/tests/test_formula_tools.py
from formula_tools import shift_formula_rows

def test_shift_simple_relative_ref():
    assert shift_formula_rows("=Data!A2", 1) == "=Data!A3"
    assert shift_formula_rows("=Data!A2", 5) == "=Data!A7"

def test_shift_leaves_absolute_row_unchanged():
    assert shift_formula_rows("=(I2-$J$2)/K2", 3) == "=(I5-$J$2)/K5"

def test_shift_leaves_absolute_col_but_shifts_row():
    assert shift_formula_rows("=$A2+$B2", 2) == "=$A4+$B4"

def test_shift_expanding_range_pattern():
    # AVERAGE($I$2:I2) filled down 5 rows -> AVERAGE($I$2:I7)
    assert shift_formula_rows("=AVERAGE($I$2:I2)", 5) == "=AVERAGE($I$2:I7)"

def test_shift_cross_sheet_reference():
    assert shift_formula_rows(
        "=IF($L3<=$DV$6,0,IF(AND($M2=0,$L3>=M$1),$DX$3,M2))", 4
    ) == "=IF($L7<=$DV$6,0,IF(AND($M6=0,$L7>=M$1),$DX$3,M6))"

def test_shift_ignores_defined_name_tokens():
    # A bare name with no row digits must never be touched
    assert shift_formula_rows("=ZThreshold_Tier1*2", 10) == "=ZThreshold_Tier1*2"

def test_shift_lookup_with_full_column_refs_unchanged():
    formula = "=(LOOKUP($A2-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D2)"
    shifted = shift_formula_rows(formula, 1)
    assert shifted == "=(LOOKUP($A3-1, Johansen_Monthly!$G:$G, Johansen_Monthly!$H:$H)*D3)"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v -k shift`
Expected: FAIL with `ImportError: cannot import name 'shift_formula_rows'`

- [ ] **Step 3: Write minimal implementation**

```python
# append to scripts/formula_tools.py
import re

# Matches an optional sheet-qualifier, optional $ + column letters, optional $ + row digits.
# Full-column refs like G:G (no digits) simply won't match the row-digit group and pass through.
_CELL_REF_RE = re.compile(
    r"(?P<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<colabs>\$)?(?P<col>[A-Za-z]{1,3})"
    r"(?P<rowabs>\$)?(?P<row>\d+)"
)


def shift_formula_rows(formula: str, row_delta: int) -> str:
    def _replace(match: "re.Match") -> str:
        if match.group("rowabs"):
            return match.group(0)  # absolute row: never shift
        sheet = match.group("sheet") or ""
        colabs = match.group("colabs") or ""
        col = match.group("col")
        new_row = int(match.group("row")) + row_delta
        return f"{sheet}{colabs}{col}{new_row}"

    return _CELL_REF_RE.sub(_replace, formula)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v -k shift`
Expected: 7 passed

- [ ] **Step 5: Run the full test file to confirm no regressions**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v`
Expected: all passed (10 tests total from Tasks 2+3)

- [ ] **Step 6: Save checkpoint**

---

### Task 4: Inspect and finalize the Config constant map

Before writing the rewrite utility, pin down the exact source cells and labels — several of Dec's Position Model parameter cells were captured earlier without fully confirming their adjacent label text. This step reads the live workbook directly rather than guessing.

**Files:**
- Create: `scripts/config_map.py`

**Interfaces:**
- Produces: `CONSTANT_MAP: dict[str, str]` mapping a Dec absolute address (e.g. `"$DV$3"`) to a defined-name string (e.g. `"ZThreshold_Tier1"`), and `CONFIG_LAYOUT: list[tuple[str, str, object]]` of `(defined_name, config_cell, default_value_or_formula)` used by Task 5.

- [ ] **Step 1: Inspect the live parameter block and its labels**

Run:
```bash
py -c "
import openpyxl
wb = openpyxl.load_workbook(r'Johansen_FM_BOHO_D4_Strategy.xlsx', data_only=False)
ws = wb['Position Model']
for r in range(1, 12):
    for c in range(108, 136):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            print(ws.cell(row=r, column=c).coordinate, repr(v))
"
```
Record the exact label text found in the columns immediately left of DK/DL/DQ/DV/DW/DX (columns roughly DI-DJ, DP-DT) for every populated row 1-11. Use these confirmed labels to name each Config field in Step 2 — do not rely on labels guessed from memory.

- [ ] **Step 2: Write the constant map using the confirmed labels**

```python
# scripts/config_map.py

# Defined name -> (Config sheet cell, default value or formula, source Dec address it replaces)
CONFIG_LAYOUT = [
    ("ContractLabel",        "B2",  "TEMPLATE"),
    ("ZLContractMonth",      "B3",  "TEMPLATE"),
    ("HOContractMonth",      "B4",  "TEMPLATE"),
    ("HOGalPerContract",     "B7",  42000),          # was Position Model!$DK$2 / $DK$5
    ("ZLMultiplier",         "B8",  7.5),             # was Position Model!$DK$3
    ("JohansenWindowLen",    "B11", 63),              # was Matrices!$C$3
    ("ZThreshold_Tier1",     "B14", 2.25),            # was Position Model!$DV$3
    ("ZThreshold_Tier2",     "B15", 2.1),             # was Position Model!$DV$4
    ("ZThreshold_Tier3",     "B16", 1.75),            # was Position Model!$DV$5
    ("ZExitThreshold",       "B17", 0.1),             # was Position Model!$DV$6
    ("SizingWeight_Tier1",   "B20", 0.25),            # was Position Model!$DL$7
    ("SizingWeight_Tier2",   "B21", 0.5),             # was Position Model!$DM$7
    ("SizingWeight_Tier3",   "B22", 0.25),            # was Position Model!$DN$7
    ("EWMA_Lambda",          "B25", 0.94),            # was Position Model!$DQ$7
    ("ShockMultiplierThresh","B26", 1.25),            # was Position Model!$DQ$8
    ("StopLossUSD_Base",     "B27", -65000),          # was Position Model!$DQ$10 (before *DQ11 scalar)
    ("RiskLevelScalar",      "B28", 1.5),             # was Position Model!$DK$11
    ("RiskLevelInput",       "B29", 1),               # was Position Model!$DQ$11
]

# Dec absolute address -> defined name, used by the formula rewriter.
CONSTANT_MAP = {
    "$DK$2": "HOGalPerContract",
    "$DK$5": "HOGalPerContract",
    "$DK$3": "ZLMultiplier",
    "$DV$3": "ZThreshold_Tier1",
    "$DV$4": "ZThreshold_Tier2",
    "$DV$5": "ZThreshold_Tier3",
    "$DV$6": "ZExitThreshold",
    "$DL$7": "SizingWeight_Tier1",
    "$DM$7": "SizingWeight_Tier2",
    "$DN$7": "SizingWeight_Tier3",
    "$DQ$7": "EWMA_Lambda",
    "$DQ$8": "ShockMultiplierThresh",
    "$DQ$10": "StopLossUSD_Base",
    "$DK$11": "RiskLevelScalar",
    "$DQ$11": "RiskLevelInput",
}
```

- [ ] **Step 3: Reconcile against Step 1's actual output**

Compare every address in `CONSTANT_MAP` against the real labels captured in Step 1. If a label contradicts the name chosen above (e.g. `$DK$2` turns out to label something other than HO gallons/contract), rename the defined name and the `CONFIG_LAYOUT` entry to match reality before proceeding — this map must reflect the live file, not the assumption made when this plan was written.

- [ ] **Step 4: Verify by import**

Run: `cd scripts && py -c "from config_map import CONSTANT_MAP, CONFIG_LAYOUT; print(len(CONSTANT_MAP), len(CONFIG_LAYOUT))"`
Expected: two positive integers printed, no error.

- [ ] **Step 5: Save checkpoint**

---

### Task 5: Constant-to-name formula rewriter

**Files:**
- Modify: `scripts/formula_tools.py`
- Test: `scripts/tests/test_formula_tools.py`

**Interfaces:**
- Consumes: `CONSTANT_MAP` from `config_map.py` (Task 4).
- Produces: `rewrite_constants(formula: str, constant_map: dict[str, str]) -> str`.

- [ ] **Step 1: Write the failing tests**

```python
# append to scripts/tests/test_formula_tools.py
from formula_tools import rewrite_constants
from config_map import CONSTANT_MAP

def test_rewrite_single_constant():
    formula = "=IF($L3<=$DV$6,0,IF(AND($M2=0,$L3>=M$1),$DX$3,M2))"
    result = rewrite_constants(formula, CONSTANT_MAP)
    assert "$DV$6" not in result
    assert "ZExitThreshold" in result
    assert "$DX$3" in result  # not in the map, left untouched

def test_rewrite_leaves_unmapped_addresses_untouched():
    formula = "=Data!A2"
    assert rewrite_constants(formula, CONSTANT_MAP) == "=Data!A2"

def test_rewrite_multiple_constants_in_one_formula():
    formula = "=S2*($Z2*$DK$2-$AA2*$DK$5)"
    result = rewrite_constants(formula, CONSTANT_MAP)
    assert result == "=S2*($Z2*HOGalPerContract-$AA2*HOGalPerContract)"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v -k rewrite`
Expected: FAIL with `ImportError: cannot import name 'rewrite_constants'`

- [ ] **Step 3: Write minimal implementation**

```python
# append to scripts/formula_tools.py

def rewrite_constants(formula: str, constant_map: dict) -> str:
    # Longest addresses first so "$DK$2" doesn't get partially matched inside a longer token.
    for address in sorted(constant_map, key=len, reverse=True):
        formula = formula.replace(address, constant_map[address])
    return formula
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v -k rewrite`
Expected: 3 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 6: Config sheet writer with defined names

**Files:**
- Create: `scripts/build_config_sheet.py`
- Test: `scripts/tests/test_build_config_sheet.py`

**Interfaces:**
- Consumes: `CONFIG_LAYOUT` from `config_map.py`.
- Produces: `write_config_sheet(wb: openpyxl.Workbook) -> None`, which creates a `Config` sheet and registers one workbook-scoped `DefinedName` per `CONFIG_LAYOUT` entry.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_build_config_sheet.py
import openpyxl
from build_config_sheet import write_config_sheet

def test_config_sheet_created_with_defined_names():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    assert "Config" in wb.sheetnames
    ws = wb["Config"]
    assert ws["B14"].value == 2.25  # ZThreshold_Tier1 default
    assert "ZThreshold_Tier1" in wb.defined_names
    dn = wb.defined_names["ZThreshold_Tier1"]
    assert dn.attr_text == "Config!$B$14"

def test_all_layout_entries_produce_a_defined_name():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    from config_map import CONFIG_LAYOUT
    for name, cell, _default in CONFIG_LAYOUT:
        assert name in wb.defined_names, f"missing defined name {name}"
        assert wb.defined_names[name].attr_text == f"Config!${cell[0]}${cell[1:]}"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_config_sheet.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'build_config_sheet'`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_config_sheet.py
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from config_map import CONFIG_LAYOUT


def write_config_sheet(wb: "openpyxl.Workbook") -> None:
    if "Config" in wb.sheetnames:
        del wb["Config"]
    ws = wb.create_sheet("Config")
    ws["A1"] = "Field"
    ws["B1"] = "Value"
    for name, cell, default in CONFIG_LAYOUT:
        row = int("".join(ch for ch in cell if ch.isdigit()))
        label_cell = f"A{row}"
        ws[label_cell] = name
        ws[cell] = default
        col_letters = "".join(ch for ch in cell if ch.isalpha())
        ref = f"Config!${col_letters}${row}"
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_config_sheet.py -v`
Expected: 2 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 7: Generic row-fill sheet cloner

Combines Tasks 2, 3, and 5 into the one utility every per-day sheet builder will call: take a source pattern row from Dec, shift + rewrite it, and fill a target sheet down N rows.

**Files:**
- Create: `scripts/sheet_cloner.py`
- Test: `scripts/tests/test_sheet_cloner.py`

**Interfaces:**
- Consumes: `extract_row_formulas`, `shift_formula_rows`, `rewrite_constants` from `formula_tools.py`; `CONSTANT_MAP` from `config_map.py`.
- Produces: `fill_sheet_from_pattern(target_ws, source_path, source_sheet, pattern_row, target_start_row, num_rows, max_col, constant_map)`.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_sheet_cloner.py
import openpyxl
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP
import pathlib

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_fill_diff_lag_pattern_five_rows():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=5,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B3"].value == '=IF(Data!B3="","",Data!A3)'
    assert ws["B6"].value == '=IF(Data!B6="","",Data!A6)'
    assert ws["C4"].value == "=Data!B4"

def test_fill_position_model_rewrites_constants():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="Position Model",
        pattern_row=3,
        target_start_row=3,
        num_rows=3,
        max_col=13,
        constant_map=CONSTANT_MAP,
    )
    m_formula = ws["M4"].value
    assert "$DV$6" not in m_formula
    assert "ZExitThreshold" in m_formula
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_sheet_cloner.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'sheet_cloner'`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/sheet_cloner.py
from openpyxl.utils import get_column_letter, column_index_from_string
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants


def fill_sheet_from_pattern(
    target_ws,
    source_path: str,
    source_sheet: str,
    pattern_row: int,
    target_start_row: int,
    num_rows: int,
    max_col: int,
    constant_map: dict,
) -> None:
    pattern = extract_row_formulas(source_path, source_sheet, pattern_row, max_col)
    for i in range(num_rows):
        target_row = target_start_row + i
        row_delta = target_row - pattern_row
        for col_letter, value in pattern.items():
            if isinstance(value, str) and value.startswith("="):
                shifted = shift_formula_rows(value, row_delta)
                rewritten = rewrite_constants(shifted, constant_map)
                target_ws[f"{col_letter}{target_row}"] = rewritten
            else:
                target_ws[f"{col_letter}{target_row}"] = value
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_sheet_cloner.py -v`
Expected: 2 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 8: Notes and Data sheets

**Files:**
- Create: `scripts/build_static_sheets.py`
- Test: `scripts/tests/test_build_static_sheets.py`

**Interfaces:**
- Produces: `write_notes_sheet(wb)`, `write_data_sheet(wb)`.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_build_static_sheets.py
import openpyxl
from build_static_sheets import write_notes_sheet, write_data_sheet

def test_notes_sheet_describes_config_workflow():
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    text = " ".join(str(c.value) for c in wb["Notes"]["A"] if c.value)
    assert "Config" in text
    assert "Data" in text

def test_data_sheet_headers_and_no_price_values():
    wb = openpyxl.Workbook()
    write_data_sheet(wb)
    ws = wb["Data"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "BOHO"
    assert ws["C1"].value == "D4"
    assert ws["S2"].value == "ZL cents/LBs"
    assert ws["T2"].value == "=ZLCentsPerLB"
    assert ws["B2"].value is None  # no live price data in the template
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_static_sheets.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_static_sheets.py

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
        "3) DiffLag, Matrices, Johansen, Johansen_Monthly, Monthly Beta, and Position",
        "   Model recompute automatically from Data + Config; no formulas need editing.",
        "4) Results summarizes the Johansen cointegration test for the current window.",
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
```

Add `"ZLCentsPerLB"` to `CONFIG_LAYOUT`/`CONSTANT_MAP` in `config_map.py` (Task 4's map) bound to default `60000`, mapping from Dec's `Data!$T$2`, since this constant lives on Data in Dec, not Position Model — update Task 4's files accordingly before running this test.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_static_sheets.py -v`
Expected: 2 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 9: DiffLag and Monthly Beta sheets

**Files:**
- Create: `scripts/build_diff_lag.py`, `scripts/build_monthly_beta.py`
- Test: `scripts/tests/test_build_diff_lag.py`, `scripts/tests/test_build_monthly_beta.py`

**Interfaces:**
- Consumes: `fill_sheet_from_pattern` (Task 7).
- Produces: `write_diff_lag_sheet(wb, source_path, num_rows=3000)`, `write_monthly_beta_sheet(wb, source_path, num_rows=3000)`.

- [ ] **Step 1: Write the failing tests**

```python
# scripts/tests/test_build_diff_lag.py
import openpyxl, pathlib
from build_diff_lag import write_diff_lag_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_diff_lag_header_and_fill():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["DiffLag"]
    assert ws["A1"].value == "Row"
    assert ws["B1"].value == "Date"
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B11"].value == '=IF(Data!B11="","",Data!A11)'
```

```python
# scripts/tests/test_build_monthly_beta.py
import openpyxl, pathlib
from build_monthly_beta import write_monthly_beta_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_monthly_beta_header_and_fill():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Monthly Beta"]
    assert ws["A1"].value == "Date"
    assert ws["H1"].value == "Beta_Locked_Monthly"
    assert ws["A2"].value == "=Data!A2"
    assert ws["F2"].value == "=EOMONTH(A2,0)"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_diff_lag.py tests/test_build_monthly_beta.py -v`
Expected: both FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write minimal implementations**

```python
# scripts/build_diff_lag.py
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Row", "Date", "BOHO", "D4", "dBOHO", "dD4", "BOHO(-1)", "D4(-1)"]


def write_diff_lag_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "DiffLag" in wb.sheetnames:
        del wb["DiffLag"]
    ws = wb.create_sheet("DiffLag")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
```

```python
# scripts/build_monthly_beta.py
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Date", "BOHO", "D4", "BOHO_\u0394", "D4_\u0394", "MonthEnd", "Beta_120d", "Beta_Locked_Monthly"]


def write_monthly_beta_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Monthly Beta" in wb.sheetnames:
        del wb["Monthly Beta"]
    ws = wb.create_sheet("Monthly Beta")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Monthly Beta",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_diff_lag.py tests/test_build_monthly_beta.py -v`
Expected: both pass

- [ ] **Step 5: Save checkpoint**

---

### Task 10: Matrices and Johansen sheets

These are small fixed-size blocks (not per-day fills). Matrices' window length becomes a Config reference; its "End Row" pointer is wired to Johansen_Monthly's new dynamic pointer, built in Task 11 — agree on the target cell address (`Johansen_Monthly!$K$3`) now so both tasks are consistent.

**Files:**
- Create: `scripts/build_matrices_johansen.py`
- Test: `scripts/tests/test_build_matrices_johansen.py`

**Interfaces:**
- Consumes: `extract_row_formulas`, `rewrite_constants` from `formula_tools.py`.
- Produces: `write_matrices_sheet(wb, source_path)`, `write_johansen_sheet(wb, source_path)`.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_build_matrices_johansen.py
import openpyxl, pathlib
from build_matrices_johansen import write_matrices_sheet, write_johansen_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_matrices_window_length_uses_config_name():
    wb = openpyxl.Workbook()
    write_matrices_sheet(wb, SOURCE_PATH)
    ws = wb["Matrices"]
    assert ws["C3"].value == "=JohansenWindowLen"
    assert ws["B2"].value == "=Johansen_Monthly!$K$3"

def test_johansen_sheet_replicated():
    wb = openpyxl.Workbook()
    write_johansen_sheet(wb, SOURCE_PATH)
    ws = wb["Johansen"]
    assert ws["B3"].value == "=Matrices!C50"
    assert ws["B8"].value == '=IF(OR(B3="",B6=""),"",B3+B6)'
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_matrices_johansen.py
import openpyxl
from formula_tools import extract_row_formulas, rewrite_constants
from config_map import CONSTANT_MAP


def write_matrices_sheet(wb, source_path: str) -> None:
    if "Matrices" in wb.sheetnames:
        del wb["Matrices"]
    ws = wb.create_sheet("Matrices")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Matrices"]
    for row in src.iter_rows(min_row=1, max_row=53, max_col=8):
        for cell in row:
            if cell.value is None:
                continue
            ws[cell.coordinate] = cell.value
    # Structural template-specific overrides:
    ws["C3"] = "=JohansenWindowLen"          # was the literal 63
    ws["B2"] = "=Johansen_Monthly!$K$3"       # unchanged reference target, now resolved dynamically there
    for row in ws.iter_rows(min_row=1, max_row=53, max_col=8):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = rewrite_constants(cell.value, CONSTANT_MAP)


def write_johansen_sheet(wb, source_path: str) -> None:
    if "Johansen" in wb.sheetnames:
        del wb["Johansen"]
    ws = wb.create_sheet("Johansen")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Johansen"]
    for row in src.iter_rows(min_row=1, max_row=26, max_col=6):
        for cell in row:
            if cell.value is not None:
                ws[cell.coordinate] = cell.value
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v`
Expected: 2 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 11: Johansen_Monthly sheet with the dynamic Active-End-Row fix

**Files:**
- Create: `scripts/build_johansen_monthly.py`
- Test: `scripts/tests/test_build_johansen_monthly.py`

**Interfaces:**
- Produces: `write_johansen_monthly_sheet(wb, source_path, num_rows=3000)`.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_build_johansen_monthly.py
import openpyxl, pathlib
from build_johansen_monthly import write_johansen_monthly_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_active_end_row_is_dynamic_not_hardcoded():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    formula = ws["K3"].value
    assert formula != "=F69"
    assert "LOOKUP(2,1/(" in formula
    assert "F:F" in formula  # resolves against the log table's own F column

def test_headers_and_array_formula_present():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, num_rows=10)
    ws = wb["Johansen_Monthly"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "EndRow"
    a2 = ws["A2"].value
    assert "FILTER" in a2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_johansen_monthly.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_johansen_monthly.py
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" in wb.sheetnames:
        del wb["Johansen_Monthly"]
    ws = wb.create_sheet("Johansen_Monthly")
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Johansen_Monthly",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=4,
        constant_map=CONSTANT_MAP,
    )
    ws["K1"] = "Active End Row"
    # Self-updating last-row lookup, replacing Dec's hardcoded "=F69":
    # finds the last row in column F (the month-end log) that is non-blank.
    ws["K3"] = '=LOOKUP(2,1/(F:F<>""),F:F)'
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_johansen_monthly.py -v`
Expected: 2 passed

- [ ] **Step 5: Save checkpoint**

---

### Task 12: Position Model sheet — full column build + dynamic year breakdown

The largest task. Uses the generic cloner (Task 7) for the full per-day formula block (all populated columns, row pattern from Dec's row 3 — row 3 is the first row with real position-state formulas, unlike row 2 which has boundary-condition placeholders), then separately builds the dynamic per-calendar-year summary that replaces Dec's hardcoded 2021-2026 columns.

**Files:**
- Create: `scripts/build_position_model.py`
- Test: `scripts/tests/test_build_position_model.py`

**Interfaces:**
- Consumes: `fill_sheet_from_pattern` (Task 7).
- Produces: `write_position_model_sheet(wb, source_path, num_rows=3000)`.

- [ ] **Step 1: Determine the sheet's real column count**

Run:
```bash
py -c "
import openpyxl
wb = openpyxl.load_workbook(r'Johansen_FM_BOHO_D4_Strategy.xlsx', data_only=False)
ws = wb['Position Model']
print(ws.max_column, ws.dimensions)
"
```
Record the printed `max_column` value — use it as `MAX_COL` in Step 3 (this sheet's true width, don't guess).

- [ ] **Step 2: Write the failing test**

```python
# scripts/tests/test_build_position_model.py
import openpyxl, pathlib
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
```

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_position_model.py
import openpyxl
from sheet_cloner import fill_sheet_from_pattern
from config_map import CONSTANT_MAP
from formula_tools import extract_row_formulas


def write_position_model_sheet(wb, source_path: str, num_rows: int = 3000, max_col: int = None) -> None:
    if "Position Model" in wb.sheetnames:
        del wb["Position Model"]
    ws = wb.create_sheet("Position Model")

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Position Model"]
    if max_col is None:
        max_col = src.max_column

    # Header row, copied as-is (labels only, no formulas to rewrite).
    for cell in next(src.iter_rows(min_row=1, max_row=1, max_col=max_col)):
        if cell.value is not None:
            ws.cell(row=1, column=cell.column, value=cell.value)

    # Per-day formula block: pattern from row 3 (row 2 carries boundary-condition
    # placeholders in Dec, e.g. self-referencing differences, not representative
    # of the steady-state row-down formula).
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Position Model",
        pattern_row=3,
        target_start_row=3,
        num_rows=num_rows - 1,
        max_col=max_col,
        constant_map=CONSTANT_MAP,
    )
    # Row 2 filled from its own literal Dec pattern (boundary row, not shifted from row 3).
    row2 = extract_row_formulas(source_path, "Position Model", row=2, max_col=max_col)
    from formula_tools import rewrite_constants
    from openpyxl.utils import column_index_from_string
    for col_letter, value in row2.items():
        if isinstance(value, str) and value.startswith("="):
            value = rewrite_constants(value, CONSTANT_MAP)
        ws.cell(row=2, column=column_index_from_string(col_letter), value=value)

    _write_dynamic_year_breakdown(ws)


def _write_dynamic_year_breakdown(ws) -> None:
    """Replaces Dec's hardcoded per-year row-bound COUNT columns (e.g. AH$2:AH$249
    for 2021 specifically) with COUNTIFS keyed on this sheet's own Year column (CO),
    so a new calendar year needs no manual edit."""
    ws["EA2"] = "Pos1"
    ws["EB2"] = "Pos2"
    ws["EC2"] = "Pos3"
    ws["ED2"] = "Total"
    ws["DZ3"] = "=YEAR(A3)"
    ws["EA3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AH$3:$AH$3000,"<>")'
    ws["EB3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AJ$3:$AJ$3000,"<>")'
    ws["EC3"] = '=COUNTIFS($CO$3:$CO$3000,DZ3,$AL$3:$AL$3000,"<>")'
    ws["ED3"] = "=SUM(EA3:EC3)"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_position_model.py -v`
Expected: 2 passed. If `test_position_model_fills_and_rewrites_constants` fails because column `I` or `M`'s actual content differs from what's asserted, re-run Step 1's inspection command against the real file for those specific cells and correct the test's expected substring — do not weaken the assertion to pass artificially.

- [ ] **Step 5: Save checkpoint**

---

### Task 13: Results sheet

**Files:**
- Modify: `scripts/build_matrices_johansen.py` (add alongside the other two builders in this file, since it's the same "small fixed block, direct copy" shape)
- Test: `scripts/tests/test_build_matrices_johansen.py`

**Interfaces:**
- Produces: `write_results_sheet(wb, source_path)`.

- [ ] **Step 1: Write the failing test**

```python
# append to scripts/tests/test_build_matrices_johansen.py
from build_matrices_johansen import write_results_sheet

def test_results_sheet_replicated():
    wb = openpyxl.Workbook()
    write_results_sheet(wb, SOURCE_PATH)
    ws = wb["Results"]
    assert ws["A2"].value == "Effective T (\u0394 observations)"
    assert ws["B4"].value == "=Johansen!B12"
```

Before asserting `B4`'s exact formula, run:
```bash
py -c "
import openpyxl
wb = openpyxl.load_workbook(r'Johansen_FM_BOHO_D4_Strategy.xlsx', data_only=False)
ws = wb['Results']
for row in ws.iter_rows(min_row=1, max_row=11, max_col=2):
    print([c.value for c in row])
"
```
and correct the test's expected string for `B4` (and any other asserted cell) to match the real formula text found, not a guess.

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v -k results`
Expected: FAIL with `ImportError: cannot import name 'write_results_sheet'`

- [ ] **Step 3: Write minimal implementation**

```python
# append to scripts/build_matrices_johansen.py

def write_results_sheet(wb, source_path: str) -> None:
    if "Results" in wb.sheetnames:
        del wb["Results"]
    ws = wb.create_sheet("Results")
    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    src = source_wb["Results"]
    for row in src.iter_rows(min_row=1, max_row=11, max_col=2):
        for cell in row:
            if cell.value is not None:
                ws[cell.coordinate] = cell.value
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v`
Expected: all passed (3 tests total from Tasks 10+13)

- [ ] **Step 5: Save checkpoint**

---

### Task 14: Orchestration script — assemble the full template

**Files:**
- Create: `scripts/build_template.py`
- Test: `scripts/tests/test_build_template.py`

**Interfaces:**
- Consumes: every `write_*_sheet` function from Tasks 6, 8, 9, 10, 11, 12, 13.
- Produces: `build_template(source_path: str, output_path: str, num_rows: int = 3000) -> None`; a runnable `if __name__ == "__main__":` block.

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_build_template.py
import openpyxl, pathlib, tempfile, os
from build_template import build_template

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_build_template_creates_expected_sheets_in_order():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb.sheetnames == [
            "Notes", "Config", "Data", "DiffLag", "Matrices",
            "Johansen", "Johansen_Monthly", "Monthly Beta",
            "Position Model", "Results",
        ]

def test_build_template_data_sheet_has_no_price_values():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb["Data"]["B2"].value is None
        assert wb["Data"]["C2"].value is None

def test_build_template_no_leftover_dec_absolute_constants():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        ws = wb["Position Model"]
        for row in ws.iter_rows(min_row=2, max_row=11):
            for cell in row:
                if isinstance(cell.value, str):
                    assert "$DV$6" not in cell.value
                    assert "$DQ$7" not in cell.value
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_build_template.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/build_template.py
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_build_template.py -v`
Expected: 3 passed

- [ ] **Step 5: Run the full test suite to confirm no regressions across all prior tasks**

Run: `cd scripts && py -m pytest tests/ -v`
Expected: all tests from Tasks 2-14 pass.

- [ ] **Step 6: Generate the real template file**

Run: `cd scripts && py build_template.py`
Expected: `Template written to ...\Johansen_Template_Engine.xlsx`

- [ ] **Step 7: Save checkpoint**

---

### Task 15: Numeric validation against Dec's live output (Excel COM automation)

openpyxl never calculates formulas — it only ever shows the last value Excel cached. To prove the template's generated formulas produce the *same numbers* as Dec (not just structurally similar formulas), this task drives real Excel via COM to force a full recalculation and read back live results, using Dec's own historical data and Config values as the input to a throwaway copy of the template.

**Files:**
- Create: `scripts/validate_against_dec.py`
- Test: `scripts/tests/test_validate_against_dec.py`

**Interfaces:**
- Consumes: `build_template` (Task 14).
- Produces: `run_validation(template_path: str, dec_source_path: str, output_diff_path: str) -> list[dict]` — a list of row-level mismatches (empty list = match).

- [ ] **Step 1: Write the failing test**

```python
# scripts/tests/test_validate_against_dec.py
import pathlib
from validate_against_dec import run_validation

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()

def test_validation_reports_no_mismatches_on_dec_replay():
    mismatches = run_validation(
        template_path=str(pathlib.Path(__file__).parent.parent / "Johansen_Template_Engine.xlsx"),
        dec_source_path=SOURCE_PATH,
        output_diff_path=str(pathlib.Path(__file__).parent / "_validation_diff.csv"),
    )
    assert mismatches == [], f"{len(mismatches)} mismatched rows, see _validation_diff.csv"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_validate_against_dec.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'validate_against_dec'`

- [ ] **Step 3: Write minimal implementation**

```python
# scripts/validate_against_dec.py
import csv
import shutil
import openpyxl
import win32com.client
from config_map import CONFIG_LAYOUT

# Dec's live Config-equivalent values, read from the same cells Task 4's
# CONSTANT_MAP was built from, used to populate the throwaway template copy
# so it trades on identical parameters to Dec.
DEC_LIVE_VALUES = {
    "HOGalPerContract": 42000,
    "ZLMultiplier": 7.5,
    "JohansenWindowLen": 63,
    "ZThreshold_Tier1": 2.25,
    "ZThreshold_Tier2": 2.1,
    "ZThreshold_Tier3": 1.75,
    "ZExitThreshold": 0.1,
    "SizingWeight_Tier1": 0.25,
    "SizingWeight_Tier2": 0.5,
    "SizingWeight_Tier3": 0.25,
    "EWMA_Lambda": 0.94,
    "ShockMultiplierThresh": 1.25,
    "StopLossUSD_Base": -65000,
    "RiskLevelScalar": 1.5,
    "RiskLevelInput": 1,
    "ZLCentsPerLB": 60000,
}

COMPARE_COLUMNS = ["I", "L", "S", "T", "U", "V", "X"]  # spread, z-score, positions 1-3, total, daily PnL


def _populate_throwaway_copy(template_path: str, dec_source_path: str, throwaway_path: str) -> None:
    shutil.copy(template_path, throwaway_path)
    wb = openpyxl.load_workbook(throwaway_path, data_only=False)
    config_ws = wb["Config"]
    layout_by_name = {name: cell for name, cell, _default in CONFIG_LAYOUT}
    for name, value in DEC_LIVE_VALUES.items():
        cell = layout_by_name[name]
        config_ws[cell] = value
    dec_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    dec_data = dec_wb["Data"]
    data_ws = wb["Data"]
    for row in dec_data.iter_rows(min_row=2, max_row=dec_data.max_row, max_col=20):
        for cell in row:
            if cell.value is not None:
                data_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    wb.save(throwaway_path)


def _recalculate_with_excel(path: str) -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(path)
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()


def run_validation(template_path: str, dec_source_path: str, output_diff_path: str) -> list:
    throwaway_path = template_path.replace(".xlsx", "_validation_copy.xlsx")
    _populate_throwaway_copy(template_path, dec_source_path, throwaway_path)
    _recalculate_with_excel(throwaway_path)
    _recalculate_with_excel(dec_source_path)  # ensure Dec's own cache is also current

    template_wb = openpyxl.load_workbook(throwaway_path, data_only=True)
    dec_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    template_pm = template_wb["Position Model"]
    dec_pm = dec_wb["Position Model"]

    mismatches = []
    max_row = min(dec_pm.max_row, template_pm.max_row)
    for row in range(3, max_row + 1):
        for col in COMPARE_COLUMNS:
            dec_val = dec_pm[f"{col}{row}"].value
            tmpl_val = template_pm[f"{col}{row}"].value
            if isinstance(dec_val, (int, float)) and isinstance(tmpl_val, (int, float)):
                if abs(dec_val - tmpl_val) > 1e-6:
                    mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val})
            elif dec_val != tmpl_val:
                mismatches.append({"row": row, "col": col, "dec": dec_val, "template": tmpl_val})

    if mismatches:
        with open(output_diff_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["row", "col", "dec", "template"])
            writer.writeheader()
            writer.writerows(mismatches)

    return mismatches
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd scripts && py -m pytest tests/test_validate_against_dec.py -v -s`
Expected: PASS with 0 mismatches. If it fails, open `scripts/_validation_diff.csv` to see exactly which rows/columns diverge — this points directly at a formula-translation defect in one of Tasks 8-13, not a genuine strategy difference, since both workbooks are fed identical data and identical parameters. Fix the offending sheet builder, regenerate the template (Task 14 Step 6), and re-run this validation before proceeding.

- [ ] **Step 5: Save checkpoint**

- [ ] **Step 6: Confirm the live December workbook was never modified**

Run: `py -c "import openpyxl; wb=openpyxl.load_workbook(r'Johansen_FM_BOHO_D4_Strategy.xlsx', data_only=True); print(wb['Position Model']['A3'].value, wb['Position Model']['I3'].value)"` and confirm these values are unchanged from what was observed at the start of this project (the source file may have been recalculated in-place by Step 3's `_recalculate_with_excel` call on `dec_source_path` — if so, restore it from a fresh, untouched copy before treating this task as complete, since Global Constraints forbid writing into the live source file).

---

## Definition of done

- `Johansen_Template_Engine.xlsx` exists in the `BOHO Trading` folder with the 9 sheets in the specified order, Config-driven throughout, no scattered literal constants remaining in Position Model or Matrices.
- Johansen_Monthly's Active End Row is a self-updating formula.
- Position Model's per-calendar-year breakdown is dynamic (COUNTIFS/SUMIFS-keyed, no hardcoded per-year row bounds).
- `py -m pytest scripts/tests/ -v` passes in full.
- `run_validation(...)` returns an empty mismatch list against Dec's real historical data and parameters.
- The live December workbook is provably unmodified.
