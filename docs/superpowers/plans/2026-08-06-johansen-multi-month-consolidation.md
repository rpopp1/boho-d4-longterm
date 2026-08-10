# Johansen Multi-Month BOHO/D4 Consolidation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the per-contract-month "Save-As a new workbook" template process with one consolidated workbook: a single `Data` tab covering all 8 BOHO contract months, shared engine tabs (`DiffLag`, `Matrices_Johansen_Results`, `Johansen_Monthly`, `Monthly Beta`) laid out as 8 side-by-side per-month column-blocks, and 8 independent `Position Model` tabs (one per month) that a trader flips between.

**Architecture:** Every per-day/per-test builder script (`build_diff_lag.py`, `build_matrices_johansen.py`, `build_johansen_monthly.py`, `build_monthly_beta.py`) is parameterized to write one month's column-block at a computed column offset instead of always starting at column A, using two new primitives: `multi_month.py` (the 8-month list plus column-offset arithmetic) and `formula_tools.shift_formula_columns` (a new sibling to the existing `shift_formula_rows`, which relocates column letters in cell/range/full-column references — including cross-sheet ones — by a computed delta). `Config` drops workbook-scoped defined names entirely in favor of one column per month; every formula that used to read a defined name now reads that month's `Config` column directly. `Position Model` tabs are NOT column-shifted (each stays a full standalone sheet starting at column A, exactly like today's single-month sheet) — they get a single post-build pass that relocates only their *cross-sheet* references (to `Data` and `Johansen_Monthly`) to the correct month's block.

**Tech Stack:** Python 3.14, `openpyxl` (workbook generation), `pywin32`/COM (Excel-driven validation recalculation), `pytest`.

## Global Constraints

- Never write to the live source file `Johansen_FM_BOHO_D4_Strategy.xlsx`. Verify its mtime (`1785880251.4742322`) and sha256 (`09eef5f8bfc4fd81f4b94e8822b7264cbcb7d1352d1db01f03ab31bb08b6f9d8`) are unchanged after any session touching this code.
- Never commit `.xlsx`/`.xls`/`.xlsm` files (already gitignored).
- This plan executes in an isolated git worktree/branch (the user chose this explicitly over the project's earlier no-worktree convention, given this plan's size) — merge back to `master` only via `finishing-a-development-branch` once the final review is clean.
- No fix for the known hardcoded-Dec-row-boundary defect in `Position Model`'s summary/Sharpe panel (`docs/superpowers/HANDOFF.md` punch-list item 1), nor items 2–4 (the `#N/A` summary cells, header label inconsistencies, validation-harness blind spots). All four are explicitly deferred per the design spec — do not fix them as part of this plan.
- No change to the underlying Johansen/z-score/position-sizing/PnL statistical logic.
- Dynamic-array/spill formulas cannot be reproduced via `openpyxl` — never introduce one; keep the existing `NotImplementedError` guards (`sheet_cloner.fill_sheet_from_pattern`, `formula_tools.unwrap_scalar_array_formula`) intact and unweakened.
- Prefer real Excel-COM recalculation (`validate_against_dec.py`) over `openpyxl`-only checks for anything that depends on how a formula actually evaluates — `openpyxl` never calculates formulas.
- Run `cd scripts && py -m pytest tests/ -v` after every task; it must be green (aside from the COM-recalculation tests, which take ~10–15 minutes and only need a full run before the final commit of Tasks 13/14 — run the non-COM unit tests after every task, and the full suite including COM tests at least once before Task 13/14 are considered done).

---

### Task 1: `shift_formula_columns` — the column-relocation primitive

**Files:**
- Modify: `scripts/formula_tools.py`
- Test: `scripts/tests/test_formula_tools.py`

**Interfaces:**
- Produces: `shift_formula_columns(formula: str, sheet_map: dict[str | None, tuple[int, str | None]]) -> str`. `sheet_map` keys are the qualifying sheet name as it appears in the source formula (or `None` for an unqualified/same-sheet reference); values are `(col_delta, new_sheet_name_or_None)`. A reference whose qualifier is not a key in `sheet_map` is returned completely unchanged. `new_sheet_name_or_None = None` writes the reference unqualified in the output (used both for genuine same-sheet references and for references crossing into a sheet that has since been merged into the formula's own sheet).

- [ ] **Step 1: Write the failing tests**

Append to `scripts/tests/test_formula_tools.py`:

```python
def test_shift_columns_unqualified_self_reference():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=B3+B6", {None: (2, None)}) == "=D3+D6"

def test_shift_columns_cross_sheet_reference_keeps_qualifier():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=Johansen_Monthly!$K$3", {"Johansen_Monthly": (13, "Johansen_Monthly")}
    ) == "=Johansen_Monthly!$X$3"

def test_shift_columns_full_column_reference():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=DiffLag!E:E", {"DiffLag": (5, "DiffLag")}) == "=DiffLag!J:J"

def test_shift_columns_strips_qualifier_after_sheet_merge():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=Matrices!C50", {"Matrices": (10, None)}) == "=M50"

def test_shift_columns_renames_qualifier_after_sheet_merge():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns(
        "=Johansen!D18", {"Johansen": (10, "Matrices_Johansen_Results")}
    ) == "=Matrices_Johansen_Results!N18"

def test_shift_columns_leaves_unmapped_sheet_untouched():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=Config!C14", {"DiffLag": (5, "DiffLag")}) == "=Config!C14"

def test_shift_columns_leaves_defined_name_tokens_untouched():
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=ZThreshold_Tier1*2", {None: (10, None)}) == "=ZThreshold_Tier1*2"

def test_shift_columns_ignores_dollar_anchoring():
    # Unlike shift_formula_rows, $ never blocks a column shift here -- see the
    # shift_formula_columns docstring for why.
    from formula_tools import shift_formula_columns
    assert shift_formula_columns("=$B$5", {None: (3, None)}) == "=$E$5"

def test_shift_columns_real_matrices_formula():
    # Verbatim from the live source workbook's Matrices!C7 (read-only inspection).
    from formula_tools import shift_formula_columns
    formula = (
        '=IF($B$5=0,"",SUMPRODUCT(INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2), '
        'INDEX(DiffLag!E:E,$B$4):INDEX(DiffLag!E:E,$B$2))/$B$5)'
    )
    result = shift_formula_columns(formula, {None: (10, None), "DiffLag": (5, "DiffLag")})
    assert result == (
        '=IF($L$5=0,"",SUMPRODUCT(INDEX(DiffLag!J:J,$L$4):INDEX(DiffLag!J:J,$L$2), '
        'INDEX(DiffLag!J:J,$L$4):INDEX(DiffLag!J:J,$L$2))/$L$5)'
    )
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v -k shift_columns`
Expected: FAIL with `ImportError` / `AttributeError` — `shift_formula_columns` doesn't exist yet.

- [ ] **Step 3: Implement `shift_formula_columns` in `formula_tools.py`**

Add `column_index_from_string` to the existing `openpyxl.utils` import line (it currently only imports `get_column_letter`), then append:

```python
_SHEET_QUALIFIER = r"(?:'(?P<qname>[^']+)'|(?P<uname>[A-Za-z_][A-Za-z0-9_.]*))!"

_COL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:" + _SHEET_QUALIFIER + r")?"
    r"(?:"
    r"(?P<colabs>\$)?(?P<col>[A-Za-z]{1,3})(?P<rowabs>\$)?(?P<row>\d+)"
    r"|"
    r"(?P<fcol1>[A-Za-z]{1,3}):(?P<fcol2>[A-Za-z]{1,3})"
    r")"
    r"(?![A-Za-z0-9_(])"
)

_NEEDS_QUOTING_RE = re.compile(r"[^A-Za-z0-9_]")


def _qualify_sheet(name: str | None) -> str:
    if name is None:
        return ""
    if _NEEDS_QUOTING_RE.search(name):
        return f"'{name}'!"
    return f"{name}!"


def shift_formula_columns(formula: str, sheet_map: dict) -> str:
    """Relocate the COLUMN portion of every cell/range/full-column reference
    in `formula` whose qualifying sheet name (or None, for an unqualified/
    same-sheet reference) is a key in `sheet_map`.

    sheet_map: {qualifier_or_None: (col_delta, new_sheet_name_or_None)}.
    A reference whose qualifier is NOT a key in sheet_map (e.g. a Config!
    reference, when "Config" isn't in the map) is left byte-for-byte
    unchanged -- this is how Config references pass through untouched
    during the multi-month column-relocation pass.

    Unlike shift_formula_rows, this ignores $ anchoring entirely: $ was
    only ever used in this codebase to protect a reference from *row*
    shifting when a formula is filled down many rows (see
    shift_formula_rows). Column shifting here is applied exactly once per
    formula -- relocating a whole per-month block to its column position --
    so there is no "moving target" to protect against, and every column
    reference into a relocated sheet must move, anchored or not.
    """
    def _replace(match: "re.Match") -> str:
        sheet_key = match.group("qname") or match.group("uname")
        if sheet_key not in sheet_map:
            return match.group(0)
        col_delta, new_sheet = sheet_map[sheet_key]
        prefix = _qualify_sheet(new_sheet)

        if match.group("col"):
            colabs = match.group("colabs") or ""
            rowabs = match.group("rowabs") or ""
            new_col = get_column_letter(column_index_from_string(match.group("col")) + col_delta)
            return f"{prefix}{colabs}{new_col}{rowabs}{match.group('row')}"

        new_col1 = get_column_letter(column_index_from_string(match.group("fcol1")) + col_delta)
        new_col2 = get_column_letter(column_index_from_string(match.group("fcol2")) + col_delta)
        return f"{prefix}{new_col1}:{new_col2}"

    return _COL_REF_RE.sub(_replace, formula)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_formula_tools.py -v`
Expected: PASS (all tests in the file, including the pre-existing ones — confirms nothing regressed).

- [ ] **Step 5: Commit**

```bash
git add scripts/formula_tools.py scripts/tests/test_formula_tools.py
git commit -m "Add shift_formula_columns primitive for multi-month column relocation"
```

---

### Task 2: `multi_month.py` — month list and column-offset arithmetic

**Files:**
- Create: `scripts/multi_month.py`
- Test: `scripts/tests/test_multi_month.py`

**Interfaces:**
- Produces: `MONTHS: list[str]`, `GAP: int`, `DATA_BLOCK_WIDTH/DIFFLAG_BLOCK_WIDTH/JOHANSEN_MONTHLY_BLOCK_WIDTH/MONTHLY_BETA_BLOCK_WIDTH: int`, `MATRICES_BLOCK_WIDTH/JOHANSEN_BLOCK_WIDTH/RESULTS_BLOCK_WIDTH/MJR_SUB_GAP/MJR_MONTH_GAP/MJR_BLOCK_WIDTH: int`, `month_index(month: str) -> int`, `column_offset(month: str, block_width: int, gap: int = GAP) -> int`, `mjr_offsets(month: str) -> dict[str, int]`, `shift_col(col_letter: str, delta: int) -> str`.

These block widths are not arbitrary — they were measured by read-only `openpyxl` inspection of the live source workbook: `Data` A:T = 20 cols, `DiffLag`/`Monthly Beta` A:H = 8 cols each, `Matrices` A:H = 8 cols, `Johansen` A:F = 6 cols, `Results` A:B = 2 cols, `Johansen_Monthly` A:K (with blank spacer columns E and J already built in) = 11 cols.

- [ ] **Step 1: Write the failing tests**

Create `scripts/tests/test_multi_month.py`:

```python
from multi_month import (
    MONTHS, GAP, month_index, column_offset, mjr_offsets, shift_col,
    DATA_BLOCK_WIDTH, DIFFLAG_BLOCK_WIDTH,
)


def test_months_in_calendar_order():
    assert MONTHS == [
        "January", "March", "May", "July", "August", "September", "October", "December",
    ]


def test_month_index():
    assert month_index("January") == 0
    assert month_index("December") == 7


def test_column_offset_january_is_zero_for_every_block_width():
    # January must land exactly where today's single-month template already
    # puts things -- a load-bearing sanity property for every later task.
    assert column_offset("January", DIFFLAG_BLOCK_WIDTH, GAP) == 0
    assert column_offset("January", DATA_BLOCK_WIDTH, GAP) == 0


def test_column_offset_march_diff_lag():
    assert column_offset("March", DIFFLAG_BLOCK_WIDTH, GAP) == 10  # 1 * (8 + 2)


def test_column_offset_december_data():
    assert column_offset("December", DATA_BLOCK_WIDTH, GAP) == 154  # 7 * (20 + 2)


def test_mjr_offsets_january():
    assert mjr_offsets("January") == {"Matrices": 0, "Johansen": 10, "Results": 18}


def test_mjr_offsets_march():
    # March's mega-block starts at 1 * (20 + 4) = 24
    assert mjr_offsets("March") == {"Matrices": 24, "Johansen": 34, "Results": 42}


def test_shift_col():
    assert shift_col("B", 10) == "L"
    assert shift_col("E", 5) == "J"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_multi_month.py -v`
Expected: FAIL — `multi_month` module doesn't exist.

- [ ] **Step 3: Implement `scripts/multi_month.py`**

```python
from openpyxl.utils import get_column_letter, column_index_from_string

MONTHS = ["January", "March", "May", "July", "August", "September", "October", "December"]

GAP = 2

DATA_BLOCK_WIDTH = 20
DIFFLAG_BLOCK_WIDTH = 8
JOHANSEN_MONTHLY_BLOCK_WIDTH = 11
MONTHLY_BETA_BLOCK_WIDTH = 8

MATRICES_BLOCK_WIDTH = 8
JOHANSEN_BLOCK_WIDTH = 6
RESULTS_BLOCK_WIDTH = 2
MJR_SUB_GAP = 2
MJR_MONTH_GAP = 4
MJR_BLOCK_WIDTH = (
    MATRICES_BLOCK_WIDTH + MJR_SUB_GAP + JOHANSEN_BLOCK_WIDTH + MJR_SUB_GAP + RESULTS_BLOCK_WIDTH
)


def month_index(month: str) -> int:
    return MONTHS.index(month)


def column_offset(month: str, block_width: int, gap: int = GAP) -> int:
    """How many columns to add to a block's original (single-month-template-
    style, starting at column A) column indices so that `month`'s copy of
    the block lands at the correct position, given every month's block is
    `block_width` columns wide with `gap` blank columns between consecutive
    months' blocks."""
    return month_index(month) * (block_width + gap)


def mjr_offsets(month: str) -> dict:
    """Column offsets (see column_offset) for each of the three sub-blocks
    making up `month`'s mega-block on the combined Matrices_Johansen_Results
    tab. Each offset is relative to that sub-block's own original
    column-A-start layout (Matrices A-H, Johansen A-F, Results A-B)."""
    mega_offset = column_offset(month, MJR_BLOCK_WIDTH, MJR_MONTH_GAP)
    matrices_offset = mega_offset
    johansen_offset = mega_offset + MATRICES_BLOCK_WIDTH + MJR_SUB_GAP
    results_offset = johansen_offset + JOHANSEN_BLOCK_WIDTH + MJR_SUB_GAP
    return {"Matrices": matrices_offset, "Johansen": johansen_offset, "Results": results_offset}


def shift_col(col_letter: str, delta: int) -> str:
    return get_column_letter(column_index_from_string(col_letter) + delta)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_multi_month.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/multi_month.py scripts/tests/test_multi_month.py
git commit -m "Add multi_month.py: 8-month list and column-offset arithmetic"
```

---

### Task 3: Per-month `Config` layout in `config_map.py` (additive)

**Files:**
- Modify: `scripts/config_map.py`
- Test: `scripts/tests/test_config_map.py` (new)

**Interfaces:**
- Consumes: `multi_month.MONTHS`
- Produces (all ADDED alongside the existing `CONFIG_LAYOUT`/`CONSTANT_MAP`/`MATRICES_ONLY_CONSTANT_MAP`, which stay untouched for now — see the note below): `CONFIG_FIELDS: list[tuple[str, int, object]]` (name, row, default — same content as `CONFIG_LAYOUT` minus the single-column cell address), `CONFIG_MONTH_COLUMNS: dict[str, str]` (month name -> Config column letter, `B` for January through `I` for December), `config_cell(month: str, row: int) -> str`, `constant_map_for_month(month: str) -> dict`, `matrices_only_constant_map_for_month(month: str) -> dict`.

**Note on why this is additive, not a replacement:** `CONFIG_LAYOUT`/`CONSTANT_MAP`/`MATRICES_ONLY_CONSTANT_MAP` are still imported by `build_config_sheet.py`, `build_diff_lag.py`, `build_monthly_beta.py`, `build_matrices_johansen.py`, `build_johansen_monthly.py`, and `build_position_model.py` — none of which have been migrated yet. Removing the old names now would break every one of those imports before Tasks 5–11 get a chance to migrate their consumers one at a time. The old names are removed as the last step of Task 11, once nothing references them anymore.

- [ ] **Step 1: Write the failing tests**

Create `scripts/tests/test_config_map.py`:

```python
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS, config_cell, constant_map_for_month, matrices_only_constant_map_for_month


def test_config_month_columns():
    assert CONFIG_MONTH_COLUMNS["January"] == "B"
    assert CONFIG_MONTH_COLUMNS["December"] == "I"


def test_config_cell():
    assert config_cell("January", 14) == "B14"
    assert config_cell("August", 14) == "F14"


def test_config_fields_matches_original_layout_rows():
    # Same fields/rows/defaults as the single-month CONFIG_LAYOUT, just
    # without a hardcoded single-column cell address.
    by_name = {name: (row, default) for name, row, default in CONFIG_FIELDS}
    assert by_name["ZThreshold_Tier1"] == (14, 2.25)
    assert by_name["ContractLabel"] == (2, "TEMPLATE")
    assert by_name["StopLossUSD_Base"] == (27, -65000)


def test_constant_map_for_month_points_at_that_months_config_column():
    jan_map = constant_map_for_month("January")
    aug_map = constant_map_for_month("August")
    assert jan_map["$DV$6"] == "Config!B17"       # ZExitThreshold, January column B
    assert aug_map["$DV$6"] == "Config!F17"        # ZExitThreshold, August column F
    assert jan_map["$DK$5"] == "Config!B7"         # HOGalPerContract


def test_matrices_only_constant_map_for_month():
    assert matrices_only_constant_map_for_month("January") == {"$B$3": "Config!B11"}
    assert matrices_only_constant_map_for_month("December") == {"$B$3": "Config!I11"}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_config_map.py -v`
Expected: FAIL — new names don't exist yet.

- [ ] **Step 3: Add the new per-month structures to `config_map.py`**

Append to `scripts/config_map.py` (leave every existing line untouched):

```python
from openpyxl.utils import get_column_letter
from multi_month import MONTHS

# Per-month Config layout: same fields/rows/defaults as CONFIG_LAYOUT above,
# but with every field independently valued per contract month (one column
# per month) instead of a single shared column + workbook-scoped defined
# name. CONFIG_LAYOUT/CONSTANT_MAP/MATRICES_ONLY_CONSTANT_MAP above are
# unused by the multi-month template once Tasks 5-11 finish migrating their
# consumers; they are removed as the last step of Task 11.
CONFIG_FIELDS = [(name, int("".join(ch for ch in cell if ch.isdigit())), default)
                  for name, cell, default in CONFIG_LAYOUT]

CONFIG_MONTH_COLUMNS = {month: get_column_letter(2 + i) for i, month in enumerate(MONTHS)}

_CONFIG_ROW_BY_FIELD = {name: row for name, row, _default in CONFIG_FIELDS}

# Dec absolute Position-Model-local-param-block address -> defined name,
# reconciled against the same live-workbook inspection CONSTANT_MAP above
# was built from. Kept as a private lookup so constant_map_for_month below
# only needs to know "which Config row", not re-derive the address list.
_ADDRESS_TO_FIELD = {
    "$DK$2": "ZLCentsPerLB", "$DK$5": "HOGalPerContract", "$DK$3": "ZLMultiplier",
    "$DV$3": "ZThreshold_Tier1", "$DV$4": "ZThreshold_Tier2", "$DV$5": "ZThreshold_Tier3",
    "$DV$6": "ZExitThreshold", "$DL$7": "SizingWeight_Tier1", "$DM$7": "SizingWeight_Tier2",
    "$DN$7": "SizingWeight_Tier3", "$DQ$7": "EWMA_Lambda", "$DQ$8": "ShockMultiplierThreshold",
    "$DQ$10": "StopLossUSD_Base", "$DK$11": "D4Multiplier", "$DQ$11": "RiskLevelInput",
}


def config_cell(month: str, row: int) -> str:
    return f"{CONFIG_MONTH_COLUMNS[month]}{row}"


def constant_map_for_month(month: str) -> dict:
    """Dec absolute address -> that month's Config sheet cell reference
    (e.g. "Config!F17"), replacing CONSTANT_MAP's address -> defined-name
    mapping now that every Config field is independently valued per month
    instead of a single shared workbook-scoped name."""
    return {
        address: f"Config!{config_cell(month, _CONFIG_ROW_BY_FIELD[field])}"
        for address, field in _ADDRESS_TO_FIELD.items()
    }


def matrices_only_constant_map_for_month(month: str) -> dict:
    return {"$B$3": f"Config!{config_cell(month, _CONFIG_ROW_BY_FIELD['JohansenWindowLen'])}"}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_config_map.py tests/test_formula_tools.py -v`
Expected: PASS — including the pre-existing formula_tools tests, confirming the additive change broke nothing.

- [ ] **Step 5: Commit**

```bash
git add scripts/config_map.py scripts/tests/test_config_map.py
git commit -m "Add per-month Config layout to config_map.py (additive)"
```

---

### Task 4: `fill_sheet_from_pattern` column relocation support (backward-compatible)

**Files:**
- Modify: `scripts/sheet_cloner.py`
- Test: `scripts/tests/test_sheet_cloner.py`

**Interfaces:**
- Modifies: `fill_sheet_from_pattern(target_ws, source_path, source_sheet, pattern_row, target_start_row, num_rows, max_col, constant_map, column_shift_map: dict | None = None) -> None`. New optional parameter, default `None`, preserving today's exact behavior when omitted. When provided, `column_shift_map[None]`'s delta (default 0 if the key is absent) is where the block is physically written (added to the source's own column index), and the same map is passed through `shift_formula_columns` on every formula's text, so a block's own self-references move together with its physical placement.

- [ ] **Step 1: Write the failing tests**

Append to `scripts/tests/test_sheet_cloner.py`:

```python
def test_fill_with_column_shift_relocates_block_and_self_references():
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=3,
        max_col=8,
        constant_map=CONSTANT_MAP,
        column_shift_map={None: (10, None), "Data": (5, "Data")},
    )
    # DiffLag's own column B (source: '=IF(Data!B2="","",Data!A2)') physically
    # lands 10 columns over, at column L (B=2+10=12); its "Data!" cross-refs
    # shift by Data's own delta (5), not DiffLag's (10).
    assert ws["L2"].value == '=IF(Data!G2="","",Data!F2)'
    # DiffLag's own column E (source: '=IF(OR(C2="",C1=""),"",C2-C1)') is a
    # self-reference to DiffLag's own C column (unqualified -- no "Data!").
    # It physically lands at E(5)+10=15='O', and its self-reference shifts
    # by DiffLag's own delta (10): C(3)+10=13='M'.
    assert ws["O2"].value == '=IF(OR(M2="",M1=""),"",M2-M1)'
    assert ws["O3"].value == '=IF(OR(M3="",M2=""),"",M3-M2)'


def test_fill_without_column_shift_map_is_unchanged():
    # column_shift_map omitted entirely -- must behave exactly as before.
    wb = openpyxl.Workbook()
    ws = wb.active
    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=SOURCE_PATH,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=3,
        max_col=8,
        constant_map=CONSTANT_MAP,
    )
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_sheet_cloner.py -v`
Expected: `test_fill_with_column_shift_relocates_block_and_self_references` FAILs with `TypeError` (unexpected keyword argument); the second test passes already (it's exercising unchanged behavior).

- [ ] **Step 3: Modify `fill_sheet_from_pattern`**

In `scripts/sheet_cloner.py`, add the import and update the function:

```python
from formula_tools import extract_row_formulas, shift_formula_rows, rewrite_constants, shift_formula_columns


def fill_sheet_from_pattern(
    target_ws,
    source_path: str,
    source_sheet: str,
    pattern_row: int,
    target_start_row: int,
    num_rows: int,
    max_col: int,
    constant_map: dict,
    column_shift_map: dict | None = None,
) -> None:
    pattern = extract_row_formulas(source_path, source_sheet, pattern_row, max_col)
    spill_cols = _get_array_formula_columns(source_path, source_sheet, pattern_row, max_col)
    if spill_cols:
        raise NotImplementedError(
            f"Pattern row {pattern_row} of sheet '{source_sheet}' contains "
            f"dynamic-array spill anchor formula(s) in column(s) {sorted(spill_cols)}. "
            "openpyxl cannot reliably reproduce Excel's dynamic-array spill "
            "semantics (no supported way to write the cm=\"1\" cell-metadata "
            "Excel needs to recognize a true auto-resizing spill, vs. a "
            "legacy fixed-size CSE array formula that pads excess cells with "
            "#N/A) — this exact failure mode was found and traced by Task 15's "
            "Dec-replay validation. Redesign the sheet builder to avoid spill "
            "formulas entirely (e.g. a plain per-row mirror formula, or a "
            "manually-maintained column) rather than special-casing this here "
            "— see Task 11 v2's redesign of Johansen_Monthly for the pattern."
        )

    self_offset = (column_shift_map or {}).get(None, (0, None))[0]

    for i in range(num_rows):
        target_row = target_start_row + i
        row_delta = target_row - pattern_row
        for col_letter, value in pattern.items():
            from openpyxl.utils import column_index_from_string, get_column_letter
            target_col_letter = get_column_letter(column_index_from_string(col_letter) + self_offset)
            if isinstance(value, str) and value.startswith("="):
                shifted = shift_formula_rows(value, row_delta)
                rewritten = rewrite_constants(shifted, constant_map)
                if column_shift_map:
                    rewritten = shift_formula_columns(rewritten, column_shift_map)
                target_ws[f"{target_col_letter}{target_row}"] = rewritten
            else:
                target_ws[f"{target_col_letter}{target_row}"] = value
```

(Move the `from openpyxl.utils import ...` line to the top of the file alongside the existing `openpyxl`/`ArrayFormula`/`get_column_letter` imports rather than inline inside the loop — inline above only to show exactly what's newly needed; consolidate on write.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_sheet_cloner.py tests/test_build_diff_lag.py tests/test_build_position_model.py -v`
Expected: PASS — including the pre-existing DiffLag/Position Model tests, confirming `column_shift_map=None` (the default every existing caller uses) reproduces today's exact behavior.

- [ ] **Step 5: Commit**

```bash
git add scripts/sheet_cloner.py scripts/tests/test_sheet_cloner.py
git commit -m "Add optional column relocation to fill_sheet_from_pattern (backward-compatible)"
```

---

### Task 5: Rebuild `build_config_sheet.py` — 8-column Config, no defined names

**Files:**
- Modify: `scripts/build_config_sheet.py`
- Modify: `scripts/tests/test_build_config_sheet.py`

**Interfaces:**
- Consumes: `multi_month.MONTHS`, `config_map.CONFIG_FIELDS`, `config_map.CONFIG_MONTH_COLUMNS`
- Modifies: `write_config_sheet(wb) -> None` — same signature, now writes one column per month with no defined names.

- [ ] **Step 1: Replace the failing tests**

Replace the contents of `scripts/tests/test_build_config_sheet.py`:

```python
import openpyxl
from build_config_sheet import write_config_sheet
from multi_month import MONTHS
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS


def test_config_sheet_has_month_header_row():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    assert ws["A1"].value == "Field"
    assert ws["B1"].value == "January"
    assert ws["I1"].value == "December"


def test_config_sheet_every_month_gets_the_default_value():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for month in MONTHS:
        col = CONFIG_MONTH_COLUMNS[month]
        assert ws[f"{col}14"].value == 2.25  # ZThreshold_Tier1 default, every month


def test_config_sheet_identity_fields_default_to_template_placeholder():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for month in MONTHS:
        col = CONFIG_MONTH_COLUMNS[month]
        assert ws[f"{col}2"].value == "TEMPLATE"  # ContractLabel


def test_config_sheet_no_defined_names():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    assert len(wb.defined_names) == 0


def test_config_sheet_labels_column_a():
    wb = openpyxl.Workbook()
    write_config_sheet(wb)
    ws = wb["Config"]
    for name, row, _default in CONFIG_FIELDS:
        assert ws[f"A{row}"].value == name
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_config_sheet.py -v`
Expected: FAIL — current `write_config_sheet` still writes the single-column/defined-name layout.

- [ ] **Step 3: Rewrite `build_config_sheet.py`**

```python
import openpyxl
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS
from multi_month import MONTHS


def write_config_sheet(wb: "openpyxl.Workbook") -> None:
    if "Config" in wb.sheetnames:
        del wb["Config"]
    ws = wb.create_sheet("Config")
    ws["A1"] = "Field"
    for month in MONTHS:
        ws[f"{CONFIG_MONTH_COLUMNS[month]}1"] = month
    for name, row, default in CONFIG_FIELDS:
        ws[f"A{row}"] = name
        for month in MONTHS:
            ws[f"{CONFIG_MONTH_COLUMNS[month]}{row}"] = default
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_config_sheet.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_config_sheet.py scripts/tests/test_build_config_sheet.py
git commit -m "Rebuild Config sheet with one independently-calibrated column per month"
```

---

### Task 6: Rebuild `build_static_sheets.py` — 8-block `Data`, updated `Notes`

**Files:**
- Modify: `scripts/build_static_sheets.py`
- Modify: `scripts/tests/test_build_static_sheets.py`

**Interfaces:**
- Consumes: `multi_month.{MONTHS, column_offset, DATA_BLOCK_WIDTH, GAP, shift_col}`
- Modifies: `write_data_sheet(wb, month: str) -> None` — now takes an explicit `month` and writes only that month's block; get-or-create the `Data` sheet rather than always deleting it, since it's called once per month.
- Unmodified signature: `write_notes_sheet(wb) -> None` (text content changes only).

- [ ] **Step 1: Write the failing tests**

Replace the contents of `scripts/tests/test_build_static_sheets.py`:

```python
import openpyxl
from build_static_sheets import write_notes_sheet, write_data_sheet
from multi_month import MONTHS


def test_notes_sheet_describes_multi_month_workflow():
    wb = openpyxl.Workbook()
    write_notes_sheet(wb)
    ws = wb["Notes"]
    text = "\n".join(str(ws[f"A{i}"].value) for i in range(1, ws.max_row + 1) if ws[f"A{i}"].value)
    assert "Config" in text
    assert "one column per contract month" in text.lower()


def test_data_sheet_january_lands_at_column_a():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    ws = wb["Data"]
    assert ws["A1"].value == "Date"
    assert ws["B1"].value == "BOHO"
    assert ws["E1"].value == "Settlement Pricing"
    assert ws["K1"].value == "Open Pricing"
    assert ws["T3"].value == "=T2/T3" or ws["T4"].value == "=T2/T3"  # ZL Gal formula, see Step 3


def test_data_sheet_march_lands_at_computed_offset():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    # March offset = 1 * (20 + 2) = 22 columns -> Date lands at column A+22 = "W"
    assert ws["W1"].value == "Date"
    assert ws["X1"].value == "BOHO"


def test_data_sheet_month_label_in_gap_after_each_block():
    # A plain text label in row 1 of the first gap column right after each
    # month's block, so it's visually obvious where one month's block ends
    # and the next begins (per the design spec's per-month-block layout).
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    jan_label_col = shift_col("A", column_offset("January", DATA_BLOCK_WIDTH, GAP) + DATA_BLOCK_WIDTH)
    mar_label_col = shift_col("A", column_offset("March", DATA_BLOCK_WIDTH, GAP) + DATA_BLOCK_WIDTH)
    assert ws[f"{jan_label_col}1"].value == "January"
    assert ws[f"{mar_label_col}1"].value == "March"


def test_data_sheet_accumulates_across_months():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "January")
    write_data_sheet(wb, "March")
    ws = wb["Data"]
    assert ws["A1"].value == "Date"   # January's block is still there
    assert ws["W1"].value == "Date"   # March's block was added, not a fresh sheet


def test_data_sheet_month_header_row():
    wb = openpyxl.Workbook()
    write_data_sheet(wb, "August")
    ws = wb["Data"]
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
    offset = column_offset("August", DATA_BLOCK_WIDTH, GAP)
    header_col = shift_col("A", offset)
    assert ws[f"{header_col}{0}" if False else f"{header_col}1"].value == "Date"
```

(Drop the last, redundant assertion in `test_data_sheet_month_header_row` if it duplicates `test_data_sheet_march_lands_at_computed_offset`'s intent — the point of this test is just to exercise the shared `column_offset`/`shift_col` helpers directly against the sheet, so simplify to a single clean assertion using those helpers instead of a hardcoded letter, unlike the two tests above it which intentionally hardcode the expected letters as an independent cross-check.)

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_static_sheets.py -v`
Expected: FAIL — `write_data_sheet` doesn't take a `month` argument yet.

- [ ] **Step 3: Rewrite `build_static_sheets.py`**

```python
from multi_month import MONTHS, column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
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
        "   weights, and risk parameters. Every field has one column per contract",
        "   month, independently calibrated -- there are no shared/global values.",
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_static_sheets.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_static_sheets.py scripts/tests/test_build_static_sheets.py
git commit -m "Rebuild Data sheet as 8 per-month column-blocks; update Notes for multi-month workflow"
```

---

### Task 7: Rebuild `build_diff_lag.py` — 8-block `DiffLag`

**Files:**
- Modify: `scripts/build_diff_lag.py`
- Modify: `scripts/tests/test_build_diff_lag.py`

**Interfaces:**
- Consumes: `multi_month.{column_offset, DIFFLAG_BLOCK_WIDTH, DATA_BLOCK_WIDTH, GAP}`, `config_map.constant_map_for_month`
- Modifies: `write_diff_lag_sheet(wb, source_path, month: str, num_rows=3000) -> None`.

- [ ] **Step 1: Write the failing tests**

Replace `scripts/tests/test_build_diff_lag.py`:

```python
import openpyxl, pathlib
from build_diff_lag import write_diff_lag_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_diff_lag_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["DiffLag"]
    assert ws["A1"].value == "Row"
    assert ws["B1"].value == "Date"
    assert ws["B2"].value == '=IF(Data!B2="","",Data!A2)'
    assert ws["B11"].value == '=IF(Data!B11="","",Data!A11)'


def test_diff_lag_march_relocated_and_referencing_marchs_data_block():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["DiffLag"]
    # DiffLag's own offset for March: 1 * (8 + 2) = 10 -> B(2)+10=12='L'
    # Data's offset for March: 1 * (20 + 2) = 22 -> A(1)+22=23='W', B(2)+22=24='X'
    assert ws["L1"].value == "Date"
    assert ws["L2"].value == '=IF(Data!X2="","",Data!W2)'


def test_diff_lag_accumulates_across_months():
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    write_diff_lag_sheet(wb, SOURCE_PATH, "March", num_rows=5)
    ws = wb["DiffLag"]
    assert ws["B2"].value is not None   # January's block survives
    assert ws["L2"].value is not None   # March's block was added


def test_diff_lag_month_label_in_gap_after_block():
    from multi_month import column_offset, DIFFLAG_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_diff_lag_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    label_col = shift_col("A", column_offset("January", DIFFLAG_BLOCK_WIDTH, GAP) + DIFFLAG_BLOCK_WIDTH)
    assert wb["DiffLag"][f"{label_col}1"].value == "January"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_diff_lag.py -v`
Expected: FAIL.

- [ ] **Step 3: Rewrite `build_diff_lag.py`**

```python
from openpyxl.utils import get_column_letter
from sheet_cloner import fill_sheet_from_pattern
from config_map import constant_map_for_month
from multi_month import column_offset, DIFFLAG_BLOCK_WIDTH, DATA_BLOCK_WIDTH, GAP

HEADERS = ["Row", "Date", "BOHO", "D4", "dBOHO", "dD4", "BOHO(-1)", "D4(-1)"]


def write_diff_lag_sheet(wb, source_path: str, month: str, num_rows: int = 3000) -> None:
    if "DiffLag" not in wb.sheetnames:
        ws = wb.create_sheet("DiffLag")
    else:
        ws = wb["DiffLag"]

    self_offset = column_offset(month, DIFFLAG_BLOCK_WIDTH, GAP)
    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)

    for i, h in enumerate(HEADERS, start=1):
        ws[f"{get_column_letter(i + self_offset)}1"] = h

    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="DiffLag",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=constant_map_for_month(month),
        column_shift_map={None: (self_offset, None), "Data": (data_offset, "Data")},
    )

    ws[f"{get_column_letter(1 + self_offset + DIFFLAG_BLOCK_WIDTH)}1"] = month
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_diff_lag.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_diff_lag.py scripts/tests/test_build_diff_lag.py
git commit -m "Rebuild DiffLag as 8 per-month column-blocks"
```

---

### Task 8: Rebuild `build_monthly_beta.py` — 8-block `Monthly Beta`

**Files:**
- Modify: `scripts/build_monthly_beta.py`
- Modify: `scripts/tests/test_build_monthly_beta.py`

**Interfaces:**
- Consumes: `multi_month.{column_offset, MONTHLY_BETA_BLOCK_WIDTH, DATA_BLOCK_WIDTH, GAP, shift_col}`, `config_map.constant_map_for_month`
- Modifies: `write_monthly_beta_sheet(wb, source_path, month: str, num_rows=3000) -> None`. Recall from earlier inspection: nothing in the workbook reads FROM `Monthly Beta` — it only reads FROM `Data`, so its `column_shift_map` needs only `None` and `"Data"` keys, the same shape as Task 7's `DiffLag`.

- [ ] **Step 1: Write the failing tests**

Replace `scripts/tests/test_build_monthly_beta.py`:

```python
import openpyxl, pathlib
from build_monthly_beta import write_monthly_beta_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_monthly_beta_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    ws = wb["Monthly Beta"]
    assert ws["A1"].value == "Date"
    assert ws["H1"].value == "Beta_Locked_Monthly"
    assert ws["A2"].value == "=Data!A2"


def test_monthly_beta_march_relocated():
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "March", num_rows=5)
    ws = wb["Monthly Beta"]
    # Monthly Beta's own offset for March: 1 * (8 + 2) = 10 -> A(1)+10=11='K'
    # Data's offset for March: 22 -> A(1)+22=23='W'
    assert ws["K1"].value == "Date"
    assert ws["K2"].value == "=Data!W2"


def test_monthly_beta_month_label_in_gap_after_block():
    from multi_month import column_offset, MONTHLY_BETA_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_monthly_beta_sheet(wb, SOURCE_PATH, "January", num_rows=5)
    label_col = shift_col("A", column_offset("January", MONTHLY_BETA_BLOCK_WIDTH, GAP) + MONTHLY_BETA_BLOCK_WIDTH)
    assert wb["Monthly Beta"][f"{label_col}1"].value == "January"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_monthly_beta.py -v`
Expected: FAIL.

- [ ] **Step 3: Rewrite `build_monthly_beta.py`**

```python
from openpyxl.utils import get_column_letter
from sheet_cloner import fill_sheet_from_pattern
from config_map import constant_map_for_month
from multi_month import column_offset, MONTHLY_BETA_BLOCK_WIDTH, DATA_BLOCK_WIDTH, GAP

HEADERS = ["Date", "BOHO", "D4", "BOHO_Δ", "D4_Δ", "MonthEnd", "Beta_120d", "Beta_Locked_Monthly"]


def write_monthly_beta_sheet(wb, source_path: str, month: str, num_rows: int = 3000) -> None:
    if "Monthly Beta" not in wb.sheetnames:
        ws = wb.create_sheet("Monthly Beta")
    else:
        ws = wb["Monthly Beta"]

    self_offset = column_offset(month, MONTHLY_BETA_BLOCK_WIDTH, GAP)
    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)

    for i, h in enumerate(HEADERS, start=1):
        ws[f"{get_column_letter(i + self_offset)}1"] = h

    fill_sheet_from_pattern(
        target_ws=ws,
        source_path=source_path,
        source_sheet="Monthly Beta",
        pattern_row=2,
        target_start_row=2,
        num_rows=num_rows,
        max_col=8,
        constant_map=constant_map_for_month(month),
        column_shift_map={None: (self_offset, None), "Data": (data_offset, "Data")},
    )

    ws[f"{get_column_letter(1 + self_offset + MONTHLY_BETA_BLOCK_WIDTH)}1"] = month
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_monthly_beta.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_monthly_beta.py scripts/tests/test_build_monthly_beta.py
git commit -m "Rebuild Monthly Beta as 8 per-month column-blocks"
```

---

### Task 9: Rebuild `build_matrices_johansen.py` — combined `Matrices_Johansen_Results` tab

**Files:**
- Modify: `scripts/build_matrices_johansen.py`
- Modify: `scripts/tests/test_build_matrices_johansen.py`

**Interfaces:**
- Consumes: `multi_month.{mjr_offsets, column_offset, DIFFLAG_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP}`, `config_map.{constant_map_for_month, matrices_only_constant_map_for_month}`, `formula_tools.shift_formula_columns`
- Replaces `write_matrices_sheet`/`write_johansen_sheet`/`write_results_sheet` with a single `write_matrices_johansen_results_sheet(wb, source_path, month: str) -> None`.

This is the task where the sheet-merge rename matters: Johansen's `=Matrices!C50` and Results' `=Matrices!B2`/`=Johansen!B11`/`=Johansen!B12` all currently cross into what are, after this task, sub-blocks on the *same* physical sheet — those references get their qualifier **stripped** (written unqualified). Matrices' own `=Johansen_Monthly!$K$3` stays genuinely cross-tab, so it stays qualified, just column-shifted to that month's `Johansen_Monthly` block.

- [ ] **Step 1: Write the failing tests**

Replace `scripts/tests/test_build_matrices_johansen.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v`
Expected: FAIL — `write_matrices_johansen_results_sheet` doesn't exist yet.

- [ ] **Step 3: Rewrite `build_matrices_johansen.py`**

```python
import openpyxl
from openpyxl.utils import coordinate_from_string
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
    ws[f"{b_col}3"] = matrices_map["$B$3"]
    ws[f"{b_col}2"] = f"=Johansen_Monthly!${k_col_jm}$3"

    # Plain text label in the gap right after this month's mega-block.
    mega_offset = column_offset(month, MJR_BLOCK_WIDTH, MJR_MONTH_GAP)
    ws[f"{shift_col('A', mega_offset + MJR_BLOCK_WIDTH)}1"] = month
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_matrices_johansen.py -v`
Expected: PASS. If `test_january_johansen_sub_block_references_matrices_unqualified` fails on the exact cell (`K3` vs. wherever Johansen's real `B11` lands after the offset), re-derive the expected coordinate from `mjr_offsets("January")["Johansen"]` directly in the test rather than guessing — the source values (`Johansen!B11`, `Johansen!B12`) were confirmed by the read-only live-workbook inspection during planning; only the arithmetic translating them to final coordinates is at risk of an off-by-one, and the test should compute that arithmetic from `multi_month`'s own functions, not repeat it by hand.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_matrices_johansen.py scripts/tests/test_build_matrices_johansen.py
git commit -m "Combine Matrices, Johansen, and Results into one 8-block Matrices_Johansen_Results tab"
```

---

### Task 10: Rebuild `build_johansen_monthly.py` — 8-block `Johansen_Monthly`

**Files:**
- Modify: `scripts/build_johansen_monthly.py`
- Modify: `scripts/tests/test_build_johansen_monthly.py`

**Interfaces:**
- Consumes: `multi_month.{column_offset, mjr_offsets, DIFFLAG_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col}`, `config_map.constant_map_for_month`, `formula_tools.shift_formula_columns`, `build_matrices_johansen.COMBINED_SHEET_NAME`
- Modifies: `write_johansen_monthly_sheet(wb, source_path, month: str, num_rows=3000) -> None`.

Recall the exact layout confirmed by live-source inspection: column A is a plain `=DiffLag!B{row}` mirror; B is the month-boundary detector; C2/D2 read `=Johansen!D18`/`=Johansen!D19` (now `Matrices_Johansen_Results`, at that month's Johansen sub-block offset); F/G/H/I are the manually-maintained log headers only; K2/K3 are the "Active End Row" label/self-updating lookup over this month's own F column.

- [ ] **Step 1: Write the failing tests**

Replace `scripts/tests/test_build_johansen_monthly.py` (keep any tests that still make sense unchanged in spirit; the concrete list below is what must be present after the rewrite — merge with, rather than blindly discard, whatever the existing file's other assertions cover that aren't superseded):

```python
import openpyxl, pathlib
from build_johansen_monthly import write_johansen_monthly_sheet

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_january_matches_todays_single_month_layout():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from multi_month import mjr_offsets, shift_col
    johansen_col = shift_col("D", mjr_offsets("January")["Johansen"])  # D(4)+10=14='N'
    assert ws["A2"].value == "=DiffLag!B2"
    assert ws["C2"].value == f"=Matrices_Johansen_Results!{johansen_col}18"
    assert ws["D2"].value == f"=Matrices_Johansen_Results!{johansen_col}19"
    assert ws["F2"].value == "EndRow"
    assert ws["K2"].value == "Active End Row"
    assert ws["K3"].value == '=LOOKUP(2,1/(F3:F11<>""),F3:F11)'


def test_march_relocated_and_referencing_marchs_diff_lag_block():
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Johansen_Monthly"]
    from multi_month import column_offset, mjr_offsets, JOHANSEN_MONTHLY_BLOCK_WIDTH, DIFFLAG_BLOCK_WIDTH, GAP, shift_col
    self_offset = column_offset("March", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)  # 1*(11+2)=13
    diff_lag_offset = column_offset("March", DIFFLAG_BLOCK_WIDTH, GAP)      # 1*(8+2)=10
    a_col = shift_col("A", self_offset)
    diff_lag_col = shift_col("B", diff_lag_offset)
    assert ws[f"{a_col}2"].value == f"=DiffLag!{diff_lag_col}2"

    johansen_col = shift_col("D", mjr_offsets("March")["Johansen"])
    c_col = shift_col("C", self_offset)
    assert ws[f"{c_col}2"].value == f"=Matrices_Johansen_Results!{johansen_col}18"

    k_col = shift_col("K", self_offset)
    assert ws[f"{k_col}2"].value == "Active End Row"
```

Add one more test, following the same label-in-gap pattern as Tasks 6-9:

```python
def test_january_month_label_in_gap_after_block():
    from multi_month import column_offset, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
    wb = openpyxl.Workbook()
    write_johansen_monthly_sheet(wb, SOURCE_PATH, "January", num_rows=10)
    label_col = shift_col("A", column_offset("January", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP) + JOHANSEN_MONTHLY_BLOCK_WIDTH)
    assert wb["Johansen_Monthly"][f"{label_col}1"].value == "January"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_johansen_monthly.py -v`
Expected: FAIL.

- [ ] **Step 3: Rewrite `build_johansen_monthly.py`**

```python
import openpyxl
from openpyxl.utils import get_column_letter
from formula_tools import rewrite_constants, shift_formula_columns
from config_map import constant_map_for_month
from multi_month import column_offset, mjr_offsets, DIFFLAG_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
from build_matrices_johansen import COMBINED_SHEET_NAME

HEADERS = ["Date", "EndRow", "BOHO_Weight", "D4_Weight"]


def write_johansen_monthly_sheet(wb, source_path: str, month: str, num_rows: int = 3000) -> None:
    if "Johansen_Monthly" not in wb.sheetnames:
        ws = wb.create_sheet("Johansen_Monthly")
    else:
        ws = wb["Johansen_Monthly"]

    self_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
    diff_lag_offset = column_offset(month, DIFFLAG_BLOCK_WIDTH, GAP)
    johansen_offset = mjr_offsets(month)["Johansen"]
    month_constant_map = constant_map_for_month(month)

    def col(letter):
        return shift_col(letter, self_offset)

    for i, h in enumerate(HEADERS, start=1):
        ws[f"{get_column_letter(i + self_offset)}1"] = h

    # Column A: plain per-row mirror of this month's DiffLag Date column.
    for i in range(num_rows):
        row = 2 + i
        ws[f"{col('A')}{row}"] = f"=DiffLag!{shift_col('B', diff_lag_offset)}{row}"

    # Column B: month-boundary detector, referencing this month's own column A
    # and this month's DiffLag block.
    for i in range(num_rows):
        row = 2 + i
        diff_lag_next = shift_col("B", diff_lag_offset)
        a_col, b_col = col("A"), col("B")
        ws[f"{b_col}{row}"] = (
            f'=IF(OR({a_col}{row}="",{a_col}{row + 1}=""),"",IF(MONTH({a_col}{row})<>MONTH({a_col}{row + 1}), '
            f'ROW({a_col}{row})+ROW(DiffLag!{diff_lag_next}{row + 1})-ROW({a_col}{row})-1, ""))'
        )

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Johansen_Monthly"]
        c2 = shift_formula_columns(
            rewrite_constants(src["C2"].value, month_constant_map),
            {"Johansen": (johansen_offset, COMBINED_SHEET_NAME)},
        )
        d2 = shift_formula_columns(
            rewrite_constants(src["D2"].value, month_constant_map),
            {"Johansen": (johansen_offset, COMBINED_SHEET_NAME)},
        )
        ws[f"{col('C')}2"] = c2
        ws[f"{col('D')}2"] = d2
    finally:
        source_wb.close()

    ws[f"{col('F')}2"] = "EndRow"
    ws[f"{col('G')}2"] = "Date"
    ws[f"{col('H')}2"] = "BOHO_Weight"
    ws[f"{col('I')}2"] = "D4_Weight"

    ws[f"{col('K')}2"] = "Active End Row"
    last_log_row = num_rows + 1
    f_col = col("F")
    ws[f"{col('K')}3"] = f'=LOOKUP(2,1/({f_col}3:{f_col}{last_log_row}<>""),{f_col}3:{f_col}{last_log_row})'

    ws[f"{shift_col('A', self_offset + JOHANSEN_MONTHLY_BLOCK_WIDTH)}1"] = month
```

Note: `src["C2"].value` (`"=Johansen!D18"`) is a plain string in the live source, not an `ArrayFormula` — confirmed by the read-only inspection during planning — so no `unwrap_scalar_array_formula` call is needed here, matching today's un-migrated code.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_build_johansen_monthly.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_johansen_monthly.py scripts/tests/test_build_johansen_monthly.py
git commit -m "Rebuild Johansen_Monthly as 8 per-month column-blocks"
```

---

### Task 11: Modify `build_position_model.py` — 8 separate tabs, cross-sheet relocation, old `config_map` cleanup

**Files:**
- Modify: `scripts/build_position_model.py:202-334` (the `write_position_model_sheet` entry point only — every `BOOT_END`/`PARAM_BLOCK_COLS`/`ACCUM_SELF_REF_COLS`/`_write_generalized_trailing_window_accum_cols`/`_apply_targeted_fix_round_2_overrides`/`_wrap_accumulator_self_references`/`_write_dynamic_year_breakdown` line stays byte-for-byte unchanged — none of that logic depends on which month this is, because Position Model tabs are never column-shifted)
- Modify: `scripts/config_map.py` (remove `CONFIG_LAYOUT`, `CONSTANT_MAP`, `MATRICES_ONLY_CONSTANT_MAP` — the last consumers are migrated in this task)
- Modify: `scripts/build_config_sheet.py` (already migrated in Task 5 — confirm it doesn't still import the old names; it shouldn't)
- Modify: `scripts/tests/test_build_position_model.py`
- Modify: `scripts/tests/test_formula_tools.py` (its `rewrite_constants` tests currently import the global `CONSTANT_MAP` — switch them to a small locally-defined dict, since those tests exist to verify `rewrite_constants`'s generic mechanism, not `config_map.py`'s specific content)
- Modify: `scripts/tests/test_sheet_cloner.py` (same reason)

**Interfaces:**
- Modifies: `write_position_model_sheet(wb, source_path, month: str, num_rows=3000, max_col=None) -> None`. Creates a sheet named `f"Position Model - {month}"` (always delete-then-create, matching today's behavior, since each month's tab is independent with no cross-call accumulation).
- Produces (new, appended at the very end of the function, after `_write_dynamic_year_breakdown`): `_relocate_cross_sheet_references(ws, last_target_row: int, max_col: int, sheet_map: dict) -> None` — a single blanket pass over every populated cell in the sheet, applying `shift_formula_columns` in place. Only `Data` and `Johansen_Monthly` are ever in `sheet_map` here (confirmed by read-only inspection: Position Model references no other sheet directly), and there is deliberately no `None` key, so Position Model's own internal self-references are never touched.

- [ ] **Step 1: Write the failing tests**

Append to `scripts/tests/test_build_position_model.py` (leave every existing test in the file untouched — they all still apply to `month="January"`, since January's offsets are 0, but they need the new required `month` argument added to their `write_position_model_sheet(...)` calls, and `wb["Position Model"]` lookups changed to `wb["Position Model - January"]`; do this as a mechanical find-and-replace across the file before adding the tests below):

```python
def test_position_model_march_sheet_is_named_for_its_month():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    assert "Position Model - March" in wb.sheetnames
    assert "Position Model" not in wb.sheetnames


def test_position_model_march_data_reference_relocated():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Position Model - March"]
    # A3 = "=Data!A3" in the single-month template; Data's offset for March is 22.
    assert ws["A3"].value == "=Data!W3"


def test_position_model_march_johansen_monthly_reference_relocated():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Position Model - March"]
    # I3 references Johansen_Monthly!$G:$G and $H:$H in the single-month template.
    # Johansen_Monthly's offset for March is 13 -> G(7)+13=20='T', H(8)+13=21='U'.
    formula = ws["I3"].value
    assert "Johansen_Monthly!$T:$T" in formula
    assert "Johansen_Monthly!$U:$U" in formula
    assert "Johansen_Monthly!$G:$G" not in formula


def test_position_model_march_internal_self_references_not_shifted():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Position Model - March"]
    # Y is a self-referencing accumulator column (Y{r}=IFERROR(Y{r-1},0)+...).
    # It must stay column Y -- Position Model tabs are never column-shifted.
    assert ws["Y4"].value is not None
    assert "IFERROR(Y3,0)" in ws["Y4"].value


def test_position_model_march_config_reference_uses_marchs_column():
    wb = openpyxl.Workbook()
    write_position_model_sheet(wb, SOURCE_PATH, "March", num_rows=10)
    ws = wb["Position Model - March"]
    # M4 (or wherever ZExitThreshold is consumed) must reference March's own
    # Config column ("C"), not January's ("B") and not a defined name.
    m_formula = ws["M4"].value
    assert "Config!C17" in m_formula
    assert "ZExitThreshold" not in m_formula
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_position_model.py -v`
Expected: FAIL — `write_position_model_sheet` doesn't take a `month` parameter yet.

- [ ] **Step 3: Modify `build_position_model.py`**

Replace the import line at the top of the file:

```python
from config_map import constant_map_for_month
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP
```

(removing `from config_map import CONSTANT_MAP`).

Replace the `write_position_model_sheet` signature and every `CONSTANT_MAP` reference within it (lines 202–329) with the per-month equivalent, and add the new sheet name plus the final relocation call:

```python
def write_position_model_sheet(wb, source_path: str, month: str, num_rows: int = 3000, max_col: int = None) -> None:
    sheet_name = f"Position Model - {month}"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    constant_map = constant_map_for_month(month)

    source_wb = openpyxl.load_workbook(source_path, data_only=False)
    try:
        src = source_wb["Position Model"]
        if max_col is None:
            max_col = src.max_column

        # ... every other line inside this try/finally block is UNCHANGED —
        # only replace every occurrence of the bare name `CONSTANT_MAP` with
        # `constant_map` (the new per-month dict computed above). There are
        # five such occurrences: (1) the baseline fill_sheet_from_pattern
        # call's constant_map= argument, (2) the row-2 boundary-row
        # rewrite_constants call, (3) the BOOTSTRAP-column literal-copy
        # loop's rewrite_constants call, (4) the BOOTSTRAP steady-state
        # shift's rewrite_constants call, and (5) the PARAM_BLOCK_COLS
        # literal-copy loop's rewrite_constants call. Do not change anything
        # else in this block.
    finally:
        source_wb.close()

    _write_generalized_trailing_window_accum_cols(ws, last_target_row, max_col)
    _apply_targeted_fix_round_2_overrides(ws, last_target_row, max_col)
    _wrap_accumulator_self_references(ws, last_target_row)
    _write_dynamic_year_breakdown(ws, last_target_row)

    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
    jm_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
    _relocate_cross_sheet_references(
        ws, last_target_row, max_col,
        sheet_map={"Data": (data_offset, "Data"), "Johansen_Monthly": (jm_offset, "Johansen_Monthly")},
    )
```

Add the new function at the end of the file:

```python
def _relocate_cross_sheet_references(ws, last_target_row: int, max_col: int, sheet_map: dict) -> None:
    """Position Model tabs are never column-shifted themselves (each is a
    full standalone sheet starting at column A, exactly like today's
    single-month template) -- only their CROSS-sheet references need
    relocating, to whichever column-block this month's Data/Johansen_Monthly
    live at. Read-only inspection of the live source workbook confirmed
    Position Model references no other sheet directly (not DiffLag, not
    Matrices/Johansen/Results, not Monthly Beta -- nothing references Monthly
    Beta from anywhere in the workbook), so sheet_map only ever needs "Data"
    and "Johansen_Monthly" keys, and deliberately has no None key: Position
    Model's own internal self-references (the accumulator columns, the
    dynamic year breakdown, etc.) must never be touched by this pass.

    A blanket full-sheet scan, not a targeted one, on purpose -- consistent
    with this project's established preference for an opt-out validation
    surface over an opt-in one (see docs/superpowers/HANDOFF.md's "hard-won
    gotchas"): safer to check every cell than to guess which columns might
    contain a cross-sheet reference.
    """
    for row in ws.iter_rows(min_row=1, max_row=last_target_row, max_col=max_col):
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                cell.value = shift_formula_columns(value, sheet_map)
```

Add `shift_formula_columns` to the existing `formula_tools` import line at the top of the file.

Finally, in `scripts/config_map.py`, delete `CONFIG_LAYOUT`, `CONSTANT_MAP`, and `MATRICES_ONLY_CONSTANT_MAP` (and their header comments) entirely — every consumer has now migrated to `CONFIG_FIELDS`/`constant_map_for_month`/`matrices_only_constant_map_for_month`.

Update `scripts/tests/test_formula_tools.py`'s `rewrite_constants` tests to stop importing `CONSTANT_MAP`:

```python
_TEST_CONSTANT_MAP = {"$DV$6": "ZExitThreshold", "$DK$5": "HOGalPerContract"}

def test_rewrite_single_constant():
    from formula_tools import rewrite_constants
    formula = "=IF($L3<=$DV$6,0,IF(AND($M2=0,$L3>=M$1),$DX$3,M2))"
    result = rewrite_constants(formula, _TEST_CONSTANT_MAP)
    assert "$DV$6" not in result
    assert "ZExitThreshold" in result
    assert "$DX$3" in result  # not in the map, left untouched

def test_rewrite_leaves_unmapped_addresses_untouched():
    from formula_tools import rewrite_constants
    formula = "=Data!A2"
    assert rewrite_constants(formula, _TEST_CONSTANT_MAP) == "=Data!A2"

def test_rewrite_multiple_constants_in_one_formula():
    from formula_tools import rewrite_constants
    formula = "=S2*($Z2*$DK$5-$AA2*$DK$5)"
    result = rewrite_constants(formula, _TEST_CONSTANT_MAP)
    assert result == "=S2*($Z2*HOGalPerContract-$AA2*HOGalPerContract)"
```

Remove the file's `from config_map import CONSTANT_MAP` import line, and remove the (now-redundant) `test_extract_position_model_row2`/`test_extract_diff_lag_row2` tests' dependency on `CONSTANT_MAP` if any (they don't call `rewrite_constants`, so check first — they likely only used `extract_row_formulas`, unaffected).

Apply the same `_TEST_CONSTANT_MAP` substitution to `scripts/tests/test_sheet_cloner.py`'s two `CONSTANT_MAP`-importing tests.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/ -v -k "not validate_against_dec"`
Expected: PASS across the board — this is the point where every builder has been migrated off the old `config_map.py` names, so the full non-COM test suite should be green.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_position_model.py scripts/config_map.py scripts/tests/test_build_position_model.py scripts/tests/test_formula_tools.py scripts/tests/test_sheet_cloner.py
git commit -m "Position Model: 8 separate tabs with relocated cross-sheet references; retire single-column Config API"
```

---

### Task 12: Rebuild `build_template.py` orchestrator

**Files:**
- Modify: `scripts/build_template.py`
- Modify: `scripts/tests/test_build_template.py`

**Interfaces:**
- Consumes every builder from Tasks 5–11, plus `multi_month.MONTHS`.
- Modifies: `build_template(source_path, output_path, num_rows=3000) -> None` — same signature, now loops over `MONTHS` for every per-month sheet.

- [ ] **Step 1: Write the failing tests**

Replace `scripts/tests/test_build_template.py`'s sheet-order assertion and add new ones (keep the file's other tests — array-formula scan, leftover-Dec-constant scan — updating them only where they hardcode the old sheet name `"Position Model"`):

```python
import openpyxl, pathlib, tempfile, os
from build_template import build_template
from multi_month import MONTHS

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_build_template_creates_expected_sheets_in_order():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        assert wb.sheetnames == [
            "Notes", "Config", "Data", "DiffLag", "Matrices_Johansen_Results",
            "Johansen_Monthly", "Monthly Beta",
        ] + [f"Position Model - {m}" for m in MONTHS]


def test_build_template_data_sheet_has_no_price_values_for_every_month():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col
        for month in MONTHS:
            offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
            assert wb["Data"][f"{shift_col('B', offset)}2"].value is None
            assert wb["Data"][f"{shift_col('C', offset)}2"].value is None


def test_build_template_no_leftover_dec_absolute_constants_in_any_month():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "test_template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=10)
        wb = openpyxl.load_workbook(out, data_only=False)
        for month in MONTHS:
            ws = wb[f"Position Model - {month}"]
            for row in ws.iter_rows(min_row=2, max_row=11):
                for cell in row:
                    if isinstance(cell.value, str):
                        assert "$DV$6" not in cell.value
                        assert "$DQ$7" not in cell.value
```

Keep the existing `test_build_template_has_no_array_formula_cells_anywhere` test as-is (it already loops over `wb.sheetnames`, so it needs no change).

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_build_template.py -v`
Expected: FAIL — orchestrator still builds the single-month layout.

- [ ] **Step 3: Rewrite `build_template.py`**

```python
import openpyxl

from build_config_sheet import write_config_sheet
from build_static_sheets import write_notes_sheet, write_data_sheet
from build_diff_lag import write_diff_lag_sheet
from build_matrices_johansen import write_matrices_johansen_results_sheet, COMBINED_SHEET_NAME
from build_johansen_monthly import write_johansen_monthly_sheet
from build_monthly_beta import write_monthly_beta_sheet
from build_position_model import write_position_model_sheet
from multi_month import MONTHS

SHEET_ORDER = [
    "Notes", "Config", "Data", "DiffLag", COMBINED_SHEET_NAME,
    "Johansen_Monthly", "Monthly Beta",
] + [f"Position Model - {m}" for m in MONTHS]


def build_template(source_path: str, output_path: str, num_rows: int = 3000) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    write_notes_sheet(wb)
    write_config_sheet(wb)
    for month in MONTHS:
        write_data_sheet(wb, month)
    for month in MONTHS:
        write_diff_lag_sheet(wb, source_path, month, num_rows=num_rows)
    for month in MONTHS:
        write_matrices_johansen_results_sheet(wb, source_path, month)
    for month in MONTHS:
        write_johansen_monthly_sheet(wb, source_path, month, num_rows=num_rows)
    for month in MONTHS:
        write_monthly_beta_sheet(wb, source_path, month, num_rows=num_rows)
    for month in MONTHS:
        write_position_model_sheet(wb, source_path, month, num_rows=num_rows)
    wb.remove(default_sheet)
    wb._sheets.sort(key=lambda ws: SHEET_ORDER.index(ws.title))
    wb.save(output_path)


if __name__ == "__main__":
    import pathlib
    source = (pathlib.Path(__file__).parent / "DEC_SOURCE_PATH.txt").read_text().strip()
    out = pathlib.Path(__file__).parent.parent / "Johansen_MultiMonth_Template_Engine.xlsx"
    build_template(source, str(out))
    print(f"Template written to {out}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/ -v -k "not validate_against_dec"`
Expected: PASS across the entire non-COM suite.

- [ ] **Step 5: Commit**

```bash
git add scripts/build_template.py scripts/tests/test_build_template.py
git commit -m "Rebuild build_template.py orchestrator for the 8-month consolidated workbook"
```

---

### Task 13: Update `validate_against_dec.py` for December's block within the multi-month workbook

**Files:**
- Modify: `scripts/validate_against_dec.py`
- Modify: `scripts/tests/test_validate_against_dec.py`

**Interfaces:**
- Consumes: `multi_month.{column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP}`, `config_map.{CONFIG_FIELDS, CONFIG_MONTH_COLUMNS}`
- Modifies: `_populate_throwaway_copy` (write Dec's real `Data`/`Johansen_Monthly` values into December's block, and Dec's real values into December's `Config` column, instead of column A / column B), `run_validation` (compare `template_wb["Position Model - December"]` against `dec_wb["Position Model"]`, everything else — `compare_columns_for`, `param_block_value_parity_mismatches`, `EXCLUDED_COMPARE_COLUMNS` — unchanged, since December's Position Model tab is never column-shifted and uses the exact same column letters A through the sheet's `max_col` as today).

- [ ] **Step 1: Write the failing test**

Add to `scripts/tests/test_validate_against_dec.py` (append; do not remove the existing accepted-exception tests, which still apply once the target sheet name changes):

```python
def test_populate_throwaway_copy_writes_into_decembers_block_not_column_a(tmp_path):
    import openpyxl
    from build_template import build_template
    from validate_against_dec import _populate_throwaway_copy
    from multi_month import column_offset, DATA_BLOCK_WIDTH, GAP, shift_col

    template_path = str(tmp_path / "template.xlsx")
    build_template(SOURCE_PATH, template_path, num_rows=20)
    throwaway_path = str(tmp_path / "throwaway.xlsx")

    _populate_throwaway_copy(template_path, SOURCE_PATH, throwaway_path)

    wb = openpyxl.load_workbook(throwaway_path, data_only=False)
    december_offset = column_offset("December", DATA_BLOCK_WIDTH, GAP)
    date_col = shift_col("A", december_offset)
    boho_col = shift_col("B", december_offset)
    assert wb["Data"][f"{date_col}2"].value is not None
    assert wb["Data"][f"{boho_col}2"].value is not None
    # January's block (offset 0) must NOT have been touched by Dec's data.
    assert wb["Data"]["A2"].value is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd scripts && py -m pytest tests/test_validate_against_dec.py -v -k decembers_block`
Expected: FAIL — `_populate_throwaway_copy` still writes at column A.

- [ ] **Step 3: Modify `validate_against_dec.py`**

Replace the `DEC_LIVE_VALUES` dict's consumption and the `_populate_throwaway_copy`/`run_validation` functions:

```python
from config_map import CONFIG_FIELDS
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col

# ... DEC_LIVE_VALUES dict unchanged ...

def _populate_throwaway_copy(template_path: str, dec_source_path: str, throwaway_path: str) -> None:
    shutil.copy(template_path, throwaway_path)
    wb = openpyxl.load_workbook(throwaway_path, data_only=False)

    config_ws = wb["Config"]
    row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
    for name, value in DEC_LIVE_VALUES.items():
        config_ws[f"I{row_by_field[name]}"] = value  # I = December's Config column

    data_offset = column_offset("December", DATA_BLOCK_WIDTH, GAP)
    jm_offset = column_offset("December", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)

    dec_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    dec_data = dec_wb["Data"]
    data_ws = wb["Data"]
    for row in dec_data.iter_rows(min_row=2, max_row=dec_data.max_row, max_col=20):
        for cell in row:
            if cell.value is not None:
                from openpyxl.utils import coordinate_from_string
                col_letter, _ = coordinate_from_string(cell.coordinate)
                target_col = shift_col(col_letter, data_offset)
                data_ws[f"{target_col}{cell.row}"] = cell.value

    dec_jm = dec_wb["Johansen_Monthly"]
    jm_ws = wb["Johansen_Monthly"]
    for row in dec_jm.iter_rows(min_row=2, max_row=dec_jm.max_row, min_col=6, max_col=9):  # F:I
        for cell in row:
            if cell.value is not None:
                from openpyxl.utils import coordinate_from_string
                col_letter, _ = coordinate_from_string(cell.coordinate)
                target_col = shift_col(col_letter, jm_offset)
                jm_ws[f"{target_col}{cell.row}"] = cell.value

    wb.save(throwaway_path)
```

(Move both inline `from openpyxl.utils import coordinate_from_string` lines to the top of the file with the other imports.)

In `run_validation`, change:

```python
    template_pm = template_wb["Position Model"]
```

to:

```python
    template_pm = template_wb["Position Model - December"]
```

Everything else in `run_validation` — `compare_columns_for`, the error-parity check, `param_block_value_parity_mismatches` — is unchanged, since `template_pm` is still a full A-through-`max_col` sheet, just under a new name.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_validate_against_dec.py -v -k decembers_block`
Expected: PASS.

- [ ] **Step 5: Run the full validation end-to-end (not just the unit test)**

Run: `cd scripts && py -m pytest tests/ -v` (the full suite, including the real Excel-COM recalculation tests this time — expect ~10–15 minutes)
Expected: PASS, including `run_validation`'s real comparison of `Position Model - December` against Dec's real recalculated output. This is the actual proof this refactor didn't silently change December's numbers — trust it over every other test in this plan. If the accepted-exception count (`== 414` or whatever it currently is) no longer matches, investigate before adjusting the assertion — per this project's established discipline (see `docs/superpowers/HANDOFF.md`'s gotchas), a changed count must be individually root-caused, never just widened.

- [ ] **Step 6: Commit**

```bash
git add scripts/validate_against_dec.py scripts/tests/test_validate_against_dec.py
git commit -m "Point validate_against_dec.py at December's block within the multi-month workbook"
```

---

### Task 14: One-time migration script for December's and January's existing real data

**Files:**
- Create: `scripts/migrate_existing_months.py`
- Create: `scripts/tests/test_migrate_existing_months.py`

**Interfaces:**
- Produces: `migrate_month_data(wb, source_wb, month: str, config_values: dict[str, object] | None = None) -> None` — copies `source_wb`'s `Data` (A:T) and `Johansen_Monthly` (F:I) values into `wb`'s consolidated tabs at `month`'s block offset, and (if `config_values` given) writes them into `wb`'s `Config` column for `month`. `run_migration(template_path, dec_source_path, january_source_path, output_path) -> None` — the end-to-end script: build (or load) the multi-month template, migrate December's real data (from the live source workbook, read-only) and January's real data (from `Johansen_January_BOHO_D4_Strategy.xlsx`, read-only, including reading that file's own already-filled-in `Config` sheet for January's real calibration values), and save as the new live book.

- [ ] **Step 1: Write the failing tests**

Create `scripts/tests/test_migrate_existing_months.py`:

```python
import openpyxl, pathlib, tempfile, os
from build_template import build_template
from migrate_existing_months import migrate_month_data
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col

SOURCE_PATH = (pathlib.Path(__file__).parent.parent / "DEC_SOURCE_PATH.txt").read_text().strip()


def test_migrate_month_data_writes_data_and_log_into_that_months_block():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December")

        data_offset = column_offset("December", DATA_BLOCK_WIDTH, GAP)
        assert wb["Data"][f"{shift_col('B', data_offset)}2"].value is not None

        jm_offset = column_offset("December", JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)
        assert wb["Johansen_Monthly"][f"{shift_col('F', jm_offset)}3"].value is not None
        source_wb.close()


def test_migrate_month_data_writes_config_values_when_given():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December", config_values={"ZThreshold_Tier1": 2.25})

        assert wb["Config"]["I14"].value == 2.25  # December's Config column is I
        source_wb.close()


def test_migrate_month_data_does_not_touch_other_months_blocks():
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "template.xlsx")
        build_template(SOURCE_PATH, out, num_rows=20)
        wb = openpyxl.load_workbook(out, data_only=False)
        source_wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)

        migrate_month_data(wb, source_wb, "December")

        assert wb["Data"]["B2"].value is None  # January's block, offset 0, untouched
        source_wb.close()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd scripts && py -m pytest tests/test_migrate_existing_months.py -v`
Expected: FAIL — module doesn't exist.

- [ ] **Step 3: Implement `scripts/migrate_existing_months.py`**

```python
import pathlib
import openpyxl
from openpyxl.utils import coordinate_from_string
from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS
from multi_month import column_offset, DATA_BLOCK_WIDTH, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP, shift_col
from build_template import build_template


def migrate_month_data(wb, source_wb, month: str, config_values: dict | None = None) -> None:
    data_offset = column_offset(month, DATA_BLOCK_WIDTH, GAP)
    jm_offset = column_offset(month, JOHANSEN_MONTHLY_BLOCK_WIDTH, GAP)

    data_ws = wb["Data"]
    source_data = source_wb["Data"]
    for row in source_data.iter_rows(min_row=2, max_row=source_data.max_row, max_col=20):
        for cell in row:
            if cell.value is not None:
                col_letter, _ = coordinate_from_string(cell.coordinate)
                data_ws[f"{shift_col(col_letter, data_offset)}{cell.row}"] = cell.value

    jm_ws = wb["Johansen_Monthly"]
    source_jm = source_wb["Johansen_Monthly"]
    for row in source_jm.iter_rows(min_row=2, max_row=source_jm.max_row, min_col=6, max_col=9):  # F:I
        for cell in row:
            if cell.value is not None:
                col_letter, _ = coordinate_from_string(cell.coordinate)
                jm_ws[f"{shift_col(col_letter, jm_offset)}{cell.row}"] = cell.value

    if config_values:
        config_ws = wb["Config"]
        row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
        month_col = CONFIG_MONTH_COLUMNS[month]
        for field, value in config_values.items():
            config_ws[f"{month_col}{row_by_field[field]}"] = value


def _read_config_values(config_path: str, month_col_in_source: str) -> dict:
    """Read a previously-activated single-month template's own Config
    sheet (column B, the only column it has) and return {field: value} for
    every field in CONFIG_FIELDS, so that month's already-calibrated real
    values can be carried into the multi-month workbook's Config column for
    that same month, rather than reverting to the generic Dec-derived
    defaults build_config_sheet.py starts every month at."""
    wb = openpyxl.load_workbook(config_path, data_only=True)
    ws = wb["Config"]
    row_by_field = {name: row for name, row, _default in CONFIG_FIELDS}
    values = {field: ws[f"{month_col_in_source}{row}"].value for field, row in row_by_field.items()}
    wb.close()
    return values


def run_migration(template_path: str, dec_source_path: str, january_source_path: str, output_path: str) -> None:
    wb = openpyxl.load_workbook(template_path, data_only=False)

    dec_source_wb = openpyxl.load_workbook(dec_source_path, data_only=True)
    try:
        migrate_month_data(wb, dec_source_wb, "December")
    finally:
        dec_source_wb.close()

    january_source_wb = openpyxl.load_workbook(january_source_path, data_only=True)
    try:
        january_config = _read_config_values(january_source_path, "B")
        migrate_month_data(wb, january_source_wb, "January", config_values=january_config)
    finally:
        january_source_wb.close()

    wb.save(output_path)


if __name__ == "__main__":
    scripts_dir = pathlib.Path(__file__).parent
    root = scripts_dir.parent
    template_path = str(root / "Johansen_MultiMonth_Template_Engine.xlsx")
    dec_source_path = (scripts_dir / "DEC_SOURCE_PATH.txt").read_text().strip()
    january_source_path = str(root / "Johansen_January_BOHO_D4_Strategy.xlsx")
    output_path = str(root / "Johansen_BOHO_D4_Strategy_AllMonths.xlsx")
    run_migration(template_path, dec_source_path, january_source_path, output_path)
    print(f"Migrated workbook written to {output_path}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd scripts && py -m pytest tests/test_migrate_existing_months.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add scripts/migrate_existing_months.py scripts/tests/test_migrate_existing_months.py
git commit -m "Add one-time migration script for December's and January's existing real data"
```

---

### Task 15: Update operational docs for the multi-month workflow

**Files:**
- Modify: `docs/ACTIVATING_A_NEW_MONTH.md`
- Modify: `docs/superpowers/HANDOFF.md`

**Interfaces:** None (documentation only).

- [ ] **Step 1: Rewrite `docs/ACTIVATING_A_NEW_MONTH.md`**

Replace its "Save As" step and Config/Data/Johansen_Monthly sections to describe the new workflow: no more `Save As`/new file — activating a previously-empty month means filling in that month's `Config` column (`CONFIG_MONTH_COLUMNS` in `scripts/config_map.py` is the authoritative column-letter reference) and pasting that month's price history into that month's `Data` block (offset computed by `multi_month.column_offset("<Month>", DATA_BLOCK_WIDTH, GAP)`). Keep the recurring-manual-step section (Johansen_Monthly F:I-equivalent log, per month, still manual), updating it to note it's now a per-month block on one shared `Johansen_Monthly` tab rather than a separate file. Update the "10 sheets, in order" section to the new 15-tab list (7 shared/consolidated tabs + 8 `Position Model - <Month>` tabs). Keep the "known documentation bug" callout about the manual log's F/G column order — unaffected by this refactor.

- [ ] **Step 2: Update `docs/superpowers/HANDOFF.md`**

Add a pointer at the top to this plan and to `docs/superpowers/specs/2026-08-06-johansen-multi-month-consolidation-design.md`, noting the workbook is now multi-month-consolidated and that the "Known limitations" punch list (hardcoded-row defect, `#N/A` summary cells, header labels, validation blind spots) now applies **independently to each of the 8 `Position Model - <Month>` tabs**, not to one shared sheet — none of the four items were fixed by this plan; they're still open, just multiplied across 8 tabs instead of hidden across 8 files.

- [ ] **Step 3: Commit**

```bash
git add docs/ACTIVATING_A_NEW_MONTH.md docs/superpowers/HANDOFF.md
git commit -m "Update operational docs for the multi-month consolidated workbook"
```

---

## Definition of Done

- `cd scripts && py -m pytest tests/ -v` passes in full, including the real Excel-COM `validate_against_dec.py` recalculation test, with December's block inside the new multi-month workbook producing output that matches the live December workbook to the same accepted-exception count as today's single-month template (re-derive and re-justify that count if it changes — never just widen it).
- `py build_template.py` (from `scripts/`) produces `Johansen_MultiMonth_Template_Engine.xlsx` with the 15-tab structure: `Notes`, `Config`, `Data`, `DiffLag`, `Matrices_Johansen_Results`, `Johansen_Monthly`, `Monthly Beta`, and 8 `Position Model - <Month>` tabs.
- `py migrate_existing_months.py` (from `scripts/`) produces `Johansen_BOHO_D4_Strategy_AllMonths.xlsx` with December's and January's real historical data and log entries correctly placed in their respective blocks, and January's own already-calibrated `Config` values carried over (not reverted to generic defaults).
- No workbook-scoped Excel defined names remain anywhere in the generated template (`Config` is 100% per-month direct cell references).
- The live source workbook's mtime/sha256 are unchanged from the Global Constraints section's recorded values.
- No `.xlsx`/`.xls`/`.xlsm` file has been committed to git.
- `docs/ACTIVATING_A_NEW_MONTH.md` and `docs/superpowers/HANDOFF.md` describe the new workflow accurately.
