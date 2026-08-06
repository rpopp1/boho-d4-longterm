import re

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter


def extract_row_formulas(path: str, sheet_name: str, row: int, max_col: int) -> dict[str, object]:
    """Extract formulas and values from a row in an Excel workbook.

    Returns a dictionary mapping column letters to their cell values.
    Formula cells return formula strings, array formulas are unwrapped to their text,
    and non-formula cells (e.g., literal numbers, strings) are returned as-is.
    Empty cells are omitted from the result.
    """
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
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
    finally:
        wb.close()


# Matches an optional sheet-qualifier, optional $ + column letters, optional $ + row digits.
# Full-column refs like G:G (no digits) simply won't match the row-digit group and pass through.
# The leading negative lookbehind prevents matching in the middle of a longer identifier/defined
# name (e.g. the "r1" tail of "ZThreshold_Tier1" must never be treated as a cell reference).
_CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?P<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<colabs>\$)?(?P<col>[A-Za-z]{1,3})"
    r"(?P<rowabs>\$)?(?P<row>\d+)"
    r"(?![A-Za-z0-9_(])"
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


def rewrite_constants(formula: str, constant_map: dict) -> str:
    """Rewrite constant cell addresses in a formula to their defined names.

    Args:
        formula: A formula string (e.g., "=IF($DV$6>0,$DK$5,0)")
        constant_map: A dict mapping addresses to defined names (e.g., {"$DV$6": "ZExitThreshold"})

    Returns:
        The formula with addresses replaced by their defined names.

    Note:
        Addresses are sorted by length (longest first). This does NOT, by
        itself, prevent "$DK$2" from matching inside "$DK$25" -- a plain
        substring replace would still hit that substring regardless of
        iteration order, since $DK$2 is literally a prefix of $DK$25's
        digits. The actual protection is empirical, not structural: verified
        elsewhere in this project's history (see the Task 4/Task 7 reports),
        none of CONSTANT_MAP's real addresses collide as a prefix of another
        real cell reference actually present in the live workbook's
        formulas. The longest-first ordering only matters for the case where
        the SHORTER key is itself also a substring of a LONGER key (e.g. both
        "$DK$2" and "$DK$25" happen to be in constant_map at once) -- it
        guarantees the longer, more specific key gets its chance to match
        first rather than being partially consumed by the shorter one.
    """
    # Longest addresses first: only matters when a shorter map key is itself a
    # substring of a longer map key (see docstring Note) -- ensures the longer
    # key isn't partially clobbered before it gets its own chance to match.
    for address in sorted(constant_map, key=len, reverse=True):
        formula = formula.replace(address, constant_map[address])
    return formula


def unwrap_scalar_array_formula(value, coordinate: str):
    """If `value` is an openpyxl ArrayFormula whose `ref` spans exactly the
    single cell at `coordinate` (i.e. a legacy CSE array formula stored for
    historical reasons, not a true multi-cell/spill formula), return its
    plain formula text (e.g. "=INDEX(DiffLag!E:E,$B$2)") so it can be written
    as an ordinary string formula. Any other value (including a genuinely
    multi-cell ArrayFormula) is returned unchanged for str/None/etc, but a
    multi-cell ArrayFormula raises -- literal-copy loops have no array/spill
    handling at all (unlike sheet_cloner.fill_sheet_from_pattern's explicit
    NotImplementedError guard), so silently writing a multi-cell ArrayFormula
    object into a new cell would risk corrupting or losing its spill
    semantics with no warning.
    """
    from openpyxl.worksheet.formula import ArrayFormula

    if not isinstance(value, ArrayFormula):
        return value
    if value.ref != coordinate:
        raise NotImplementedError(
            f"ArrayFormula at {coordinate} has a multi-cell ref '{value.ref}' -- "
            "this literal-copy loop has no dynamic-array/spill handling (see "
            "sheet_cloner.fill_sheet_from_pattern's NotImplementedError guard "
            "for the same failure mode). Redesign the caller to avoid copying "
            "this cell as-is rather than risk corrupting its spill semantics."
        )
    return value.text
