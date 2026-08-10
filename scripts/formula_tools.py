import re

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter, column_index_from_string


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


_SHEET_QUALIFIER = r"(?:'(?P<qname>[^']+)'|(?P<uname>[A-Za-z_][A-Za-z0-9_.]*))!"

# NOTE (multi-month Task 11): the reference body below is deliberately matched
# as ONE unit -- an optional sheet qualifier followed by either a single cell,
# a cell:cell range, or a full-column col:col range, each side independently
# `$`-anchorable. Two real defects, both found while wiring up Position
# Model's cross-sheet relocation pass, are why it is shaped this way rather
# than as two narrow alternatives:
#
#   1. "$G:$G" (an ANCHORED full-column range) did not match the old
#      full-column alternative at all (it required bare "G:G"), so
#      shift_formula_columns SILENTLY returned the whole formula unchanged for
#      Position Model!I2's "LOOKUP($A2-1, Johansen_Monthly!$G:$G,
#      Johansen_Monthly!$H:$H)" -- every month's Position Model tab would have
#      kept reading January's Johansen_Monthly walk-forward log (columns
#      G/H/I) instead of its own. Caught by
#      test_position_model_march_johansen_monthly_reference_relocated.
#
#   2. "Data!$A$3:$A$301" was matched as TWO separate references, and only the
#      first one carried the "Data!" qualifier -- the second endpoint's
#      qualifier is implicit in Excel, so it looked unqualified, was not a
#      sheet_map key, and was left unshifted. That produced the actively
#      CORRUPT "Data!$W$3:$A$301" (a range spanning from December's block back
#      to January's) rather than merely an unshifted formula. Position Model's
#      DE43/DF43/DF46/DG49/DH49 "last non-blank row" LOOKUPs use exactly this
#      shape, so this was a live corruption, not a hypothetical one. Matching
#      the whole range as one body makes the qualifier apply to both endpoints,
#      as Excel itself does.
#
# A read-only scan of the live source workbook found $-anchored full-column
# refs in exactly three places, all of them the same Johansen_Monthly!$G/$H/$I
# LOOKUP idiom (Position Model!I2, Johansen_Monthly!O3, Spread!D2); only
# Position Model's copy is reproduced by any builder, so no already-migrated
# sheet's output changes as a result of these fixes.
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_REF_BODY = r"(?:" + _CELL + r"(?::" + _CELL + r")?" + r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})"

_COL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:" + _SHEET_QUALIFIER + r")?"
    r"(?P<body>" + _REF_BODY + r")"
    r"(?![A-Za-z0-9_(])"
)

# Every column token inside a matched reference body: an optional "$" anchor
# plus 1-3 letters. Row digits contain no letters, so they can never match.
_BODY_COL_TOKEN_RE = re.compile(r"(\$?)([A-Za-z]{1,3})")

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
    reference into a relocated sheet must move, anchored or not. The $ anchors
    themselves are preserved byte-for-byte in the output.

    A ":"-range is matched as a single reference, so a qualifier on the left
    endpoint applies to the right endpoint too (as Excel itself treats it) --
    see _COL_REF_RE's note for the corruption that resulted from not doing so.
    """
    def _replace(match: "re.Match") -> str:
        sheet_key = match.group("qname") or match.group("uname")
        if sheet_key not in sheet_map:
            return match.group(0)
        col_delta, new_sheet = sheet_map[sheet_key]
        prefix = _qualify_sheet(new_sheet)

        def _shift_token(token: "re.Match") -> str:
            anchor, col = token.group(1), token.group(2)
            return f"{anchor}{get_column_letter(column_index_from_string(col) + col_delta)}"

        return prefix + _BODY_COL_TOKEN_RE.sub(_shift_token, match.group("body"))

    return _COL_REF_RE.sub(_replace, formula)
