"""The workbook's 2-variable Johansen hedge weights (Matrices_Johansen_Results
rows 2-26 and N18:N19), as a function of a DiffLag window. Mirrors the sheet's
formulas term by term (closed-form 2x2 algebra, no library eigen-solver) so
the results match the stored weights."""

import math

from phase2.inputs import MonthInputs

WINDOW = 63  # Config!B11
FIRST_ROW = 3  # DiffLag's first row with a difference


def johansen_weight(boho_by_row: dict, d4_by_row: dict, end_row: int, window: int = WINDOW) -> tuple:
    start = max(FIRST_ROW, end_row - window + 1)
    t = end_row - start + 1
    rows = range(start, end_row + 1)
    db = [boho_by_row[r] - boho_by_row[r - 1] for r in rows]
    dd = [d4_by_row[r] - d4_by_row[r - 1] for r in rows]
    lb = [boho_by_row[r - 1] for r in rows]
    ld = [d4_by_row[r - 1] for r in rows]

    def mom(x, y):
        return sum(a * b for a, b in zip(x, y)) / t

    s00 = (mom(db, db), mom(db, dd), mom(dd, db), mom(dd, dd))
    s11 = (mom(lb, lb), mom(lb, ld), mom(ld, lb), mom(ld, ld))
    s01 = (mom(db, lb), mom(db, ld), mom(dd, lb), mom(dd, ld))
    s10 = (s01[0], s01[2], s01[1], s01[3])

    def inv(m):
        det = m[0] * m[3] - m[1] * m[2]
        return (m[3] / det, -m[1] / det, -m[2] / det, m[0] / det)

    def mul(x, y):
        return (x[0] * y[0] + x[1] * y[2], x[0] * y[1] + x[1] * y[3],
                x[2] * y[0] + x[3] * y[2], x[2] * y[1] + x[3] * y[3])

    a, b, _c, d = mul(mul(mul(inv(s11), s10), inv(s00)), s01)
    tr, det = a + d, a * d - b * _c
    lam1 = (tr + math.sqrt(tr ** 2 - 4 * det)) / 2
    lam2 = (tr - math.sqrt(tr ** 2 - 4 * det)) / 2
    lam = lam1 if lam1 >= lam2 else lam2
    v1, v2 = b, lam - a
    n = math.sqrt(v1 ** 2 + v2 ** 2)
    return (-v1 / n, -v2 / n) if v2 > 0 else (v1 / n, v2 / n)


def monthly_weights(m: MonthInputs, boho: list = None) -> list:
    """(end row, month-end date, BOHO weight, D4 weight) at each stored month-end."""
    b = dict(zip(m.rows, boho if boho is not None else m.boho))
    d = dict(zip(m.rows, m.d4))
    return [(end, dt, *johansen_weight(b, d, end)) for end, dt, _wb, _wd in m.stored]


def gate_report(m: MonthInputs) -> dict:
    worst = (None, None, None)
    max_diff = 0.0
    for (end, dt, sb, sd), (_e, _d, rb, rd) in zip(m.stored, monthly_weights(m)):
        diff = max(abs(sb - rb), abs(sd - rd))
        if diff > max_diff:
            max_diff, worst = diff, (dt, (sb, sd), (rb, rd))
    return {"max_abs_diff": max_diff, "worst": worst, "months": len(m.stored)}
