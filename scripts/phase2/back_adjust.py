"""Back-adjusted (shifted) BOHO series: each roll's old->new contract gap is
added to every earlier price, so the latest price stays real and the roll
day's change is the old contract's own move."""

from phase2.inputs import MonthInputs

ZL_MULTIPLIER = 7.5  # Config row 8: SBO cents/lb -> $/gal-equivalent factor


def boho_from_legs(sbo: float, ho: float) -> float:
    return sbo * ZL_MULTIPLIER / 100 - ho


def roll_gaps(m: MonthInputs) -> list:
    """(roll date, gap) per roll inside the data: new-contract BOHO on the
    roll day (Data) minus the old contract's BOHO that day (Roll Prices)."""
    index = {d: i for i, d in enumerate(m.dates)}
    gaps = []
    for d, old_sbo, old_ho in m.rolls:
        if m.dates and d < m.dates[0]:
            continue
        if d not in index:
            raise ValueError(f"{m.month}: roll date {d.isoformat()} is not a date in Data")
        gaps.append((d, m.boho[index[d]] - boho_from_legs(old_sbo, old_ho)))
    return gaps


def back_adjust(dates: list, boho: list, gaps: list) -> list:
    """Each price before a roll date gets that roll's gap added."""
    return [p + sum(g for rd, g in gaps if d < rd) for d, p in zip(dates, boho)]
