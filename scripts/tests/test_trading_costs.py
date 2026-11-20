import pathlib
import sys
from datetime import date, timedelta

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

from trading_costs import COSTS_PER_SIDE, daily_costs, held_tiers

D = [date(2024, 1, 1) + timedelta(days=i) for i in range(8)]
FULL = 84 * 3.66 + 16 * 3.01  # $355.60 to trade a full position one way


def test_full_position_fee_per_side():
    assert COSTS_PER_SIDE["fees"] == pytest.approx(FULL)
    assert COSTS_PER_SIDE["slippage_1tick"] == pytest.approx(84 * 6.00 + 16 * 4.20)


def test_held_follows_effective_trade_active():
    tiers = [(0, 0, 0), (0.25, 0, 0), (0.25, 0.5, 0), (0.25, 0.5, 0), (0.25, 0.5, 0.25), (0.25, 0.5, 0.25), (0, 0, 0), (0, 0, 0)]
    bs = [0, 0, 1, 1, 1, 0, 0, 0]  # PnL starts the day after the entry; the overlay exits on day 5
    held = held_tiers(tiers, bs)
    assert held == [(0, 0, 0), (0, 0, 0), (0.25, 0, 0), (0.25, 0.5, 0), (0.25, 0.5, 0), (0, 0, 0), (0, 0, 0), (0, 0, 0)]


def test_costs_charge_every_side_and_rolls():
    held = [(0, 0, 0), (0.25, 0, 0), (0.25, 0.5, 0), (-0.25, 0, 0), (0, 0, 0), (0, 0, 0)]
    c = daily_costs(D[:6], held, roll_dates=[D[2]], per_side=FULL)
    assert c[1] == pytest.approx(0.25 * FULL)              # buy tier 1
    assert c[2] == pytest.approx(0.5 * FULL + 2 * 0.75 * FULL)  # add tier 2, then roll 0.75 (two sides)
    assert c[3] == pytest.approx((0.5 + 0.5) * FULL)       # tier 1 flips long -> short (0.5), tier 2 closes (0.5)
    assert c[4] == pytest.approx(0.25 * FULL)              # close
    assert sum(c) == pytest.approx((0.25 + 0.5 + 1.5 + 1.0 + 0.25) * FULL)
