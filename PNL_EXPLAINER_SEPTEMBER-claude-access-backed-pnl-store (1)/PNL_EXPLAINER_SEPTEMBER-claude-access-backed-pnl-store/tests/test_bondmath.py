"""Pricing, risk and pull to par.

The pull-to-par tests are the important ones.  The method is forward pricing on
the prior curve, and the property that proves it is right is the no-arbitrage
forward identity:

    DirtyPx_fwd = DirtyPx_spot * (1 + y(h)) ** h        (no coupon inside h)

If that holds on a flat curve, a steep one and an inverted one, the calculation
is doing what it claims.  If it only holds on the flat one, it is repricing at
a shorter maturity and booking roll-down as PnL.
"""

from __future__ import annotations

import unittest

import numpy as np

from pnlx import bondmath as bm
from pnlx import daycount as dc
from pnlx.curves import CurveKey, CurveSet, ZeroCurve

TENORS = np.array([0.25, 1.0, 2.0, 5.0, 10.0, 30.0])


def curve_set(shape) -> CurveSet:
    """A CurveSet where every curve and snapshot takes the same shape."""
    cs = CurveSet()
    for kind in ("OIS", "GOV", "SWAP"):
        for snapshot in ("prior", "current"):
            cs.add(ZeroCurve.build(CurveKey("EUR", kind, snapshot), TENORS, shape(TENORS)))
    return cs


FLAT = curve_set(lambda t: np.full_like(t, 3.0))
STEEP = curve_set(lambda t: 1.0 + 0.15 * t)
INVERTED = curve_set(lambda t: 4.0 - 0.06 * t)


class TestPriceFromYield(unittest.TestCase):
    def test_par_bond_on_a_coupon_date_is_exactly_100(self):
        clean, dirty = bm.bond_price_from_yield(
            dc.as_date_array(["2025-06-15"]),
            dc.as_date_array(["2030-06-15"]),
            np.array([0.04]), np.array([0.04]),
            np.array([1.0]), np.array([0.0]),
        )
        self.assertAlmostEqual(float(clean[0]), 100.0, places=9)
        self.assertAlmostEqual(float(dirty[0]), 100.0, places=9)

    def test_discount_and_premium_sit_the_right_side_of_par(self):
        settle = dc.as_date_array(["2025-06-15"] * 2)
        maturity = dc.as_date_array(["2030-06-15"] * 2)
        clean, _ = bm.bond_price_from_yield(
            settle, maturity, np.array([0.02, 0.06]), np.array([0.04, 0.04]),
            np.array([2.0, 2.0]), np.array([0.0, 0.0]),
        )
        self.assertLess(float(clean[0]), 100.0)   # coupon below yield
        self.assertGreater(float(clean[1]), 100.0)  # coupon above yield

    def test_dirty_less_clean_is_the_accrued(self):
        settle = dc.as_date_array(["2025-12-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        clean, dirty = bm.bond_price_from_yield(
            settle, maturity, np.array([0.04]), np.array([0.035]),
            np.array([1.0]), np.array([0.0]),
        )
        accrued = dc.accrued_interest(
            settle, maturity, np.array([0.04]), np.array([1.0]), np.array([0.0])
        )
        self.assertAlmostEqual(float(dirty[0] - clean[0]), float(accrued[0]), places=9)

    def test_a_settled_or_matured_bond_is_nan_not_zero(self):
        clean, dirty = bm.bond_price_from_yield(
            dc.as_date_array(["2031-01-01"]),
            dc.as_date_array(["2030-06-15"]),
            np.array([0.04]), np.array([0.04]),
            np.array([1.0]), np.array([0.0]),
        )
        self.assertTrue(np.isnan(clean[0]) and np.isnan(dirty[0]))


class TestRisk(unittest.TestCase):
    def test_duration_is_positive_and_shorter_than_maturity(self):
        settle = dc.as_date_array(["2025-06-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        dur, cvx = bm.yield_risk(
            settle, maturity, np.array([0.02]), np.array([0.04]),
            np.array([2.0]), np.array([0.0]), bump_bp=1.0,
        )
        self.assertGreater(float(dur[0]), 0.0)
        self.assertLess(float(dur[0]), 5.0)
        self.assertGreater(float(cvx[0]), 0.0)

    def test_a_zero_coupon_duration_is_its_maturity(self):
        settle = dc.as_date_array(["2025-06-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        dur, _ = bm.yield_risk(
            settle, maturity, np.array([0.0]), np.array([0.04]),
            np.array([1.0]), np.array([0.0]), bump_bp=0.01,
        )
        # Modified duration of a 5-year zero at 4% annual = 5 / 1.04.
        self.assertAlmostEqual(float(dur[0]), 5.0 / 1.04, places=3)

    def test_bumped_and_analytic_convexity_agree_closely(self):
        """A 100bp bump measures AVERAGE convexity across the bump.

        It should not equal the analytic value exactly - the difference is the
        third-order term, and on a ten-year bond it is around a tenth of a
        percent.  The test is that they agree closely, not that they agree: an
        exact match would mean the bump size was being ignored.
        """
        settle = dc.as_date_array(["2025-06-15"])
        maturity = dc.as_date_array(["2035-06-15"])
        _, analytic = bm.yield_risk(
            settle, maturity, np.array([0.03]), np.array([0.035]),
            np.array([2.0]), np.array([0.0]), bump_bp=0.01,
        )
        _, bumped = bm.yield_risk(
            settle, maturity, np.array([0.03]), np.array([0.035]),
            np.array([2.0]), np.array([0.0]), bump_bp=100.0,
        )
        relative = abs(float(bumped[0]) - float(analytic[0])) / float(analytic[0])
        self.assertLess(relative, 0.005)
        self.assertGreater(relative, 0.0)


class TestPullToPar(unittest.TestCase):
    """The forward identity, on three curve shapes."""

    def _run(self, curves, maturity="2032-03-15", coupon=0.035, freq=1.0, spread=25.0):
        n = 1
        return bm.pull_to_par(
            dc.as_date_array(["2025-09-29"]),
            dc.as_date_array(["2025-09-30"]),
            dc.as_date_array([maturity]),
            np.array([coupon]), np.array([freq]), np.array([0.0]),
            np.array(["EUR"], dtype=object), np.array(["GOV"], dtype=object),
            np.array([spread]), curves,
        )

    def _identity_error(self, curves, result):
        curve = curves.require("EUR", "GOV", "prior")
        horizon = float(result.horizon_years[0])
        rate = float(curve.rate(np.array([horizon]))[0]) / 100.0 + 0.0025
        implied = float(result.dirty_spot[0]) * (1.0 + rate) ** horizon
        return abs(float(result.dirty_forward[0]) - implied)

    def test_forward_identity_holds_on_a_flat_curve(self):
        result = self._run(FLAT)
        self.assertLess(self._identity_error(FLAT, result), 1e-9)

    def test_forward_identity_holds_on_a_steep_curve(self):
        result = self._run(STEEP)
        self.assertLess(self._identity_error(STEEP, result), 1e-9)

    def test_forward_identity_holds_on_an_inverted_curve(self):
        result = self._run(INVERTED)
        self.assertLess(self._identity_error(INVERTED, result), 1e-9)

    def test_a_steep_curve_gives_less_pull_to_par_than_a_flat_one(self):
        # This is the whole reason forward pricing is used.  Repricing at a
        # shorter maturity on the spot curve would book the roll-down of an
        # upward-sloping curve as PnL the desk never earned, so the steep case
        # would look BETTER than the flat one rather than worse.
        flat = float(self._run(FLAT).clean_change[0])
        steep = float(self._run(STEEP).clean_change[0])
        self.assertLess(steep, flat)

    def test_a_premium_bond_pulls_down_on_a_flat_curve(self):
        result = self._run(FLAT, coupon=0.06)
        self.assertGreater(float(result.dirty_spot[0]), 100.0)
        self.assertLess(float(result.clean_change[0]), 0.0)

    def test_a_discount_bond_pulls_up_on_a_flat_curve(self):
        result = self._run(FLAT, coupon=0.005)
        self.assertLess(float(result.dirty_spot[0]), 100.0)
        self.assertGreater(float(result.clean_change[0]), 0.0)

    def test_a_coupon_inside_the_horizon_is_excluded(self):
        # Settlement straddles a payment date.  The coupon is cash in the bank
        # and belongs to the carry leg; compounding it into the forward price
        # would double it.
        result = bm.pull_to_par(
            dc.as_date_array(["2032-03-10"]),
            dc.as_date_array(["2032-03-20"]),
            dc.as_date_array(["2040-03-15"]),
            np.array([0.05]), np.array([1.0]), np.array([0.0]),
            np.array(["EUR"], dtype=object), np.array(["GOV"], dtype=object),
            np.array([25.0]), FLAT,
        )
        self.assertTrue(bool(result.usable[0]))
        # Nowhere near a whole coupon: this is ten days of pull, not 5 points.
        self.assertLess(abs(float(result.clean_change[0])), 0.5)

    def test_bad_inputs_are_nan_not_zero(self):
        bad = bm.pull_to_par(
            dc.as_date_array(["2025-09-29"] * 3),
            dc.as_date_array(["2025-09-30"] * 3),
            dc.as_date_array(["2024-01-01", "2032-03-15", "2032-03-15"]),  # already matured
            np.array([0.035, 0.035, 0.035]),
            np.array([1.0, 3.0, 1.0]),        # 3 is not a valid frequency
            np.array([0.0, 0.0, 99.0]),       # 99 is not a valid convention
            np.array(["EUR"] * 3, dtype=object),
            np.array(["GOV"] * 3, dtype=object),
            np.array([25.0, 25.0, 25.0]),
            FLAT,
        )
        self.assertTrue(np.all(np.isnan(bad.clean_change)))

    def test_model_dirty_price_matches_the_spot_leg(self):
        result = self._run(FLAT)
        price = bm.model_dirty_price(
            dc.as_date_array(["2025-09-29"]),
            dc.as_date_array(["2032-03-15"]),
            np.array([0.035]), np.array([1.0]),
            np.array(["EUR"], dtype=object), np.array(["GOV"], dtype=object),
            np.array([25.0]), FLAT,
        )
        self.assertAlmostEqual(float(price[0]), float(result.dirty_spot[0]), places=9)


if __name__ == "__main__":
    unittest.main()
