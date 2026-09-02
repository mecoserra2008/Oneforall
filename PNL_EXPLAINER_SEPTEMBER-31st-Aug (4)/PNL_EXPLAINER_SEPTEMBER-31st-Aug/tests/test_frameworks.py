"""Spread-framework resolution and the chain identities.

The claim the whole model rests on is that a framework changes only HOW a yield
move is split, never the total.  These tests check that claim directly: build a
yield move out of its parts, run every chain over it, and require each to
reproduce -DV01 * dy.
"""

from __future__ import annotations

import unittest

import numpy as np

from pnlx.frameworks import (
    CHAINS,
    FRAMEWORKS,
    curve_for_framework,
    resolve_frameworks,
    spread_column_for_framework,
)


def _available(n: int, **flags: bool) -> dict[str, np.ndarray]:
    return {
        code: np.full(n, flags.get(code, True), dtype=bool)
        for code in ("G", "I", "ASW", "Z", "OAS")
    }


class TestChainIdentities(unittest.TestCase):
    """y = r + g + q + i, so every chain telescopes back to the yield move."""

    def setUp(self):
        # A move built from its parts, so the identity is exact by construction.
        self.dv01 = -25_000.0
        self.dr, self.dg, self.dq, self.di = 3.2, -0.8, 0.5, -1.4
        self.dy = self.dr + self.dg + self.dq + self.di
        self.dgspread = self.dq + self.di   # y - Gov = q + i

        self.legs = {
            "pnl_ois": -self.dv01 * self.dr,
            "pnl_gov_basis": -self.dv01 * self.dg,
            "pnl_swap_gov_basis": -self.dv01 * self.dq,
            "pnl_ispread": -self.dv01 * self.di,
            "pnl_gspread": -self.dv01 * self.dgspread,
            "pnl_yield_only": -self.dv01 * self.dy,
        }

    def test_the_g_chain_ties_exactly(self):
        chain = CHAINS["G"]
        total = sum(self.legs[leg] for leg in chain.legs)
        self.assertAlmostEqual(total, self.legs["pnl_yield_only"], places=9)
        self.assertTrue(chain.exact)

    def test_the_i_chain_ties_exactly(self):
        chain = CHAINS["I"]
        total = sum(self.legs[leg] for leg in chain.legs)
        self.assertAlmostEqual(total, self.legs["pnl_yield_only"], places=9)
        self.assertTrue(chain.exact)

    def test_the_two_exact_chains_agree_with_each_other(self):
        g = sum(self.legs[leg] for leg in CHAINS["G"].legs)
        i = sum(self.legs[leg] for leg in CHAINS["I"].legs)
        self.assertAlmostEqual(g, i, places=9)

    def test_ois_is_the_undecomposed_move(self):
        self.assertEqual(CHAINS["OIS"].legs, ("pnl_yield_only",))
        self.assertTrue(CHAINS["OIS"].exact)

    def test_the_quoted_conventions_are_declared_approximate(self):
        for code in ("ASW", "Z", "OAS"):
            self.assertFalse(CHAINS[code].exact, code)

    def test_review_and_mixed_have_no_fixed_chain(self):
        self.assertEqual(CHAINS["REVIEW"].legs, ())
        self.assertEqual(CHAINS["MIXED"].legs, ())

    def test_every_declared_framework_has_a_chain(self):
        for code in FRAMEWORKS:
            self.assertIn(code, CHAINS, code)


class TestBaseCurves(unittest.TestCase):
    def test_each_framework_reprices_on_the_curve_its_spread_is_over(self):
        codes = np.array(["G", "I", "ASW", "Z", "OAS", "OIS", "SOFR"], dtype=object)
        got = list(curve_for_framework(codes))
        self.assertEqual(got, ["GOV", "SWAP", "SWAP", "SWAP", "SWAP", "OIS", "OIS"])

    def test_review_has_no_base_curve(self):
        self.assertEqual(list(curve_for_framework(np.array(["REVIEW"], dtype=object))), [""])

    def test_the_spread_column_matches_the_framework(self):
        self.assertEqual(spread_column_for_framework("G"), "delta_gspread_bp")
        self.assertEqual(spread_column_for_framework("I"), "delta_ispread_bp")
        self.assertEqual(spread_column_for_framework("OIS"), "")


class TestResolution(unittest.TestCase):
    def setUp(self):
        self.isin = np.array(["AAA", "BBB", "CCC", "DDD"], dtype=object)

    def test_a_futures_dominated_hedge_measures_against_governments(self):
        got = resolve_frameworks(
            self.isin,
            futures_dv01=np.array([-10_000.0, 0.0, -5_000.0, 0.0]),
            plain_swap_dv01=np.array([0.0, -8_000.0, -9_000.0, 0.0]),
            available=_available(4),
        )
        self.assertEqual(got.code[0], "G")   # futures only
        self.assertEqual(got.code[1], "I")   # swaps only
        self.assertEqual(got.code[2], "I")   # swap-dominated
        self.assertEqual(got.code[3], "I")   # unhedged prefers I

    def test_dominance_ignores_the_sign_of_either_hedge(self):
        positive = resolve_frameworks(
            self.isin, np.array([10_000.0] * 4), np.array([1_000.0] * 4), _available(4)
        )
        negative = resolve_frameworks(
            self.isin, np.array([-10_000.0] * 4), np.array([-1_000.0] * 4), _available(4)
        )
        self.assertEqual(list(positive.code), list(negative.code))

    def test_a_synthetic_does_not_get_a_vote(self):
        # Synthetics never reach `resolve_frameworks`: only futures and PLAIN
        # swap DV01 are passed in.  With no real hedge the bond is unhedged.
        got = resolve_frameworks(
            self.isin, np.zeros(4), np.zeros(4), _available(4)
        )
        self.assertTrue(all(c == "I" for c in got.code))

    def test_a_framework_with_no_spread_quote_is_not_chosen(self):
        got = resolve_frameworks(
            self.isin,
            np.array([-10_000.0] * 4),   # futures-dominated, so G is preferred
            np.zeros(4),
            _available(4, G=False),      # ... but no G-spread is quoted
        )
        self.assertTrue(all(c == "I" for c in got.code))

    def test_no_spread_at_all_falls_to_review(self):
        got = resolve_frameworks(
            self.isin, np.zeros(4), np.zeros(4),
            _available(4, G=False, I=False, ASW=False, Z=False, OAS=False),
        )
        self.assertTrue(all(c == "REVIEW" for c in got.code))

    def test_the_global_override_wins_over_the_automatic_choice(self):
        got = resolve_frameworks(
            self.isin, np.array([-10_000.0] * 4), np.zeros(4), _available(4),
            global_override="ASW",
        )
        self.assertTrue(all(c == "ASW" for c in got.code))

    def test_a_per_bond_override_wins_over_the_global_one(self):
        got = resolve_frameworks(
            self.isin, np.zeros(4), np.zeros(4), _available(4),
            global_override="ASW",
            per_bond_override={"BBB": "OAS"},
            override_reasons={"BBB": "callable"},
        )
        self.assertEqual(got.code[1], "OAS")
        self.assertEqual(got.code[0], "ASW")
        self.assertIn("callable", got.reason[1])
        self.assertIn("per-bond override", got.reason[1])

    def test_an_unrecognised_override_becomes_review_not_silence(self):
        got = resolve_frameworks(
            self.isin, np.zeros(4), np.zeros(4), _available(4),
            per_bond_override={"AAA": "TYPO"},
        )
        self.assertEqual(got.code[0], "REVIEW")

    def test_an_override_is_honoured_even_with_no_spread_quote(self):
        # An override is a human decision.  Overruling it silently would hide
        # the fact that the data it needs is missing.
        got = resolve_frameworks(
            self.isin, np.zeros(4), np.zeros(4), _available(4, OAS=False),
            per_bond_override={"AAA": "OAS"},
        )
        self.assertEqual(got.code[0], "OAS")

    def test_hedge_weights_sum_to_one_when_hedged(self):
        got = resolve_frameworks(
            self.isin,
            np.array([-6_000.0, 0.0, 0.0, 0.0]),
            np.array([-2_000.0, 0.0, 0.0, 0.0]),
            _available(4),
        )
        self.assertAlmostEqual(float(got.futures_weight[0]), 0.75, places=9)
        self.assertAlmostEqual(float(got.swap_weight[0]), 0.25, places=9)
        self.assertEqual(float(got.futures_weight[1]), 0.0)


if __name__ == "__main__":
    unittest.main()
