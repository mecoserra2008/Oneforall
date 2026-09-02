"""The attribution engine, end to end.

Most of these run against a tiny hand-built book rather than the sample, so the
expected answer can be worked out on paper.  The bridge-closure tests then run
against the full sample book, because closure has to hold on a book with
missing data, an ex-coupon date, a short, and unlinked hedges in it - a book
where everything is clean proves nothing.
"""

from __future__ import annotations

import datetime as _dt
import unittest
from dataclasses import replace

import numpy as np

from pnlx.aggregate import build_portfolio
from pnlx.config import AppConfig, ModelConfig, PathsConfig, RunConfig, load_config
from pnlx.curves import CurveKey, CurveSet, ZeroCurve
from pnlx.engine import AttributionEngine
from pnlx.instruments import BondBook, BondPosition, FutureBook, SwapBook
from pnlx.loaders import InputBundle, load_inputs

TENORS = np.array([0.25, 1.0, 2.0, 5.0, 10.0, 30.0])


def flat_curves(ois=3.0, gov=3.2, swap=3.4, shift_bp=0.0) -> CurveSet:
    """Flat curves, with the CURRENT snapshot shifted by `shift_bp`."""
    cs = CurveSet()
    for kind, level in (("OIS", ois), ("GOV", gov), ("SWAP", swap)):
        cs.add(ZeroCurve.build(CurveKey("EUR", kind, "prior"), TENORS, np.full_like(TENORS, level)))
        cs.add(
            ZeroCurve.build(
                CurveKey("EUR", kind, "current"), TENORS,
                np.full_like(TENORS, level + shift_bp / 100.0),
            )
        )
    return cs


def one_bond(**overrides) -> BondBook:
    defaults = dict(
        isin="TEST0000001",
        name="TEST BOND",
        currency="EUR",
        portfolio="P1",
        notional=10_000_000.0,
        coupon_pct=3.0,
        coupon_freq=1.0,
        maturity=_dt.date(2035, 6, 15),
        day_count_code=0.0,
        clean_px_prior=98.0,
        dirty_px_prior=99.0,
        ytm_prior=3.4,
        clean_px_current=98.0,
        dirty_px_current=99.0,
        ytm_current=3.4,
        mod_duration_prior=8.0,
        mod_duration_current=8.0,
        convexity=80.0,
        zspread_prior=20.0, zspread_current=20.0,
        asw_prior=18.0, asw_current=18.0,
        oas_prior=15.0, oas_current=15.0,
        fx_prior=1.0, fx_current=1.0,
    )
    defaults.update(overrides)
    return BondBook.from_positions([BondPosition(**defaults)])


def config_for(model: ModelConfig | None = None) -> AppConfig:
    return AppConfig(
        run=RunConfig(as_of=_dt.date(2025, 9, 30), prior=_dt.date(2025, 9, 29)),
        paths=PathsConfig(),
        model=model or ModelConfig(),
    )


def bundle(bonds: BondBook, curves: CurveSet) -> InputBundle:
    return InputBundle(
        bonds=bonds,
        futures=FutureBook.from_positions([]),
        swaps=SwapBook.from_positions([]),
        curves=curves,
        fx={"EUR": (1.0, 1.0)},
        override_codes={},
        override_notes={},
    )


class TestFirstOrder(unittest.TestCase):
    def test_dv01_is_duration_times_price_times_notional(self):
        book = one_bond()
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        # 10m * 8.0 * 99/100 * 1e-4 = 7,920 per bp
        self.assertAlmostEqual(float(result.bonds["dv01_prior"][0]), 7_920.0, places=6)

    def test_a_short_position_has_negative_dv01(self):
        book = one_bond(notional=-10_000_000.0)
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        self.assertAlmostEqual(float(result.bonds["dv01_prior"][0]), -7_920.0, places=6)

    def test_a_rate_rise_loses_money_on_a_long(self):
        # Yield up 10bp, curve up 10bp: PnL should be -DV01 * 10.
        book = one_bond(ytm_current=3.5)
        result = AttributionEngine(config_for(), bundle(book, flat_curves(shift_bp=10.0))).run()
        frame = result.bonds
        self.assertAlmostEqual(float(frame["delta_y_bp"][0]), 10.0, places=6)
        self.assertAlmostEqual(float(frame["pnl_yield_only"][0]), -79_200.0, places=3)
        self.assertLess(float(frame["pnl_duration_total"][0]), 0.0)

    def test_the_chain_reproduces_the_yield_move(self):
        book = one_bond(ytm_current=3.5)
        result = AttributionEngine(config_for(), bundle(book, flat_curves(shift_bp=6.0))).run()
        frame = result.bonds
        self.assertAlmostEqual(float(frame["duration_identity_check"][0]), 0.0, places=6)

    def test_a_short_gains_when_rates_rise(self):
        book = one_bond(notional=-10_000_000.0, ytm_current=3.5)
        result = AttributionEngine(config_for(), bundle(book, flat_curves(shift_bp=10.0))).run()
        self.assertGreater(float(result.bonds["pnl_yield_only"][0]), 0.0)


class TestRiskDateConvention(unittest.TestCase):
    def test_attribution_uses_opening_risk_by_default(self):
        # Price falls, so closing risk is smaller than opening risk.
        book = one_bond(dirty_px_current=95.0, ytm_current=3.9)
        result = AttributionEngine(config_for(), bundle(book, flat_curves(shift_bp=50.0))).run()
        frame = result.bonds
        self.assertAlmostEqual(
            float(frame["attribution_dv01"][0]), float(frame["dv01_prior"][0]), places=9
        )
        self.assertNotAlmostEqual(
            float(frame["dv01_prior"][0]), float(frame["dv01_current"][0]), places=3
        )

    def test_the_legacy_convention_can_be_selected_and_differs(self):
        book = one_bond(dirty_px_current=95.0, ytm_current=3.9)
        legacy = config_for(replace(ModelConfig(), attribution_risk_date="current"))
        result = AttributionEngine(legacy, bundle(book, flat_curves(shift_bp=50.0))).run()
        self.assertAlmostEqual(
            float(result.bonds["attribution_dv01"][0]),
            float(result.bonds["dv01_current"][0]),
            places=9,
        )

    def test_the_timing_bias_is_the_difference_between_the_two(self):
        book = one_bond(dirty_px_current=95.0, ytm_current=3.9)
        opening = AttributionEngine(config_for(), bundle(book, flat_curves(shift_bp=50.0))).run()
        legacy = AttributionEngine(
            config_for(replace(ModelConfig(), attribution_risk_date="current")),
            bundle(one_bond(dirty_px_current=95.0, ytm_current=3.9), flat_curves(shift_bp=50.0)),
        ).run()
        reported = float(opening.bonds["dv01_timing_bias"][0])
        actual = float(legacy.bonds["pnl_yield_only"][0]) - float(opening.bonds["pnl_yield_only"][0])
        self.assertAlmostEqual(reported, actual, places=6)


class TestCarry(unittest.TestCase):
    def test_exact_carry_survives_an_ex_coupon_date(self):
        """A coupon paying inside the period must not open a residual.

        This is the case the legacy sheet had no term for: the dirty price
        drops by the coupon and, with only a smooth accrual on the explained
        side, that bond showed a one-day break the size of its coupon.
        """
        # Coupon pays on 30 September, the closing date.
        book = one_bond(maturity=_dt.date(2031, 9, 30), coupon_pct=5.0)
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        frame = result.bonds
        self.assertGreater(float(frame["coupons_paid_per_100"][0]), 4.9)
        # Carry is one day of accrual, not a whole coupon.
        self.assertLess(abs(float(frame["carry_coupon"][0])), 10_000.0)
        # The cash appears on the practical side.
        self.assertAlmostEqual(
            float(frame["coupon_cash"][0]),
            10_000_000.0 / 100.0 * float(frame["coupons_paid_per_100"][0]),
            places=6,
        )

    def test_smooth_carry_is_notional_times_coupon_times_year_fraction(self):
        smooth = config_for(replace(ModelConfig(), coupon_carry="smooth"))
        result = AttributionEngine(smooth, bundle(one_bond(), flat_curves())).run()
        expected = 10_000_000.0 * 0.03 * (1.0 / 365.0)
        self.assertAlmostEqual(float(result.bonds["carry_coupon"][0]), expected, places=6)

    def test_carry_is_positive_on_a_quiet_day_for_a_long(self):
        result = AttributionEngine(config_for(), bundle(one_bond(), flat_curves())).run()
        self.assertGreater(float(result.bonds["carry_total"][0]), 0.0)


class TestFx(unittest.TestCase):
    def test_the_fx_decomposition_is_an_exact_identity(self):
        """dMV_base = FX translation + local move at the opening fix + cross."""
        book = one_bond(
            currency="USD", fx_prior=0.85, fx_current=0.87,
            dirty_px_current=100.5, ytm_current=3.3,
        )
        curves = flat_curves()
        for kind in ("OIS", "GOV", "SWAP"):
            for snapshot in ("prior", "current"):
                curve = curves.require("EUR", kind, snapshot)
                curves.add(ZeroCurve.build(CurveKey("USD", kind, snapshot), curve.tenors, curve.rates))

        result = AttributionEngine(config_for(), bundle(book, curves)).run()
        frame = result.bonds

        d_mv = float(frame["delta_mv_base"][0])
        fx = float(frame["pnl_fx"][0])
        cross = float(frame["pnl_fx_cross"][0])
        local_at_open = (
            float(frame["notional"][0])
            * (float(frame["dirty_px_current"][0]) - float(frame["dirty_px_prior"][0]))
            / 100.0
            * float(frame["fx_prior"][0])
        )
        self.assertAlmostEqual(d_mv, fx + local_at_open + cross, places=6)

    def test_a_base_currency_position_has_no_fx_leg(self):
        result = AttributionEngine(config_for(), bundle(one_bond(), flat_curves())).run()
        self.assertEqual(float(result.bonds["pnl_fx"][0]), 0.0)
        self.assertEqual(float(result.bonds["pnl_fx_cross"][0]), 0.0)


class TestMissingData(unittest.TestCase):
    def test_a_missing_mark_blanks_rather_than_zeroes(self):
        book = one_bond(ytm_current=float("nan"))
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        frame = result.bonds
        # Nothing derived from the missing yield may be fabricated.
        for column in ("delta_y_bp", "delta_ispread_bp", "delta_gspread_bp",
                       "pnl_yield_only", "duration_identity_check"):
            self.assertTrue(np.isnan(frame[column][0]), column)
        self.assertNotEqual(frame["attribution_status"][0], "OK")

    def test_a_missing_yield_falls_back_to_a_quoted_spread_framework(self):
        """No yield, but a quoted ASW: the model measures against ASW instead.

        This is deliberate.  The frameworks whose spread is DERIVED from the
        yield (G and I) become unavailable, so the resolver drops to one that
        is quoted independently.  The row still reports 'Missing market data',
        because the identity check cannot be computed without a yield and
        nobody should read the number as fully corroborated.
        """
        book = one_bond(ytm_current=float("nan"))
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        frame = result.bonds
        self.assertIn(str(frame["spread_framework"][0]), ("ASW", "Z", "OAS"))
        self.assertFalse(np.isnan(frame["pnl_duration_total"][0]))
        self.assertTrue(np.isnan(frame["duration_identity_check"][0]))
        self.assertEqual(frame["attribution_status"][0], "Missing market data")

    def test_no_usable_spread_at_all_suppresses_attribution(self):
        book = one_bond(
            ytm_current=float("nan"),
            zspread_current=float("nan"),
            asw_current=float("nan"),
            oas_current=float("nan"),
        )
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        frame = result.bonds
        self.assertEqual(str(frame["spread_framework"][0]), "REVIEW")
        self.assertTrue(np.isnan(frame["pnl_duration_total"][0]))
        self.assertTrue(np.isnan(frame["total_explained"][0]))

    def test_an_incomplete_chain_gives_no_partial_total(self):
        # No curve at all for this bond's currency.
        book = one_bond(currency="JPY")
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        self.assertTrue(np.isnan(result.bonds["pnl_duration_total"][0]))

    def test_an_unknown_day_count_is_reported(self):
        book = one_bond(day_count_code=float("nan"), day_count_desc="WIBBLE")
        result = AttributionEngine(config_for(), bundle(book, flat_curves())).run()
        self.assertIn("day-count", result.bonds["attribution_status"][0])


class TestSampleBookBridge(unittest.TestCase):
    """Closure on the real sample book, imperfections and all."""

    @classmethod
    def setUpClass(cls):
        cls.config = load_config()
        cls.result = AttributionEngine(cls.config, load_inputs(cls.config)).run()
        cls.summary = build_portfolio(cls.result)

    def test_every_tie_out_closes(self):
        for line in self.summary.bridge.lines:
            if line.kind == "check":
                self.assertLess(abs(line.amount), 1e-4, line.label)

    def test_explained_is_theoretical_plus_hedge_basis(self):
        b = self.summary.bridge
        self.assertAlmostEqual(
            b.value("Theoretical PnL") + b.value("Hedge basis (actual - model)"),
            b.value("Total explained"),
            places=6,
        )

    def test_the_bond_cash_leg_residual_equals_the_bridge_residual(self):
        # The hedge legs telescope, so these must be the same number.  If they
        # drift apart, a population mismatch has crept into one of them.
        bonds = next(ac for ac in self.summary.asset_classes if ac.name.startswith("Bonds"))
        self.assertAlmostEqual(
            bonds.residual, self.summary.bridge.value("Unexplained residual"), places=6
        )

    def test_the_memo_lines_tie_to_the_duration_total(self):
        memo = sum(
            line.amount for line in self.summary.bridge.memos
            if line.label.startswith("of which")
        )
        self.assertAlmostEqual(
            memo, self.summary.bridge.value("Rates and spread (duration)"), places=6
        )

    def test_exact_frameworks_tie_to_machine_precision(self):
        frame = self.result.bonds
        codes = np.asarray(frame["spread_framework"], dtype=object)
        identity = np.asarray(frame["duration_identity_check"], dtype=np.float64)
        for code in ("G", "I"):
            mask = codes == code
            if not mask.any():
                continue
            worst = np.nanmax(np.abs(identity[mask]))
            self.assertLess(worst, 1e-6, f"{code} chain did not tie: {worst}")

    def test_synthetic_swaps_contribute_no_pnl(self):
        swaps = self.result.swaps
        plain = np.asarray(swaps["is_plain"], dtype=bool)
        synthetic = ~plain
        self.assertTrue(synthetic.any(), "sample book should contain synthetics")
        self.assertTrue(np.all(np.isnan(np.asarray(swaps["practical_pnl"])[synthetic])))
        self.assertTrue(np.all(np.isnan(np.asarray(swaps["theoretical_pnl"])[synthetic])))

    def test_unlinked_hedges_are_measured_not_hidden(self):
        q = self.summary.quality
        self.assertGreater(q["unlinked_futures_count"], 0)
        self.assertGreater(q["unlinked_swap_count"], 0)
        self.assertNotEqual(q["unlinked_futures_pnl"], 0.0)

    def test_an_unclassified_floating_leg_suppresses_the_model_leg(self):
        frame = self.result.bonds
        flagged = np.asarray(frame["plain_swap_dv01_unknown"], dtype=np.float64) != 0.0
        self.assertTrue(flagged.any(), "sample book should contain one")
        self.assertTrue(np.all(np.isnan(np.asarray(frame["model_swap_pnl"])[flagged])))

    def test_the_framework_choice_is_driven_by_the_hedge_mix(self):
        frame = self.result.bonds
        codes = np.asarray(frame["spread_framework"], dtype=object)
        futures = np.abs(np.asarray(frame["futures_dv01"], dtype=np.float64))
        swaps = np.abs(np.asarray(frame["plain_swap_dv01"], dtype=np.float64))
        overridden = np.array(
            ["override" in str(r) for r in np.asarray(frame["spread_framework_reason"])]
        )
        futures_led = (futures > 0) & (futures >= swaps) & ~overridden
        self.assertTrue(np.all(codes[futures_led] == "G"))

    def test_the_run_is_deterministic(self):
        again = AttributionEngine(self.config, load_inputs(self.config)).run()
        for column in ("practical_pnl", "theoretical_pnl", "residual_pnl"):
            np.testing.assert_allclose(
                np.asarray(self.result.bonds[column], dtype=np.float64),
                np.asarray(again.bonds[column], dtype=np.float64),
                rtol=0, atol=0, equal_nan=True,
            )


if __name__ == "__main__":
    unittest.main()
