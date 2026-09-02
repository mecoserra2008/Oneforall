"""Day counts, coupon schedules and accrued interest.

The cases that matter here are the month-end ones.  A schedule built by
stepping period by period from the previous coupon compounds the day clamp -
31 March back six months is 30 September, and six months on from THAT is
30 March - which silently moves a 31st bond onto the 30th and shifts every
accrual after it by a day.  Several of these tests exist only to catch that.
"""

from __future__ import annotations

import unittest

import numpy as np

from pnlx import daycount as dc


class TestDateArithmetic(unittest.TestCase):
    def test_add_months_clamps_to_month_end(self):
        dates = dc.as_date_array(["2025-01-31", "2025-03-31", "2024-01-31"])
        got = dc.add_months(dates, 1)
        self.assertEqual(str(got[0]), "2025-02-28")
        self.assertEqual(str(got[1]), "2025-04-30")
        self.assertEqual(str(got[2]), "2024-02-29")  # leap year

    def test_add_months_is_reversible_from_a_fixed_anchor(self):
        anchor = dc.as_date_array(["2027-03-31"] * 3)
        for n in (6, 12, 18):
            back = dc.add_months(anchor, -n)
            forward = dc.add_months(back, n)
            # Stepping back and forward from an anchor is NOT reversible once
            # the day has been clamped - which is exactly why every date the
            # schedule produces is one add from maturity, never a chain.
            self.assertTrue(np.all(forward <= anchor))

    def test_days_in_month_and_leap_years(self):
        years = np.array([2024, 2025, 2000, 1900])
        self.assertEqual(list(dc.is_leap(years)), [True, False, True, False])
        self.assertEqual(list(dc.days_in_month(years, np.array([2, 2, 2, 2]))), [29, 28, 29, 28])


class TestCouponDates(unittest.TestCase):
    def test_month_end_semiannual_schedule_does_not_drift(self):
        settle = dc.as_date_array(["2025-09-30"])
        maturity = dc.as_date_array(["2027-03-31"])
        freq = np.array([2.0])

        self.assertEqual(str(dc.prev_coupon_date(settle, maturity, freq)[0]), "2025-09-30")
        # The one a naive step-by-step walk gets wrong: it yields 30 March.
        self.assertEqual(str(dc.next_coupon_date(settle, maturity, freq)[0]), "2026-03-31")

        dates, amounts, valid = dc.coupon_schedule(
            settle, maturity, freq, np.array([0.045])
        )
        self.assertEqual(
            [str(d) for d in dates[0][valid[0]]],
            ["2026-03-31", "2026-09-30", "2027-03-31"],
        )

    def test_month_end_quarterly_schedule_stays_on_the_31st(self):
        settle = dc.as_date_array(["2025-09-30"])
        maturity = dc.as_date_array(["2035-01-31"])
        dates, _, valid = dc.coupon_schedule(
            settle, maturity, np.array([4.0]), np.array([0.0125])
        )
        got = [str(d) for d in dates[0][valid[0]][:5]]
        self.assertEqual(
            got, ["2025-10-31", "2026-01-31", "2026-04-30", "2026-07-31", "2026-10-31"]
        )

    def test_last_flow_is_always_the_maturity_date(self):
        settle = dc.as_date_array(["2025-09-30"] * 3)
        maturity = dc.as_date_array(["2030-06-15", "2027-03-31", "2035-01-31"])
        freq = np.array([1.0, 2.0, 4.0])
        dates, amounts, valid = dc.coupon_schedule(
            settle, maturity, freq, np.array([0.03, 0.045, 0.0125])
        )
        for i in range(3):
            self.assertEqual(dates[i][valid[i]][-1], maturity[i])
            # Redemption rides on the final flow.
            self.assertGreater(amounts[i][valid[i]][-1], 100.0)

    def test_settlement_inside_the_final_period_gives_redemption_only(self):
        settle = dc.as_date_array(["2025-09-30"])
        maturity = dc.as_date_array(["2025-11-15"])
        dates, amounts, valid = dc.coupon_schedule(
            settle, maturity, np.array([2.0]), np.array([0.02])
        )
        self.assertEqual(int(valid.sum()), 1)
        self.assertEqual(dates[0][0], maturity[0])

    def test_past_maturity_has_no_flows(self):
        _, _, valid = dc.coupon_schedule(
            dc.as_date_array(["2026-01-01"]),
            dc.as_date_array(["2025-11-15"]),
            np.array([2.0]),
            np.array([0.02]),
        )
        self.assertEqual(int(valid.sum()), 0)


class TestAccrued(unittest.TestCase):
    def test_zero_on_a_coupon_date(self):
        settle = dc.as_date_array(["2025-06-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        got = dc.accrued_interest(settle, maturity, np.array([0.04]), np.array([1.0]), np.array([0.0]))
        self.assertAlmostEqual(float(got[0]), 0.0, places=12)

    def test_half_a_period_is_half_a_coupon_on_icma(self):
        # 15 Dec is exactly half way between 15 Jun and 15 Jun (annual, 183/365).
        settle = dc.as_date_array(["2025-12-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        got = float(
            dc.accrued_interest(settle, maturity, np.array([0.04]), np.array([1.0]), np.array([0.0]))[0]
        )
        self.assertAlmostEqual(got, 4.0 * 183 / 365, places=6)

    def test_thirty_360_counts_whole_months(self):
        settle = dc.as_date_array(["2025-12-15"])
        maturity = dc.as_date_array(["2030-06-15"])
        got = float(
            dc.accrued_interest(
                settle, maturity, np.array([0.04]), np.array([1.0]),
                np.array([float(dc.DCC_30E_360)]),
            )[0]
        )
        self.assertAlmostEqual(got, 4.0 * 180 / 360, places=9)

    def test_unknown_convention_does_not_silently_default(self):
        got = dc.day_count_fraction(
            dc.as_date_array(["2025-01-01"]),
            dc.as_date_array(["2026-01-01"]),
            np.array([99.0]),
        )
        self.assertTrue(np.isnan(got[0]))


class TestCouponsPaid(unittest.TestCase):
    def test_a_coupon_inside_the_window_is_counted_once(self):
        got = dc.coupons_paid_between(
            dc.as_date_array(["2025-06-14"]),
            dc.as_date_array(["2025-06-16"]),
            dc.as_date_array(["2030-06-15"]),
            np.array([0.03]),
            np.array([1.0]),
        )
        self.assertAlmostEqual(float(got[0]), 3.0, places=9)

    def test_no_coupon_in_the_window_pays_nothing(self):
        got = dc.coupons_paid_between(
            dc.as_date_array(["2025-06-14"]),
            dc.as_date_array(["2025-06-16"]),
            dc.as_date_array(["2030-12-15"]),
            np.array([0.03]),
            np.array([1.0]),
        )
        self.assertAlmostEqual(float(got[0]), 0.0, places=12)

    def test_redemption_is_not_carry(self):
        # A bond maturing inside the window is a position event, not coupon
        # income, so its 100 must not appear here.
        got = dc.coupons_paid_between(
            dc.as_date_array(["2025-06-14"]),
            dc.as_date_array(["2025-06-16"]),
            dc.as_date_array(["2025-06-15"]),
            np.array([0.03]),
            np.array([1.0]),
        )
        self.assertLess(float(got[0]), 100.0)


class TestDescriptionMappers(unittest.TestCase):
    def test_the_real_feed_strings(self):
        cases = {
            "ACT/ACT": dc.DCC_ACT_ACT_ICMA,
            "ISDA ACT/ACT": dc.DCC_ACT_ACT_ISDA,
            "ACT/ACT NON-EOM": dc.DCC_ACT_ACT_ICMA,
            "ACT/360(102)": dc.DCC_ACT_360,
            "30/360(104)": dc.DCC_30_360_BOND,
            "ISMA-30/360": dc.DCC_30E_360,
            "ISMA-30/360 NONEOM": dc.DCC_30E_360,
            "ISDA SWAPS:30/360": dc.DCC_30_360_BOND,
            "ACT/365": dc.DCC_ACT_365F,
        }
        for text, expected in cases.items():
            self.assertEqual(dc.day_count_from_description(text), float(expected), text)

    def test_unknown_description_is_nan_not_a_guess(self):
        for text in ("", "  ", "WIBBLE/360", None):
            self.assertTrue(np.isnan(dc.day_count_from_description(text)), repr(text))

    def test_frequency_words_and_numbers(self):
        for text, expected in (
            ("SEMI", 2.0), ("A", 1.0), ("Q", 4.0), (2, 2.0), ("4", 4.0), ("12M", 1.0),
        ):
            self.assertEqual(dc.coupon_frequency_from_description(text), expected, text)
        for text in ("", "MONTHLY-ISH", 7):
            self.assertTrue(np.isnan(dc.coupon_frequency_from_description(text)), repr(text))


if __name__ == "__main__":
    unittest.main()
