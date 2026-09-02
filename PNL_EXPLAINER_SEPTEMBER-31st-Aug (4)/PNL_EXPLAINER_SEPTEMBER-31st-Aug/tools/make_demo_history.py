"""Run the sample book over a run of days, so the snail trail has a trail.

A snail trail with one point is a dot.  The methodology only means anything
across a series of days, so this builds one: it takes the sample book, walks
the curves forward with a small seeded random walk, re-prices every bond off
the moved curves on each day, runs the real engine, and appends each day to the
history file the trail reads back.

Everything here is the ACTUAL pipeline - the same engine, the same aggregation,
the same history writer.  Nothing is faked into the history file; it is filled
by running the model, which is the only way the resulting trail says anything
true about the model.

The walk is seeded, so the demo is reproducible.

    python tools/make_demo_history.py                # 25 business days
    python tools/make_demo_history.py --days 60
    python tools/make_demo_history.py --reset        # start the history over
"""

from __future__ import annotations

import argparse
import copy
import datetime as _dt
import sys
from dataclasses import replace
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pnlx import bondmath  # noqa: E402
from pnlx.aggregate import build_portfolio  # noqa: E402
from pnlx.config import load_config  # noqa: E402
from pnlx.curves import CurveKey, CurveSet, ZeroCurve  # noqa: E402
from pnlx.engine import AttributionEngine  # noqa: E402
from pnlx.instruments import BondBook, FutureBook, SwapBook  # noqa: E402
from pnlx.loaders import InputBundle, load_inputs  # noqa: E402
from pnlx.report.csv_out import append_history  # noqa: E402

_DAY = np.dtype("datetime64[D]")


def business_days(start: _dt.date, count: int) -> list[_dt.date]:
    """`count` consecutive weekdays, starting the day after `start`."""
    out: list[_dt.date] = []
    day = start
    while len(out) < count:
        day += _dt.timedelta(days=1)
        if day.weekday() < 5:
            out.append(day)
    return out


def shift_curves(curves: CurveSet, shifts: dict[str, float], twist: float) -> CurveSet:
    """Move every curve by a parallel shift plus a small twist.

    Parallel plus twist rather than pure parallel, because a book that is only
    ever hit by parallel moves never exercises the curve legs against each
    other - the government/OIS and swap/government bases would stay frozen and
    the trail would be flatteringly clean.
    """
    moved = CurveSet()
    for curve in curves:
        key = curve.key
        shift = shifts.get(key.curve_type, 0.0)
        # Twist pivots around the 5-year point, in percentage points.
        slope = twist * (np.log1p(curve.tenors) - np.log1p(5.0)) / 10.0
        moved.add(ZeroCurve.build(key, curve.tenors, curve.rates + shift + slope))
    return moved


def reprice(book: BondBook, curves: CurveSet, prior: _dt.date, current: _dt.date,
            spread_drift: np.ndarray, max_flows: int) -> BondBook:
    """Roll the bond book one day forward on the moved curves.

    Each bond's yield is rebuilt as `base curve at its tenor + its own spread`,
    so a yield move really is the curve move plus the spread move and the
    attribution has something coherent to explain.  Prices come from the same
    `bondmath` the engine uses.
    """
    n = len(book)
    prior_d = np.full(n, np.datetime64(prior, "D"), dtype=_DAY)
    current_d = np.full(n, np.datetime64(current, "D"), dtype=_DAY)

    years_prior = book.years_to_maturity(np.datetime64(prior, "D"))
    years_current = book.years_to_maturity(np.datetime64(current, "D"))

    # Spread implied by yesterday's close: whatever the yield was, less the
    # government curve at the bond's tenor.
    gov_prior = curves.rates_for(book.currency, years_prior, "GOV", "prior")
    spread_prior_bp = (np.asarray(book.ytm_current, dtype=np.float64) - gov_prior) * 100.0
    spread_current_bp = spread_prior_bp + spread_drift

    gov_current = curves.rates_for(book.currency, years_current, "GOV", "current")
    ytm_current = gov_current + spread_current_bp / 100.0

    coupon = book.coupon_decimal
    freq = np.asarray(book.coupon_freq, dtype=np.float64)
    code = np.asarray(book.day_count_code, dtype=np.float64)

    clean_p, dirty_p = bondmath.bond_price_from_yield(
        prior_d, book.maturity, coupon,
        np.asarray(book.ytm_current, dtype=np.float64) / 100.0,
        freq, code, max_flows=max_flows,
    )
    clean_c, dirty_c = bondmath.bond_price_from_yield(
        current_d, book.maturity, coupon, ytm_current / 100.0, freq, code, max_flows=max_flows,
    )
    dur_p, _ = bondmath.yield_risk(
        prior_d, book.maturity, coupon,
        np.asarray(book.ytm_current, dtype=np.float64) / 100.0,
        freq, code, max_flows=max_flows,
    )
    dur_c, convexity = bondmath.yield_risk(
        current_d, book.maturity, coupon, ytm_current / 100.0, freq, code, max_flows=max_flows,
    )

    swap_prior = curves.rates_for(book.currency, years_prior, "SWAP", "prior")
    swap_current = curves.rates_for(book.currency, years_current, "SWAP", "current")

    rolled = copy.copy(book)
    rolled.ytm_prior = np.asarray(book.ytm_current, dtype=np.float64).copy()
    rolled.ytm_current = ytm_current
    rolled.clean_px_prior, rolled.dirty_px_prior = clean_p, dirty_p
    rolled.clean_px_current, rolled.dirty_px_current = clean_c, dirty_c
    rolled.mod_duration_prior, rolled.mod_duration_current = dur_p, dur_c
    rolled.convexity = convexity
    # Z / ASW / OAS keep the small, stable offsets from the I-spread that make
    # their chains tie approximately rather than exactly - which is the point.
    i_prior = (np.asarray(rolled.ytm_prior, float) - swap_prior) * 100.0
    i_current = (ytm_current - swap_current) * 100.0
    rolled.zspread_prior, rolled.zspread_current = i_prior + 3.0, i_current + 3.0
    rolled.asw_prior, rolled.asw_current = i_prior - 2.5, i_current - 2.5
    rolled.oas_prior, rolled.oas_current = i_prior - 6.0, i_current - 6.0
    return rolled


def roll_futures(book: FutureBook, bump_bp: np.ndarray) -> FutureBook:
    """Move each contract by its own DV01 against the government move."""
    if len(book) == 0:
        return book
    rolled = copy.copy(book)
    rolled.fut_px_prior = np.asarray(book.fut_px_current, dtype=np.float64).copy()
    # Price points per bp, so the price move is the rate move times that.
    move = -np.asarray(book.fut_px_val_bp, dtype=np.float64) * bump_bp
    rolled.fut_px_current = np.asarray(book.fut_px_current, dtype=np.float64) + move
    return rolled


def roll_swaps(book: SwapBook, ois_bp: float, swap_bp: float) -> SwapBook:
    """Mark each deal against the curve its floating leg projects off."""
    if len(book) == 0:
        return book
    rolled = copy.copy(book)
    family = np.array([str(f) for f in book.float_curve_type.tolist()], dtype=object)
    move = np.where(
        np.isin(family, np.array(["ESTR", "SOFR", "SONIA"], dtype=object)), ois_bp,
        np.where(family == "EURIBOR", swap_bp, np.nan),
    )
    dv01 = np.asarray(book.dv01_supplied, dtype=np.float64)
    rolled.npv_prior = np.zeros(len(book))
    # A small spread mismatch on top, so the swap basis line stays non-zero.
    rolled.npv_current = -dv01 * move + np.abs(np.asarray(book.notional, float)) * 1e-7
    synthetic = book.is_synthetic
    rolled.npv_prior = np.where(synthetic, np.nan, rolled.npv_prior)
    rolled.npv_current = np.where(synthetic, np.nan, rolled.npv_current)
    return rolled


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=25, help="business days to generate")
    parser.add_argument("--seed", type=int, default=20250930)
    parser.add_argument("--config", type=Path, default=None)
    parser.add_argument("--reset", action="store_true",
                        help="delete the existing history before starting")
    args = parser.parse_args()

    config = load_config(args.config)
    base = load_inputs(config)

    history_path = config.paths.output_dir / config.paths.history
    if args.reset and history_path.is_file():
        history_path.unlink()
        print(f"cleared {history_path}")

    rng = np.random.default_rng(args.seed)
    curves = base.curves
    bonds = base.bonds
    futures = base.futures
    swaps = base.swaps

    prior = config.run.as_of
    max_flows = config.model.pull_to_par.max_cashflows

    print(f"generating {args.days} business days from {prior}")
    for day in business_days(prior, args.days):
        # A day's market: a shared rates move, a small basis move on top of it,
        # a twist, and independent credit drift per bond.
        rates_bp = float(rng.normal(0.0, 2.6))
        gov_extra = float(rng.normal(0.0, 0.9))
        swap_extra = float(rng.normal(0.0, 0.7))
        twist = float(rng.normal(0.0, 0.9))
        spread_drift = rng.normal(0.0, 1.4, size=len(bonds))

        shifts = {
            "OIS": rates_bp / 100.0,
            "GOV": (rates_bp + gov_extra) / 100.0,
            "SWAP": (rates_bp + swap_extra) / 100.0,
        }

        # Yesterday's "current" curve becomes today's "prior".
        carried = CurveSet()
        for curve in curves:
            if curve.key.snapshot == "current":
                carried.add(
                    ZeroCurve.build(
                        CurveKey(curve.key.currency, curve.key.curve_type, "prior"),
                        curve.tenors, curve.rates,
                    )
                )
        # Snapshot the prior curves before adding to the set they came from.
        for curve in list(carried.curves.values()):
            key = CurveKey(curve.key.currency, curve.key.curve_type, "current")
            shift = shifts.get(key.curve_type, 0.0)
            slope = twist * (np.log1p(curve.tenors) - np.log1p(5.0)) / 100.0
            carried.add(ZeroCurve.build(key, curve.tenors, curve.rates + shift + slope))
        curves = carried

        bonds = reprice(bonds, curves, prior, day, spread_drift, max_flows)
        futures = roll_futures(futures, (rates_bp + gov_extra) / 1.0)
        swaps = roll_swaps(swaps, rates_bp, rates_bp + swap_extra)

        day_config = replace(config, run=replace(config.run, prior=prior, as_of=day))
        bundle = InputBundle(
            bonds=bonds, futures=futures, swaps=swaps, curves=curves,
            fx=base.fx, override_codes=base.override_codes,
            override_notes=base.override_notes,
        )
        result = AttributionEngine(day_config, bundle).run()
        summary = build_portfolio(result)
        append_history(result, summary)

        print(
            f"  {day}  practical {summary.headline['practical_pnl']:>12,.0f}"
            f"   explained {summary.headline['total_explained']:>12,.0f}"
            f"   residual {summary.headline['residual_pnl']:>10,.0f}"
        )
        prior = day

    print(f"\nhistory written to {history_path}")
    print("re-run `python -m pnlx` to rebuild the workbook with the full trail")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
