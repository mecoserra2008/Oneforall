"""Bond pricing, risk and pull to par - all vectorised over a whole book.

Three jobs live here:

    1. price / yield        the street formula, so a price can be reproduced
                            from a yield and vice versa
    2. duration / convexity  struck by reprice, at a configurable bump
    3. pull to par           the passage-of-time part of a clean price change

Everything takes and returns arrays.  A book of several hundred bonds with a
hundred cash flows each is one pass of matrix arithmetic; there is no per-bond
Python loop in any of it.

-------------------------------------------------------------------------------
PULL TO PAR - what it measures and why it is done this way
-------------------------------------------------------------------------------
The part of a bond's CLEAN price change from T-1 to T0 that is pure passage of
time, with the curve and the credit spread both held at their T-1 levels.

The naive version - reprice on the SAME curve at a shorter maturity - books the
entire roll-down of an upward-sloping curve as PnL the desk never earned.  The
arbitrage-free answer discounts each remaining cash flow at the FORWARD rate the
prior curve implies between the T0 horizon `h` and that flow at `t`:

    (1 + f(h,t)) ** (t - h)  =  (1 + y(t)) ** t / (1 + y(h)) ** h

so the forward value at the horizon of a cash flow C_t is

    C_t / (1 + f(h,t)) ** (t-h)  =  C_t * (1 + y(h)) ** h / (1 + y(t)) ** t

and summing over the flows still outstanding at the horizon gives

    DirtyPx_fwd = (1 + y(h)) ** h * SUM_{t > h} C_t / (1 + y(t)) ** t

which is the classic "spot price less the PV of interim coupons, compounded to
the horizon".  Coupons paid on or before the horizon drop out on purpose: that
cash is reported by the coupon-carry leg, and counting it here as well is the
single easiest way to double the carry.

The result is a CLEAN price change per 100 nominal, accrued stripped at both
ends so coupon accrual cannot leak in:

    (DirtyPx_fwd - AI(T0)) - (DirtyPx_spot - AI(T-1))

Both legs are MODEL prices, so neither ties exactly to a quoted price.  Their
DIFFERENCE is what is used, and the calibration error is common to both terms
and cancels to first order.  Neither leg is a price on its own.

The time axis is ACT/365F, NOT the bond's own day count: the curve nodes are
quoted in decimal years against annually compounded zero rates, so the horizon
and every cash flow must be measured on the same simple axis for the forward
identity above to hold.  The bond's day count is still used where it belongs -
accrued interest.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from . import daycount as dc
from .curves import CurveSet

__all__ = [
    "PullToParResult",
    "bond_price_from_yield",
    "yield_risk",
    "pull_to_par",
    "model_dirty_price",
]

_DAY = np.dtype("datetime64[D]")


# --------------------------------------------------------------------------- #
# price / yield
# --------------------------------------------------------------------------- #


def bond_price_from_yield(
    settlement: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    ytm: np.ndarray,
    frequency: np.ndarray,
    dcc: np.ndarray,
    redemption: float = 100.0,
    max_flows: int = 1200,
) -> tuple[np.ndarray, np.ndarray]:
    """Clean and dirty price per 100 nominal from a yield to maturity.

    The street / ICMA formula, and the same one Excel's PRICE() implements:
    each flow is discounted at ``(1 + y/f) ** (k - 1 + DSC/E)`` where `DSC/E` is
    the unexpired fraction of the current coupon period measured on the bond's
    own day count.  That fractional first exponent is what makes the function
    agree with a quoted yield between coupon dates instead of only on them.

    `coupon_rate` and `ytm` are DECIMAL.  Returns `(clean, dirty)`.
    """
    settlement = np.asarray(settlement, dtype=_DAY)
    maturity = np.asarray(maturity, dtype=_DAY)
    rate = np.asarray(coupon_rate, dtype=np.float64)
    y = np.asarray(ytm, dtype=np.float64)
    freq = np.asarray(frequency, dtype=np.float64)
    code = np.asarray(dcc, dtype=np.float64)

    dates, amounts, valid = dc.coupon_schedule(
        settlement, maturity, freq, rate, max_flows=max_flows
    )
    n, width = dates.shape

    # Redemption rides on the last flow; rebuild it if it is not 100.
    if redemption != 100.0:
        counts = valid.sum(axis=1)
        last = np.maximum(counts - 1, 0)
        rows = np.arange(n)
        amounts = amounts.copy()
        amounts[rows, last] = np.where(
            counts > 0, amounts[rows, last] - 100.0 + redemption, amounts[rows, last]
        )

    prev = dc.prev_coupon_date(settlement, maturity, freq)
    nxt = dc.next_coupon_date(settlement, maturity, freq)

    accrued_frac = dc.accrued_fraction(settlement, prev, nxt, freq, code, maturity)
    accrued_frac = np.clip(np.nan_to_num(accrued_frac, nan=0.0), 0.0, 1.0)
    unexpired = 1.0 - accrued_frac                       # DSC / E

    # Exponent of flow k (1-based): k - 1 + DSC/E, in coupon periods.
    k = np.arange(1, width + 1, dtype=np.float64)[None, :]
    exponent = (k - 1.0) + unexpired[:, None]

    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        per_period = 1.0 + y[:, None] / freq[:, None]
        discounted = np.where(
            valid & (per_period > 0),
            amounts / per_period**exponent,
            0.0,
        )
        dirty = discounted.sum(axis=1)

    accrued = (rate / freq) * accrued_frac * 100.0
    clean = dirty - accrued

    usable = (
        valid.any(axis=1)
        & np.isfinite(y)
        & np.isfinite(rate)
        & np.isfinite(freq)
        & (freq > 0)
        & (settlement < maturity)
    )
    clean = np.where(usable, clean, np.nan)
    dirty = np.where(usable, dirty, np.nan)
    return clean, dirty


def yield_risk(
    settlement: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    ytm: np.ndarray,
    frequency: np.ndarray,
    dcc: np.ndarray,
    bump_bp: float = 100.0,
    max_flows: int = 1200,
) -> tuple[np.ndarray, np.ndarray]:
    """Modified duration and convexity by central-difference reprice.

    Returns `(mod_duration, convexity)`, both against the DIRTY price - which
    is the price the position is actually marked at, so a duration struck
    against the clean price would under-state the risk of a high-coupon bond by
    the accrued fraction of its value.

        ModDur    = -(P_up - P_down) / (2 * P_0 * dy)
        Convexity =  (P_up + P_down - 2 * P_0) / (P_0 * dy ** 2)

    The bump is configurable for one reason: at a large bump these are the
    average duration and convexity ACROSS the bump, which is what a desk wants
    when the day's move was large; at a vanishing bump they converge on the
    analytic derivatives.  Both are reported so the difference is visible rather
    than assumed away.  100bp reproduces the legacy workbook.
    """
    dy = bump_bp / 10_000.0
    y = np.asarray(ytm, dtype=np.float64)

    _, p0 = bond_price_from_yield(
        settlement, maturity, coupon_rate, y, frequency, dcc, max_flows=max_flows
    )
    _, p_up = bond_price_from_yield(
        settlement, maturity, coupon_rate, y + dy, frequency, dcc, max_flows=max_flows
    )
    _, p_dn = bond_price_from_yield(
        settlement, maturity, coupon_rate, y - dy, frequency, dcc, max_flows=max_flows
    )

    with np.errstate(divide="ignore", invalid="ignore"):
        ok = np.isfinite(p0) & (p0 != 0) & np.isfinite(p_up) & np.isfinite(p_dn)
        mod_dur = np.where(ok, -(p_up - p_dn) / (2.0 * p0 * dy), np.nan)
        convexity = np.where(ok, (p_up + p_dn - 2.0 * p0) / (p0 * dy * dy), np.nan)

    return mod_dur, convexity


# --------------------------------------------------------------------------- #
# pull to par
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class PullToParResult:
    """The pull-to-par calculation, with both legs kept for audit.

    `clean_change` is the deliverable; the two dirty legs are exposed because
    the only honest way to check the number is to look at the prices it is a
    difference of.  Neither leg is a usable price on its own - see the module
    docstring.
    """

    clean_change: np.ndarray
    dirty_spot: np.ndarray
    dirty_forward: np.ndarray
    accrued_prior: np.ndarray
    accrued_current: np.ndarray
    horizon_years: np.ndarray
    usable: np.ndarray

    @property
    def n(self) -> int:
        return int(self.clean_change.size)


def _zero_rates_on_matrix(
    curve_set: CurveSet,
    currencies: np.ndarray,
    curve_types: np.ndarray,
    years: np.ndarray,
    snapshot: str,
) -> np.ndarray:
    """Interpolate a (n_bonds, n_flows) matrix of tenors, in PERCENT.

    Grouped by (currency, curve type) so each distinct curve is touched once,
    however many cash flows reference it.
    """
    out = np.full(years.shape, np.nan, dtype=np.float64)
    ccy = np.asarray(currencies, dtype=object)
    ctype = np.asarray(curve_types, dtype=object)

    pairs = {
        (str(c).strip().upper(), str(t).strip().upper())
        for c, t in zip(ccy.tolist(), ctype.tolist())
        if c and t
    }
    for code, kind in pairs:
        curve = curve_set.get(code, kind, snapshot)
        if curve is None:
            continue
        rows = np.array(
            [
                str(c).strip().upper() == code and str(t).strip().upper() == kind
                for c, t in zip(ccy.tolist(), ctype.tolist())
            ]
        )
        if not rows.any():
            continue
        out[rows, :] = curve.rate(years[rows, :])
    return out


def pull_to_par(
    prior: np.ndarray,
    current: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    frequency: np.ndarray,
    dcc: np.ndarray,
    currency: np.ndarray,
    curve_type: np.ndarray,
    spread_bp: np.ndarray,
    curve_set: CurveSet,
    day_basis: float = 365.0,
    max_flows: int = 1200,
) -> PullToParResult:
    """Clean price change per 100 nominal from pure passage of time.

    `curve_type` is the base curve the bond's spread is quoted over, chosen by
    its spread framework (see `frameworks.curve_for_framework`).  `spread_bp` is
    the T-1 spread on that framework, applied flat across tenors - the same
    calibration the rest of the model uses.

    Every discount factor comes off the PRIOR curve.  That is the definition:
    holding the curve still is what makes the answer time and nothing else.
    """
    prior = np.asarray(prior, dtype=_DAY)
    current = np.asarray(current, dtype=_DAY)
    maturity = np.asarray(maturity, dtype=_DAY)
    rate = np.asarray(coupon_rate, dtype=np.float64)
    freq = np.asarray(frequency, dtype=np.float64)
    code = np.asarray(dcc, dtype=np.float64)
    spread = np.asarray(spread_bp, dtype=np.float64) / 10_000.0

    n = prior.size
    nan = np.full(n, np.nan)

    horizon = (
        (current.astype("int64") - prior.astype("int64")).astype(np.float64) / day_basis
    )

    dates, amounts, valid = dc.coupon_schedule(
        prior, maturity, freq, rate, max_flows=max_flows
    )

    # Time to each flow, measured from T-1 on the same ACT/365F axis.
    flow_years = (
        dates.astype("datetime64[D]").astype("int64") - prior.astype("int64")[:, None]
    ).astype(np.float64) / day_basis
    flow_years = np.where(valid, flow_years, np.nan)

    zero_pct = _zero_rates_on_matrix(curve_set, currency, curve_type, np.nan_to_num(flow_years), "prior")
    zero = zero_pct / 100.0 + spread[:, None]

    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        base = 1.0 + zero
        priceable = valid & (flow_years > 0) & np.isfinite(zero) & (base > 0)
        pv_each = np.where(priceable, amounts / base**flow_years, 0.0)

        dirty_spot = pv_each.sum(axis=1)

        # Only what is still outstanding at the horizon gets compounded forward.
        # Anything paid on or before it is cash in the bank, already reported by
        # the coupon-carry leg.
        outstanding = priceable & (flow_years > horizon[:, None])
        remaining_pv = np.where(outstanding, pv_each, 0.0).sum(axis=1)

    horizon_rate_pct = np.full(n, np.nan)
    ccy = np.asarray(currency, dtype=object)
    ctype = np.asarray(curve_type, dtype=object)
    for kcode, kkind in {
        (str(c).strip().upper(), str(t).strip().upper())
        for c, t in zip(ccy.tolist(), ctype.tolist())
        if c and t
    }:
        curve = curve_set.get(kcode, kkind, "prior")
        if curve is None:
            continue
        rows = np.array(
            [
                str(c).strip().upper() == kcode and str(t).strip().upper() == kkind
                for c, t in zip(ccy.tolist(), ctype.tolist())
            ]
        )
        horizon_rate_pct[rows] = curve.rate(horizon[rows])

    spot_to_horizon = horizon_rate_pct / 100.0 + spread

    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        growth_base = 1.0 + spot_to_horizon
        growth = np.where(growth_base > 0, growth_base**horizon, np.nan)
        dirty_forward = remaining_pv * growth

    accrued_prior = dc.accrued_interest(prior, maturity, rate, freq, code)
    accrued_current = dc.accrued_interest(current, maturity, rate, freq, code)

    usable = (
        (current > prior)
        & (prior < maturity)
        & (current < maturity)
        & np.isin(freq, np.array(dc.VALID_FREQUENCIES, dtype=np.float64))
        & np.isfinite(code)
        & (code >= 0)
        & (code <= 6)
        & np.isfinite(spread)
        & np.isfinite(dirty_spot)
        & (dirty_spot > 0)
        & np.isfinite(dirty_forward)
        & valid.any(axis=1)
    )

    clean_change = np.where(
        usable,
        (dirty_forward - accrued_current) - (dirty_spot - accrued_prior),
        np.nan,
    )

    return PullToParResult(
        clean_change=clean_change,
        dirty_spot=np.where(usable, dirty_spot, nan),
        dirty_forward=np.where(usable, dirty_forward, nan),
        accrued_prior=accrued_prior,
        accrued_current=accrued_current,
        horizon_years=horizon,
        usable=usable,
    )


def model_dirty_price(
    settlement: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    frequency: np.ndarray,
    currency: np.ndarray,
    curve_type: np.ndarray,
    spread_bp: np.ndarray,
    curve_set: CurveSet,
    snapshot: str = "prior",
    day_basis: float = 365.0,
    max_flows: int = 1200,
) -> np.ndarray:
    """Dirty price per 100 nominal on a zero curve plus a flat spread.

    Exposed because it is the natural check on the pull-to-par result: when no
    coupon falls inside the horizon, the forward dirty price must equal this
    price grown at the horizon zero rate.  Also usable straight against a quoted
    dirty price to see how well the flat-spread calibration holds.
    """
    settlement = np.asarray(settlement, dtype=_DAY)
    freq = np.asarray(frequency, dtype=np.float64)
    rate = np.asarray(coupon_rate, dtype=np.float64)
    spread = np.asarray(spread_bp, dtype=np.float64) / 10_000.0

    dates, amounts, valid = dc.coupon_schedule(
        settlement, maturity, freq, rate, max_flows=max_flows
    )
    flow_years = (
        dates.astype("datetime64[D]").astype("int64") - settlement.astype("int64")[:, None]
    ).astype(np.float64) / day_basis
    flow_years = np.where(valid, flow_years, np.nan)

    zero_pct = _zero_rates_on_matrix(
        curve_set, currency, curve_type, np.nan_to_num(flow_years), snapshot
    )
    zero = zero_pct / 100.0 + spread[:, None]

    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        base = 1.0 + zero
        good = valid & (flow_years > 0) & np.isfinite(zero) & (base > 0)
        pv = np.where(good, amounts / base**flow_years, 0.0).sum(axis=1)

    return np.where(good.any(axis=1) & (pv > 0), pv, np.nan)
