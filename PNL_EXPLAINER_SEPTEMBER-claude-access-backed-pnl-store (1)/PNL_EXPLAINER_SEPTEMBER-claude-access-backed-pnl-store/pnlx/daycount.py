"""Day-count conventions, coupon schedules and accrued interest - vectorised.

This is the Python of `modImpRepo.bas`, with one structural difference: every
routine takes and returns numpy ARRAYS.  The VBA versions were scalar and were
called once per bond per cash flow from a worksheet formula, which is what made
a full recalculation of the book slow enough to be a workflow problem.  Here a
whole book of several hundred bonds is one pass of array arithmetic.

Convention codes are the same seven integers the workbook used, so a stored
`BondDCC_Code` still means what it meant:

    0  ACT/ACT ICMA
    1  ACT/ACT ISDA
    2  ACT/365F
    3  ACT/360
    4  30/360 Bond (US SIA / NASD)
    5  30E/360 Eurobond
    6  30E/360 ISDA

Dates are `numpy.datetime64[D]` throughout.  Day differences are then plain
integer subtraction, which is both exact and fast; there is no floating-point
date arithmetic anywhere in this module.
"""

from __future__ import annotations

import datetime as _dt
from typing import Final, Iterable, Sequence

import numpy as np

__all__ = [
    "DCC_ACT_ACT_ICMA",
    "DCC_ACT_ACT_ISDA",
    "DCC_ACT_365F",
    "DCC_ACT_360",
    "DCC_30_360_BOND",
    "DCC_30E_360",
    "DCC_30E_360_ISDA",
    "DCC_NAMES",
    "VALID_FREQUENCIES",
    "as_date_array",
    "to_datetime64",
    "ymd",
    "days_in_month",
    "add_months",
    "is_leap",
    "is_end_of_month",
    "day_count_fraction",
    "accrued_fraction",
    "accrued_interest",
    "prev_coupon_date",
    "next_coupon_date",
    "coupon_schedule",
    "coupons_paid_between",
    "day_count_from_description",
    "coupon_frequency_from_description",
    "excel_price_basis",
]

# --------------------------------------------------------------------------- #
# convention codes
# --------------------------------------------------------------------------- #

DCC_ACT_ACT_ICMA: Final[int] = 0
DCC_ACT_ACT_ISDA: Final[int] = 1
DCC_ACT_365F: Final[int] = 2
DCC_ACT_360: Final[int] = 3
DCC_30_360_BOND: Final[int] = 4
DCC_30E_360: Final[int] = 5
DCC_30E_360_ISDA: Final[int] = 6

DCC_NAMES: Final[dict[int, str]] = {
    DCC_ACT_ACT_ICMA: "ACT/ACT ICMA",
    DCC_ACT_ACT_ISDA: "ACT/ACT ISDA",
    DCC_ACT_365F: "ACT/365F",
    DCC_ACT_360: "ACT/360",
    DCC_30_360_BOND: "30/360 BOND",
    DCC_30E_360: "30E/360",
    DCC_30E_360_ISDA: "30E/360 ISDA",
}

VALID_FREQUENCIES: Final[tuple[int, ...]] = (1, 2, 4, 12)

_DAY = np.dtype("datetime64[D]")
_MONTH_LENGTHS = np.array([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31], dtype=np.int64)


# --------------------------------------------------------------------------- #
# date plumbing
# --------------------------------------------------------------------------- #


def to_datetime64(value) -> np.datetime64:
    """One date, as `datetime64[D]`.  Accepts date, datetime, string or None."""
    if value is None:
        return np.datetime64("NaT", "D")
    if isinstance(value, np.datetime64):
        return value.astype(_DAY)
    if isinstance(value, _dt.datetime):
        return np.datetime64(value.date(), "D")
    if isinstance(value, _dt.date):
        return np.datetime64(value, "D")
    return np.datetime64(str(value)[:10], "D")


def as_date_array(values: Iterable | np.ndarray) -> np.ndarray:
    """A sequence of dates, as a `datetime64[D]` array.  Blanks become NaT."""
    if isinstance(values, np.ndarray) and values.dtype.kind == "M":
        return values.astype(_DAY)
    return np.array([to_datetime64(v) for v in values], dtype=_DAY)


def ymd(dates: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Split `datetime64[D]` into (year, month, day) int64 arrays.

    Done through the calendar-unit casts rather than by conversion to Python
    `date` objects, so it stays a vector operation.
    """
    d = np.asarray(dates, dtype=_DAY)
    years = d.astype("datetime64[Y]").astype(np.int64) + 1970
    months = d.astype("datetime64[M]").astype(np.int64) % 12 + 1
    days = (d - d.astype("datetime64[M]")).astype(np.int64) + 1
    return years, months, days


def is_leap(years: np.ndarray) -> np.ndarray:
    y = np.asarray(years, dtype=np.int64)
    return ((y % 4 == 0) & (y % 100 != 0)) | (y % 400 == 0)


def days_in_month(years: np.ndarray, months: np.ndarray) -> np.ndarray:
    """Length of each (year, month), leap years included."""
    y = np.asarray(years, dtype=np.int64)
    m = np.asarray(months, dtype=np.int64)
    base = _MONTH_LENGTHS[m - 1]
    return np.where((m == 2) & is_leap(y), 29, base)


def is_end_of_month(dates: np.ndarray) -> np.ndarray:
    y, m, d = ymd(dates)
    return d == days_in_month(y, m)


def add_months(dates: np.ndarray, months: np.ndarray | int) -> np.ndarray:
    """Shift dates by whole months, clamping the day to the target month length.

    Matches VBA `DateAdd("m", n, d)`: 31 Jan + 1 month is 28/29 Feb, not 3 March.
    A schedule built by repeated month-adds therefore never walks off the
    month-end, which is exactly the drift the workbook guarded against by
    forcing the last cash flow onto the maturity date.
    """
    d = np.asarray(dates, dtype=_DAY)
    n = np.asarray(months, dtype=np.int64)
    y0, m0, day0 = ymd(d)

    idx = y0 * 12 + (m0 - 1) + n
    y1 = idx // 12
    m1 = idx % 12 + 1
    day1 = np.minimum(day0, days_in_month(y1, m1))

    out = (
        (y1 - 1970).astype("timedelta64[Y]").astype("datetime64[Y]")
        .astype("datetime64[M]")
        + (m1 - 1).astype("timedelta64[M]")
    ).astype(_DAY) + (day1 - 1).astype("timedelta64[D]")

    return np.where(np.isnat(d), np.datetime64("NaT", "D"), out)


def _days(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    """Calendar days from a to b, as float (NaN where either end is NaT)."""
    a = np.asarray(a, dtype=_DAY)
    b = np.asarray(b, dtype=_DAY)
    bad = np.isnat(a) | np.isnat(b)
    raw = (b.astype("int64") - a.astype("int64")).astype(np.float64)
    return np.where(bad, np.nan, raw)


# --------------------------------------------------------------------------- #
# the seven conventions
# --------------------------------------------------------------------------- #


def _days_30_360_bond(d1: np.ndarray, d2: np.ndarray) -> np.ndarray:
    """US SIA / NASD 30/360."""
    y1, m1, dd1 = ymd(d1)
    y2, m2, dd2 = ymd(d2)
    dd1 = np.where(dd1 == 31, 30, dd1)
    dd2 = np.where((dd2 == 31) & (dd1 == 30), 30, dd2)
    return 360.0 * (y2 - y1) + 30.0 * (m2 - m1) + (dd2 - dd1)


def _days_30e_360(d1: np.ndarray, d2: np.ndarray) -> np.ndarray:
    """Eurobond 30E/360: both ends capped at 30, unconditionally."""
    y1, m1, dd1 = ymd(d1)
    y2, m2, dd2 = ymd(d2)
    dd1 = np.where(dd1 == 31, 30, dd1)
    dd2 = np.where(dd2 == 31, 30, dd2)
    return 360.0 * (y2 - y1) + 30.0 * (m2 - m1) + (dd2 - dd1)


def _days_30e_360_isda(d1: np.ndarray, d2: np.ndarray, maturity: np.ndarray) -> np.ndarray:
    """30E/360 ISDA: month-ends go to 30, except a February maturity."""
    y1, m1, dd1 = ymd(d1)
    y2, m2, dd2 = ymd(d2)
    dd1 = np.where(is_end_of_month(d1), 30, dd1)
    end2 = is_end_of_month(d2)
    feb_maturity = (np.asarray(d2, dtype=_DAY) == np.asarray(maturity, dtype=_DAY)) & (m2 == 2)
    dd2 = np.where(end2 & ~feb_maturity, 30, dd2)
    return 360.0 * (y2 - y1) + 30.0 * (m2 - m1) + (dd2 - dd1)


def _act_act_isda_fraction(d1: np.ndarray, d2: np.ndarray) -> np.ndarray:
    """ACT/ACT ISDA: each calendar year contributes on its own basis."""
    d1 = np.asarray(d1, dtype=_DAY)
    d2 = np.asarray(d2, dtype=_DAY)
    y1, _, _ = ymd(d1)
    y2, _, _ = ymd(d2)

    jan1_next = (y1 + 1 - 1970).astype("timedelta64[Y]").astype("datetime64[Y]").astype(_DAY)
    jan1_this = (y2 - 1970).astype("timedelta64[Y]").astype("datetime64[Y]").astype(_DAY)

    basis1 = np.where(is_leap(y1), 366.0, 365.0)
    basis2 = np.where(is_leap(y2), 366.0, 365.0)

    same_year = _days(d1, d2) / basis1
    spanning = (
        _days(d1, jan1_next) / basis1
        + (y2 - y1 - 1).astype(np.float64)
        + _days(jan1_this, d2) / basis2
    )
    return np.where(y1 == y2, same_year, spanning)


def day_count_fraction(
    start: np.ndarray,
    end: np.ndarray,
    dcc: np.ndarray,
    frequency: np.ndarray | int = 2,
    maturity: np.ndarray | None = None,
) -> np.ndarray:
    """Year fraction between two dates under each row's convention.

    ACT/ACT ICMA needs the coupon period it sits in, so it needs `maturity` and
    `frequency`; the others ignore them.  Rows whose convention code is outside
    0..6 come back NaN rather than defaulting - an unknown convention must not
    quietly become ACT/365.
    """
    start = np.asarray(start, dtype=_DAY)
    end = np.asarray(end, dtype=_DAY)
    code = np.asarray(dcc, dtype=np.float64)
    freq = np.broadcast_to(np.asarray(frequency, dtype=np.float64), start.shape)

    out = np.full(start.shape, np.nan, dtype=np.float64)
    raw_days = _days(start, end)

    if maturity is None:
        maturity = end
    maturity = np.broadcast_to(np.asarray(maturity, dtype=_DAY), start.shape)

    m = code == DCC_ACT_ACT_ICMA
    if m.any():
        prev = prev_coupon_date(start, maturity, freq)
        nxt = next_coupon_date(start, maturity, freq)
        period = _days(prev, nxt)
        with np.errstate(divide="ignore", invalid="ignore"):
            icma = raw_days / (freq * period)
        out = np.where(m, icma, out)

    m = code == DCC_ACT_ACT_ISDA
    if m.any():
        out = np.where(m, _act_act_isda_fraction(start, end), out)

    out = np.where(code == DCC_ACT_365F, raw_days / 365.0, out)
    out = np.where(code == DCC_ACT_360, raw_days / 360.0, out)
    out = np.where(code == DCC_30_360_BOND, _days_30_360_bond(start, end) / 360.0, out)
    out = np.where(code == DCC_30E_360, _days_30e_360(start, end) / 360.0, out)
    out = np.where(
        code == DCC_30E_360_ISDA, _days_30e_360_isda(start, end, maturity) / 360.0, out
    )
    return out


def accrued_fraction(
    settlement: np.ndarray,
    prev_coupon: np.ndarray,
    next_coupon: np.ndarray,
    frequency: np.ndarray,
    dcc: np.ndarray,
    maturity: np.ndarray,
) -> np.ndarray:
    """Fraction of the current coupon period that has accrued at settlement.

    Expressed as a fraction of ONE COUPON PERIOD (so accrued interest is
    `coupon / frequency * fraction`), which is why every non-ICMA branch divides
    by `basis / frequency` rather than by `basis`.
    """
    settlement = np.asarray(settlement, dtype=_DAY)
    code = np.asarray(dcc, dtype=np.float64)
    freq = np.asarray(frequency, dtype=np.float64)

    elapsed = _days(prev_coupon, settlement)
    period = _days(prev_coupon, next_coupon)

    with np.errstate(divide="ignore", invalid="ignore"):
        icma = elapsed / period
        act365 = elapsed / (365.0 / freq)
        act360 = elapsed / (360.0 / freq)
        b30_360 = _days_30_360_bond(prev_coupon, settlement) / (360.0 / freq)
        b30e_360 = _days_30e_360(prev_coupon, settlement) / (360.0 / freq)
        b30e_isda = _days_30e_360_isda(prev_coupon, settlement, maturity) / (360.0 / freq)

    isda = _act_act_isda_fraction(prev_coupon, settlement) * freq

    # ICMA is also the fallback for an unrecognised code: it is the definition
    # that needs no basis at all, so it cannot mis-scale a coupon period.
    out = icma.copy()
    out = np.where(code == DCC_ACT_ACT_ISDA, isda, out)
    out = np.where(code == DCC_ACT_365F, act365, out)
    out = np.where(code == DCC_ACT_360, act360, out)
    out = np.where(code == DCC_30_360_BOND, b30_360, out)
    out = np.where(code == DCC_30E_360, b30e_360, out)
    out = np.where(code == DCC_30E_360_ISDA, b30e_isda, out)
    return out


def accrued_interest(
    settlement: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    frequency: np.ndarray,
    dcc: np.ndarray,
) -> np.ndarray:
    """Accrued interest per 100 nominal.  `coupon_rate` is DECIMAL (0.0375)."""
    settlement = np.asarray(settlement, dtype=_DAY)
    maturity = np.asarray(maturity, dtype=_DAY)
    freq = np.asarray(frequency, dtype=np.float64)
    rate = np.asarray(coupon_rate, dtype=np.float64)

    prev = prev_coupon_date(settlement, maturity, freq)
    nxt = next_coupon_date(settlement, maturity, freq)
    frac = accrued_fraction(settlement, prev, nxt, freq, dcc, maturity)

    with np.errstate(divide="ignore", invalid="ignore"):
        return (rate / freq) * frac * 100.0


# --------------------------------------------------------------------------- #
# coupon schedules
# --------------------------------------------------------------------------- #


def _months_per_period(frequency: np.ndarray) -> np.ndarray:
    freq = np.asarray(frequency, dtype=np.float64)
    with np.errstate(divide="ignore", invalid="ignore"):
        per = np.where(freq > 0, np.round(12.0 / freq), np.nan)
    return per


def _periods_back(
    reference: np.ndarray, maturity: np.ndarray, frequency: np.ndarray
) -> tuple[np.ndarray, np.ndarray]:
    """How many whole coupon periods before maturity the current period starts.

    Returns `(k, months_per_period)` where `k` is the smallest non-negative
    integer with ``maturity - k periods <= reference``.  Every coupon date the
    rest of this module produces is then ONE month-add away from maturity:

        prev coupon  = maturity - k     periods
        next coupon  = maturity - (k-1) periods
        j-th of the k remaining flows = maturity - (k-1-j) periods

    Anchoring on maturity for every date, rather than stepping period by period
    from the previous one, is what keeps a month-end schedule exact.  Stepping
    compounds the day clamp: 31 Mar back six months is 30 Sep, and stepping six
    months on from THAT gives 30 March, silently moving a 31st bond onto the
    30th and shifting every subsequent accrual by a day.  A single add from
    maturity cannot drift, because the anchor never changes.

    The seed is the whole number of months between the two dates divided by the
    period; it is never more than one period out, so the two correction passes
    below are bounded regardless of book size.
    """
    reference = np.asarray(reference, dtype=_DAY)
    maturity = np.broadcast_to(np.asarray(maturity, dtype=_DAY), reference.shape)
    per = np.broadcast_to(_months_per_period(frequency), reference.shape)
    per_i = np.nan_to_num(per, nan=1.0).astype(np.int64)
    per_i = np.where(per_i < 1, 1, per_i)

    ry, rm, _ = ymd(reference)
    my, mm, _ = ymd(maturity)
    month_gap = (my - ry) * 12 + (mm - rm)

    k = np.floor(month_gap / per_i).astype(np.int64)
    k = np.maximum(k, 0)
    bad = np.isnat(reference) | np.isnat(maturity) | ~np.isfinite(per)
    k = np.where(bad, 0, k)

    # Too far forward: the candidate is still after the reference date.
    for _ in range(4):
        cand = add_months(maturity, -(k * per_i))
        too_late = (cand > reference) & ~bad
        if not too_late.any():
            break
        k = np.where(too_late, k + 1, k)

    # Too far back: a later coupon would also still be on or before reference.
    for _ in range(4):
        cand_prev = add_months(maturity, -((k - 1) * per_i))
        too_early = (k >= 1) & (cand_prev <= reference) & ~bad
        if not too_early.any():
            break
        k = np.where(too_early, k - 1, k)

    return k, per_i


def prev_coupon_date(
    reference: np.ndarray, maturity: np.ndarray, frequency: np.ndarray
) -> np.ndarray:
    """Last coupon date on or before `reference`, rolled BACK from maturity.

    Rolling back from maturity rather than forward from issue reproduces the
    bond's real payment dates without needing the issue date at all: the final
    coupon is on the maturity date by construction and every earlier one is a
    whole number of periods before it.
    """
    k, per_i = _periods_back(reference, maturity, frequency)
    maturity = np.broadcast_to(np.asarray(maturity, dtype=_DAY), np.asarray(reference, dtype=_DAY).shape)
    return add_months(maturity, -(k * per_i))


def next_coupon_date(
    reference: np.ndarray, maturity: np.ndarray, frequency: np.ndarray
) -> np.ndarray:
    """First coupon date strictly after `reference`."""
    k, per_i = _periods_back(reference, maturity, frequency)
    maturity = np.broadcast_to(np.asarray(maturity, dtype=_DAY), np.asarray(reference, dtype=_DAY).shape)
    return add_months(maturity, -((k - 1) * per_i))


def coupon_schedule(
    settlement: np.ndarray,
    maturity: np.ndarray,
    frequency: np.ndarray,
    coupon_rate: np.ndarray,
    max_flows: int = 1200,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Remaining cash flows for a whole book, as one padded matrix.

    Returns `(dates, amounts, valid)`, each shaped `(n_bonds, max_n)`:

        dates    the payment date of each flow (NaT where padded)
        amounts  cash per 100 nominal, redemption folded into the last flow
        valid    True where the column is a real flow for that bond

    Building the schedule as a rectangular matrix is the whole reason pull to
    par is fast here.  The workbook walked a per-bond loop that re-read a
    13-row curve range for every cash flow - roughly 70,000 range reads per
    recalculation.  One padded matrix turns the same job into a single
    vectorised discounting step.

    Every date is one month-add from maturity (see `_periods_back`), so the
    last real flow lands exactly on the maturity date by construction and a
    month-end bond cannot drift off its own payment dates.

    A settlement date inside the final coupon period yields a single flow, the
    redemption - no special case is needed for it.
    """
    settlement = np.asarray(settlement, dtype=_DAY)
    maturity = np.asarray(maturity, dtype=_DAY)
    freq = np.asarray(frequency, dtype=np.float64)
    rate = np.asarray(coupon_rate, dtype=np.float64)
    n = settlement.shape[0]

    # k periods back from maturity is the coupon on or before settlement, so
    # exactly k flows remain strictly after it.
    counts, per_i = _periods_back(settlement, maturity, freq)
    counts = np.where(np.isnat(settlement) | np.isnat(maturity), 0, counts)
    counts = np.where(settlement >= maturity, 0, counts)
    counts = np.clip(counts, 0, max_flows)

    width = max(int(counts.max()) if counts.size else 0, 1)

    steps = np.arange(width, dtype=np.int64)[None, :]
    valid = steps < counts[:, None]

    # Column j is the j-th remaining flow: maturity minus (count-1-j) periods.
    back = (counts[:, None] - 1 - steps) * per_i[:, None]
    dates = add_months(np.broadcast_to(maturity[:, None], (n, width)), -back)

    last_idx = np.maximum(counts - 1, 0)
    rows = np.arange(n)

    with np.errstate(divide="ignore", invalid="ignore"):
        coupon_amt = np.where(freq > 0, 100.0 * rate / freq, np.nan)
    amounts = np.broadcast_to(coupon_amt[:, None], (n, width)).copy()
    amounts[rows, last_idx] = amounts[rows, last_idx] + 100.0
    amounts = np.where(valid, amounts, 0.0)
    dates = np.where(valid, dates, np.datetime64("NaT", "D"))

    return dates, amounts, valid


def coupons_paid_between(
    start: np.ndarray,
    end: np.ndarray,
    maturity: np.ndarray,
    coupon_rate: np.ndarray,
    frequency: np.ndarray,
    max_flows: int = 1200,
) -> np.ndarray:
    """Coupon cash per 100 nominal actually PAID in `(start, end]`.

    This is the term the legacy workbook did not have, and its absence is
    exactly why a bond that went ex-coupon inside the period showed a one-day
    residual the size of its coupon: the dirty price dropped by the coupon and
    nothing on the explained side put the cash back.  Redemption is excluded -
    a bond maturing inside the period is a position event, not carry.
    """
    dates, amounts, valid = coupon_schedule(start, maturity, frequency, coupon_rate, max_flows)
    start_b = np.asarray(start, dtype=_DAY)[:, None]
    end_b = np.asarray(end, dtype=_DAY)[:, None]
    maturity_b = np.asarray(maturity, dtype=_DAY)[:, None]

    paid = valid & (dates > start_b) & (dates <= end_b) & (dates < maturity_b)
    with np.errstate(invalid="ignore"):
        return np.where(paid, amounts, 0.0).sum(axis=1)


# --------------------------------------------------------------------------- #
# text -> code mappers (the Bloomberg descriptions the sheet used to carry)
# --------------------------------------------------------------------------- #

_DCC_EXACT: Final[dict[str, int]] = {
    "ACT/ACT": DCC_ACT_ACT_ICMA,
    "ACT/ACTICMA": DCC_ACT_ACT_ICMA,
    "ACT/ACTISDA": DCC_ACT_ACT_ISDA,
    "ACT/365": DCC_ACT_365F,
    "ACT/365F": DCC_ACT_365F,
    "ACT/365FIXED": DCC_ACT_365F,
    "ACT/365L": DCC_ACT_365F,
    "NL/365": DCC_ACT_365F,
    "ACT/360": DCC_ACT_360,
    "30/360": DCC_30_360_BOND,
    "30/360BOND": DCC_30_360_BOND,
    "30E/360": DCC_30E_360,
    "ISMA/30/360": DCC_30E_360,
    "30E/360ISDA": DCC_30E_360_ISDA,
}


def day_count_from_description(description: object) -> float:
    """Map a `DAY_CNT_DES` string onto a convention code, or NaN.

    The workbook's mapper, kept because the strings it normalises are the real
    ones the market-data feed serves: `ACT/360(102)`, `ISMA-30/360 NONEOM`,
    `ISDA SWAPS:30/360` and friends.  An unrecognised description returns NaN
    rather than guessing: a bond priced on the wrong accrual basis is worse
    than a bond with no accrued at all, because it looks right.
    """
    if description is None:
        return float("nan")
    s = str(description)
    if not s.strip():
        return float("nan")

    for ch in ("\r", "\n", "\t", "\xa0"):
        s = s.replace(ch, " ")
    s = s.strip().upper()
    for a, b in (("-", "/"), ("_", "/"), ("\\", "/"), (":", " "), ("(", " "), (")", " ")):
        s = s.replace(a, b)
    for a, b in (
        ("ACTUAL", "ACT"),
        ("NON EOM", "NONEOM"),
        ("NON/EOM", "NONEOM"),
        ("BOND BASIS", "BOND"),
        ("US SIA", "US"),
        ("NASD", "US"),
    ):
        s = s.replace(a, b)
    while "  " in s:
        s = s.replace("  ", " ")
    compact = s.replace(" ", "")

    if compact in _DCC_EXACT:
        return float(_DCC_EXACT[compact])

    if "ISDA" in compact and ("ACT/ACT" in compact or "ACTACT" in compact):
        return float(DCC_ACT_ACT_ISDA)
    if "ACT/ACT" in compact or "ACTACT" in compact:
        return float(DCC_ACT_ACT_ICMA)
    if any(t in compact for t in ("ACT/365", "ACT365", "A/365", "365F")):
        return float(DCC_ACT_365F)
    if any(t in compact for t in ("ACT/360", "ACT360", "A/360")):
        return float(DCC_ACT_360)
    if "ISMA" in compact and "30/360" in compact:
        return float(DCC_30E_360)
    if "30E/360" in compact and "ISDA" in compact:
        return float(DCC_30E_360_ISDA)
    if "30E/360" in compact:
        return float(DCC_30E_360)
    if "ISDA" in compact and "SWAPS" in compact and "30/360" in compact:
        return float(DCC_30_360_BOND)
    if "30/360" in compact:
        return float(DCC_30_360_BOND)

    return float("nan")


_FREQ_WORDS: Final[dict[str, int]] = {
    "1": 1, "A": 1, "ANNUAL": 1, "ANNUALLY": 1, "12M": 1, "Y": 1, "YEARLY": 1,
    "2": 2, "S": 2, "SA": 2, "SEMI": 2, "SEMIANNUAL": 2, "SEMI-ANNUAL": 2, "6M": 2,
    "4": 4, "Q": 4, "QUARTERLY": 4, "3M": 4,
    "12": 12, "M": 12, "MONTHLY": 12, "1M": 12,
}


def coupon_frequency_from_description(value: object) -> float:
    """Map a coupon-frequency description onto payments per year, or NaN."""
    if value is None:
        return float("nan")
    s = str(value).strip().upper()
    if not s:
        return float("nan")
    if s in _FREQ_WORDS:
        return float(_FREQ_WORDS[s])
    try:
        n = int(float(s))
    except ValueError:
        return float("nan")
    return float(n) if n in VALID_FREQUENCIES else float("nan")


def excel_price_basis(dcc: np.ndarray) -> np.ndarray:
    """Internal convention code -> the basis code Excel's PRICE() expects.

    Kept so a figure produced here can be reconciled against a spreadsheet the
    desk builds by hand, which is the only reason the mapping exists.
    """
    code = np.asarray(dcc, dtype=np.float64)
    out = np.full(code.shape, 1.0)                       # Actual/Actual
    out = np.where(code == DCC_ACT_365F, 3.0, out)       # Actual/365
    out = np.where(code == DCC_ACT_360, 2.0, out)        # Actual/360
    out = np.where(code == DCC_30_360_BOND, 0.0, out)    # US 30/360
    out = np.where(np.isin(code, (DCC_30E_360, DCC_30E_360_ISDA)), 4.0, out)  # European
    return out
