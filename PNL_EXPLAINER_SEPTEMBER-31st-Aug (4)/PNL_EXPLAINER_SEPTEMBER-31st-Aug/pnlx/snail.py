"""The snail methodology.

A snail trail is a PATH, not a bar.  Instead of asking "how big was each PnL
component today", it asks "where does the book walk to as each component is
added, and where has it walked over time".  The shape of the path is the
information: a tight coil round the origin is a book whose model tracks its
marks; a trail that drifts steadily away in one direction is a systematic bias,
not noise, and the direction says which leg is causing it.

Two trails are built, because there are two useful questions.

INTRADAY SNAIL - one day, walked leg by leg
    Start at zero.  Add the duration leg, then convexity, then carry, then FX,
    then the model hedge, then the hedge basis, and finally the residual.  Each
    step plots the CUMULATIVE explained PnL against the practical PnL still
    outstanding at that point.  The walk ends, by construction, on the practical
    total with nothing outstanding.

    Read it as: how far does each leg carry the explanation, and which one does
    the heavy lifting?  A long horizontal run on "hedge basis" means the model
    hedge is not tracking the real one.  A big final jump on "residual" means
    the bond leg is not being explained.

HISTORICAL SNAIL - many days, cumulative
    Each day is a point at (cumulative practical PnL, cumulative explained PnL).
    A perfect model traces the 45-degree line.  Distance from that line is the
    accumulated unexplained PnL, and the shape says what kind of error it is:

        along the diagonal          the model is tracking
        drifting one side steadily  a systematic bias - a convention, a stale
                                    input, or a missing leg
        looping back and forth      noise, or marks arriving on different days
        a sudden right angle        something changed on that date

    A second pair of axes plots cumulative residual against cumulative absolute
    risk taken, which is the risk/return reading of the same trail: it answers
    "is the unexplained PnL growing with the risk we are running, or
    independently of it".

The history is fed from the daily CSV the run appends to, so the trail extends
by one point per run with no separate bookkeeping.
"""

from __future__ import annotations

import csv
import datetime as _dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np

from .aggregate import PortfolioSummary
from .columns import ColumnFrame

__all__ = ["SnailStep", "SnailPoint", "SnailTrail", "build_intraday_snail", "load_history", "build_history_snail"]


@dataclass(frozen=True, slots=True)
class SnailStep:
    """One leg of the intraday walk."""

    order: int
    label: str
    amount: float
    cumulative_explained: float
    outstanding: float
    share_of_practical: float
    note: str = ""

    def to_record(self) -> dict[str, object]:
        return {
            "order": self.order,
            "label": self.label,
            "amount": self.amount,
            "cumulative_explained": self.cumulative_explained,
            "outstanding": self.outstanding,
            "share_of_practical": self.share_of_practical,
            "note": self.note,
        }


@dataclass(frozen=True, slots=True)
class SnailPoint:
    """One day on the historical trail."""

    date: _dt.date
    practical: float
    explained: float
    residual: float
    cum_practical: float
    cum_explained: float
    cum_residual: float
    cum_abs_dv01: float
    label: str = ""

    def to_record(self) -> dict[str, object]:
        return {
            "date": self.date.isoformat() if self.date else None,
            "practical_pnl": self.practical,
            "explained_pnl": self.explained,
            "residual_pnl": self.residual,
            "cum_practical_pnl": self.cum_practical,
            "cum_explained_pnl": self.cum_explained,
            "cum_residual_pnl": self.cum_residual,
            "cum_abs_dv01": self.cum_abs_dv01,
            "label": self.label,
        }


@dataclass(slots=True)
class SnailTrail:
    """Both trails, plus the diagnostics that read them."""

    steps: list[SnailStep] = field(default_factory=list)
    points: list[SnailPoint] = field(default_factory=list)
    diagnostics: dict[str, object] = field(default_factory=dict)

    @property
    def has_history(self) -> bool:
        return len(self.points) > 1

    def step_records(self) -> list[dict[str, object]]:
        return [s.to_record() for s in self.steps]

    def point_records(self) -> list[dict[str, object]]:
        return [p.to_record() for p in self.points]


# --------------------------------------------------------------------------- #
# intraday
# --------------------------------------------------------------------------- #

#: The walk order, and why it is this order: model first, from the most
#: mechanical leg to the least, then the marks.  Each step should shrink what is
#: still outstanding; a step that does not is the one to look at.
_WALK = (
    ("Duration", "pnl_duration_total",
     "First order in the yield move, split by the row's spread framework."),
    ("Convexity", "pnl_convexity",
     "Second order in the same move. Small on a quiet day, not on a violent one."),
    ("Carry", "carry_total",
     "Coupon accrual plus pull to par: what the book earns for holding still."),
    ("FX", "pnl_fx",
     "Translation of the opening position at the new fix."),
    ("FX cross", "pnl_fx_cross",
     "Local move revalued at the FX move; booked so the FX identity is exact."),
    ("Hedge (model)", "model_hedge_pnl",
     "What the hedges should have made from the curves they neutralise."),
    ("Hedge basis", "hedge_basis_pnl",
     "Actual less model hedge: CTD basis and swap spread mismatch."),
)


def build_intraday_snail(bonds: ColumnFrame, summary: PortfolioSummary) -> list[SnailStep]:
    """Walk one day's attribution leg by leg.

    Only rows with a complete explained figure are walked, so the walk lands on
    the same practical total the bridge uses.  Rows that could not be attributed
    are the bridge's "not attributed" line and are not part of this path - a
    trail that ended somewhere other than its own total would be worse than
    useless.
    """
    attributed = np.isfinite(np.asarray(bonds.get("total_explained"), dtype=np.float64))
    practical = float(
        np.nansum(np.asarray(bonds.get("practical_pnl"), dtype=np.float64)[attributed])
    )
    denominator = abs(practical) if practical else float("nan")

    steps: list[SnailStep] = []
    running = 0.0
    for order, (label, column, note) in enumerate(_WALK, start=1):
        values = np.asarray(bonds.get(column), dtype=np.float64)[attributed]
        amount = float(np.nansum(values))
        running += amount
        steps.append(
            SnailStep(
                order=order,
                label=label,
                amount=amount,
                cumulative_explained=running,
                outstanding=practical - running,
                share_of_practical=amount / denominator if denominator == denominator else float("nan"),
                note=note,
            )
        )

    residual = practical - running
    steps.append(
        SnailStep(
            order=len(steps) + 1,
            label="Residual",
            amount=residual,
            cumulative_explained=practical,
            outstanding=0.0,
            share_of_practical=residual / denominator if denominator == denominator else float("nan"),
            note="What the model did not reach. The walk closes on the practical total "
                 "of the attributed rows by construction, so this is the whole of the gap. "
                 "Rows that could not be attributed are the bridge's 'not attributed' "
                 "line and are outside this walk.",
        )
    )
    return steps


# --------------------------------------------------------------------------- #
# history
# --------------------------------------------------------------------------- #


def load_history(path: Path) -> list[dict[str, object]]:
    """Read the appended daily history file, oldest first.

    A missing file is not an error: the first run of a new book has no history
    and should produce a one-point trail rather than fail.
    """
    if not path.is_file():
        return []
    rows: list[dict[str, object]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        for raw in csv.DictReader(fh):
            rows.append(dict(raw))
    rows.sort(key=lambda r: str(r.get("as_of", "")))
    return rows


def build_history_snail(
    history: Sequence[dict[str, object]],
    *,
    today: dict[str, object] | None = None,
) -> tuple[list[SnailPoint], dict[str, object]]:
    """Turn the daily history into a cumulative trail and read its shape."""
    rows = list(history)
    if today is not None:
        stamp = str(today.get("as_of", ""))
        rows = [r for r in rows if str(r.get("as_of", "")) != stamp]
        rows.append(dict(today))
        rows.sort(key=lambda r: str(r.get("as_of", "")))

    points: list[SnailPoint] = []
    cum_p = cum_e = cum_r = cum_dv01 = 0.0

    for row in rows:
        practical = _num(row.get("practical_pnl"))
        explained = _num(row.get("total_explained"))
        residual = _num(row.get("residual_pnl"))
        dv01 = abs(_num(row.get("bond_dv01")))

        cum_p += practical
        cum_e += explained
        cum_r += residual
        cum_dv01 += dv01

        points.append(
            SnailPoint(
                date=_date(row.get("as_of")),
                practical=practical,
                explained=explained,
                residual=residual,
                cum_practical=cum_p,
                cum_explained=cum_e,
                cum_residual=cum_r,
                cum_abs_dv01=cum_dv01,
                label=str(row.get("label", "") or ""),
            )
        )

    return points, _read_the_trail(points)


def _read_the_trail(points: Sequence[SnailPoint]) -> dict[str, object]:
    """Say in words what the shape of the trail means.

    A chart nobody can read is decoration.  These are the readings a desk would
    take by eye, computed so they are in the report whether or not anyone looks
    at the picture.
    """
    n = len(points)
    if n == 0:
        return {"verdict": "no history yet", "days": 0}
    if n == 1:
        return {
            "verdict": "first day - a trail needs at least two points",
            "days": 1,
            "cum_residual": points[0].cum_residual,
        }

    residuals = np.array([p.residual for p in points], dtype=np.float64)
    practicals = np.array([p.practical for p in points], dtype=np.float64)
    finite = np.isfinite(residuals)

    cum_residual = points[-1].cum_residual
    cum_practical = points[-1].cum_practical
    mean_residual = float(np.nanmean(residuals[finite])) if finite.any() else float("nan")
    std_residual = float(np.nanstd(residuals[finite])) if finite.sum() > 1 else float("nan")

    # Is the mean residual distinguishable from zero?
    #
    # The test has to be the mean against the STANDARD ERROR of the mean
    # (std / sqrt(n)), not against the standard deviation.  Comparing the two
    # directly is not sample-size aware: a genuinely random series of 25 days
    # has a mean around 0.2 standard deviations from zero purely by chance, so
    # a fixed threshold on that ratio calls honest noise a drift and gets more
    # wrong the shorter the history is.  Dividing by sqrt(n) is what makes the
    # reading mean the same thing after 10 days as after 200.
    n_finite = int(finite.sum())
    standard_error = (
        std_residual / np.sqrt(n_finite)
        if std_residual and np.isfinite(std_residual) and std_residual > 0 and n_finite
        else float("nan")
    )
    t_stat = (
        mean_residual / standard_error
        if standard_error and np.isfinite(standard_error) and standard_error > 0
        else float("nan")
    )
    drift_ratio = (
        abs(mean_residual) / std_residual
        if std_residual and np.isfinite(std_residual) and std_residual > 0
        else float("nan")
    )
    same_sign = (
        float(np.mean(np.sign(residuals[finite]) == np.sign(mean_residual)))
        if finite.any()
        else float("nan")
    )

    if np.isfinite(t_stat) and abs(t_stat) > 2.0 and same_sign > 0.65:
        verdict = (
            "systematic drift - the residual keeps landing on the same side by more "
            "than chance explains, so something is missing from the model rather "
            "than noisy"
        )
    elif np.isfinite(t_stat) and abs(t_stat) < 1.0:
        verdict = (
            "coiled - the mean residual is indistinguishable from zero, which is "
            "what a working model looks like"
        )
    else:
        verdict = "mixed - some one-sided drift, but not beyond what chance would give"

    with np.errstate(divide="ignore", invalid="ignore"):
        hit_rate = (
            float(np.mean(np.abs(residuals[finite]) < 0.05 * np.abs(practicals[finite])))
            if finite.any()
            else float("nan")
        )

    return {
        "verdict": verdict,
        "days": n,
        "cum_practical": cum_practical,
        "cum_explained": points[-1].cum_explained,
        "cum_residual": cum_residual,
        "cum_residual_pct": cum_residual / abs(cum_practical) if cum_practical else float("nan"),
        "mean_daily_residual": mean_residual,
        "std_daily_residual": std_residual,
        "standard_error": standard_error,
        "t_statistic": t_stat,
        "drift_ratio": drift_ratio,
        "same_sign_share": same_sign,
        "days_within_5pct": hit_rate,
        "worst_day": max(points, key=lambda p: abs(p.residual)).date.isoformat()
        if points and points[0].date
        else None,
    }


def _num(value: object) -> float:
    if value is None:
        return 0.0
    try:
        text = str(value).strip()
        return float(text) if text else 0.0
    except (TypeError, ValueError):
        return 0.0


def _date(value: object) -> _dt.date | None:
    if value is None:
        return None
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()[:10]
    try:
        return _dt.date.fromisoformat(text)
    except ValueError:
        return None
