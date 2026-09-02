"""Zero curves and the interpolation the whole model is calibrated on.

The workbook held three curves per currency on one wide sheet (OIS_Curves),
with each currency owning its own tenor column - EUR in B, USD in S, GBP in AJ.
Adding a currency meant editing a resolver in two places, and the interpolation
itself re-read a 13-row worksheet range for every single lookup.

Here a curve is an object.  `CurveSet` is a dict of them keyed on
`(currency, curve_type, snapshot)`, and interpolation is one vectorised pass
over a whole book of tenors at once.

Three curve types, and the reason each exists:

    OIS   the risk-free discounting curve (ESTR in EUR, SOFR in USD, SONIA in
          GBP).  `r` in the decomposition below.
    GOV   the government benchmark curve.  Gov - OIS is the sovereign/collateral
          basis `g`.
    SWAP  the IBOR-projection swap curve (EURIBOR in EUR).  Swap - Gov is `q`.

which is what makes the identity the attribution rests on hold:

    y  =  r  +  g  +  q  +  i

        y  bond yield to maturity
        r  OIS zero rate at the bond's tenor
        g  Gov  - OIS   government/OIS basis
        q  Swap - Gov   swap/government basis
        i  y    - Swap  the bond's own I-spread

Every term is a difference of two quoted rates, so the chain telescopes exactly.
That is the whole reason the spread frameworks can split the same yield move
several different ways without any of them changing the total.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Iterator, Literal, Mapping

import numpy as np

__all__ = [
    "Snapshot",
    "CurveType",
    "CurveKey",
    "ZeroCurve",
    "CurveSet",
    "CurveError",
]

Snapshot = Literal["prior", "current"]
CurveType = Literal["OIS", "GOV", "SWAP"]

CURVE_TYPES: tuple[str, ...] = ("OIS", "GOV", "SWAP")
SNAPSHOTS: tuple[str, ...] = ("prior", "current")


class CurveError(ValueError):
    """Raised when a curve cannot be built or cannot be found."""


@dataclass(frozen=True, slots=True)
class CurveKey:
    """What identifies one curve: currency, type and which day it is from."""

    currency: str
    curve_type: str
    snapshot: str

    def __post_init__(self) -> None:
        object.__setattr__(self, "currency", self.currency.strip().upper())
        object.__setattr__(self, "curve_type", self.curve_type.strip().upper())
        object.__setattr__(self, "snapshot", self.snapshot.strip().lower())
        if self.curve_type not in CURVE_TYPES:
            raise CurveError(f"curve_type must be one of {CURVE_TYPES}, got {self.curve_type!r}")
        if self.snapshot not in SNAPSHOTS:
            raise CurveError(f"snapshot must be one of {SNAPSHOTS}, got {self.snapshot!r}")

    def __str__(self) -> str:  # pragma: no cover - display only
        return f"{self.currency}/{self.curve_type}/{self.snapshot}"


@dataclass(frozen=True, slots=True)
class ZeroCurve:
    """A zero curve as sorted (tenor in years, rate in PERCENT) node arrays.

    Rates are held in percent because that is how every upstream source quotes
    them and how they sit on the report; `rates_decimal` is the one place the
    conversion happens.  Mixing the two units is the classic way to make a
    discount factor look plausible and be a hundred times wrong, so the unit is
    part of the attribute name, not a comment.

    Interpolation is linear in tenor with FLAT extrapolation beyond both ends,
    which is what the workbook did.  Flat rather than linear extrapolation
    matters at the short end: extending the slope of the 3M-6M segment back to
    a two-day bond can produce a negative discount factor on a steep curve.
    """

    key: CurveKey
    tenors: np.ndarray  # float64, strictly ascending
    rates: np.ndarray   # float64, percent

    @classmethod
    def build(
        cls,
        key: CurveKey,
        tenors: Iterable[float],
        rates: Iterable[float],
    ) -> "ZeroCurve":
        """Sort, drop non-finite nodes, and refuse a curve too thin to use."""
        t = np.asarray(list(tenors), dtype=np.float64).ravel()
        r = np.asarray(list(rates), dtype=np.float64).ravel()
        if t.shape != r.shape:
            raise CurveError(f"{key}: {t.size} tenors against {r.size} rates")

        good = np.isfinite(t) & np.isfinite(r) & (t > 0)
        t, r = t[good], r[good]

        order = np.argsort(t, kind="stable")
        t, r = t[order], r[order]

        # Collapse duplicate tenors onto their mean rather than letting a
        # repeated node create a zero-width interpolation segment.
        if t.size and np.any(np.diff(t) == 0):
            uniq, inverse = np.unique(t, return_inverse=True)
            summed = np.zeros(uniq.size, dtype=np.float64)
            counts = np.zeros(uniq.size, dtype=np.float64)
            np.add.at(summed, inverse, r)
            np.add.at(counts, inverse, 1.0)
            t, r = uniq, summed / counts

        if t.size < 2:
            raise CurveError(
                f"{key}: needs at least 2 usable nodes, got {t.size}. "
                "A single point cannot be interpolated and must not be "
                "silently treated as a flat curve."
            )

        t.flags.writeable = False
        r.flags.writeable = False
        return cls(key=key, tenors=t, rates=r)

    # -- access ------------------------------------------------------------- #

    @property
    def rates_decimal(self) -> np.ndarray:
        return self.rates / 100.0

    @property
    def n_nodes(self) -> int:
        return int(self.tenors.size)

    def rate(self, years) -> np.ndarray:
        """Interpolated zero rate in PERCENT at the given tenor(s).

        `np.interp` already flat-extrapolates at both ends, which is precisely
        the behaviour required, so no clipping is needed around it.  Non-finite
        tenors return NaN instead of being clamped onto the front node: a bond
        with no maturity has no rate, and a fabricated one would flow straight
        into its DV01.
        """
        y = np.asarray(years, dtype=np.float64)
        out = np.interp(y, self.tenors, self.rates)
        return np.where(np.isfinite(y), out, np.nan)

    def rate_decimal(self, years) -> np.ndarray:
        return self.rate(years) / 100.0

    def discount_factor(self, years, spread_decimal=0.0) -> np.ndarray:
        """Annually compounded discount factor `(1 + z + s) ** -t`.

        Annual compounding, not continuous, because that is the convention the
        pull-to-par forward identity below is written in.  Using `exp(-zt)` here
        and the compounded form there would leave a small standing difference in
        carry that looks like model error and is only a unit mismatch.
        """
        y = np.asarray(years, dtype=np.float64)
        z = self.rate_decimal(y) + np.asarray(spread_decimal, dtype=np.float64)
        base = 1.0 + z
        with np.errstate(invalid="ignore", divide="ignore"):
            df = np.where(base > 0, base ** (-y), np.nan)
        return np.where(np.isfinite(y), df, np.nan)


@dataclass(slots=True)
class CurveSet:
    """Every curve for a run, addressable by (currency, type, snapshot)."""

    curves: dict[CurveKey, ZeroCurve] = field(default_factory=dict)

    # -- construction ------------------------------------------------------- #

    def add(self, curve: ZeroCurve) -> None:
        self.curves[curve.key] = curve

    @classmethod
    def from_records(cls, records: Iterable[Mapping[str, object]]) -> "CurveSet":
        """Build from long-format rows.

        Each row carries one node of one curve for BOTH snapshots:

            currency, curve_type, tenor_years, rate_prior, rate_current

        Long format on purpose.  The workbook's wide layout - one column block
        per currency, hard-coded letters, a separate tenor column each - is what
        forced a resolver to exist at all, and what made adding a fourth
        currency a change in two files.  Here a new currency is new rows.
        """
        buckets: dict[CurveKey, list[tuple[float, float]]] = {}

        for row in records:
            ccy = str(row.get("currency", "")).strip().upper()
            ctype = str(row.get("curve_type", "")).strip().upper()
            if not ccy or not ctype:
                continue
            try:
                tenor = float(row.get("tenor_years"))  # type: ignore[arg-type]
            except (TypeError, ValueError):
                continue
            if not np.isfinite(tenor):
                continue

            for snapshot, column in (("prior", "rate_prior"), ("current", "rate_current")):
                raw = row.get(column)
                if raw is None or str(raw).strip() == "":
                    continue
                try:
                    rate = float(raw)  # type: ignore[arg-type]
                except (TypeError, ValueError):
                    continue
                if not np.isfinite(rate):
                    continue
                buckets.setdefault(CurveKey(ccy, ctype, snapshot), []).append((tenor, rate))

        out = cls()
        for key, nodes in buckets.items():
            tenors = [n[0] for n in nodes]
            rates = [n[1] for n in nodes]
            try:
                out.add(ZeroCurve.build(key, tenors, rates))
            except CurveError:
                # A curve with one usable node is dropped, not faked.  Every
                # bond that needed it reports a missing-curve status instead of
                # a rate invented from a single point.
                continue
        return out

    # -- lookup ------------------------------------------------------------- #

    def get(self, currency: str, curve_type: str, snapshot: str) -> ZeroCurve | None:
        try:
            return self.curves.get(CurveKey(currency, curve_type, snapshot))
        except CurveError:
            return None

    def require(self, currency: str, curve_type: str, snapshot: str) -> ZeroCurve:
        curve = self.get(currency, curve_type, snapshot)
        if curve is None:
            raise CurveError(
                f"no {curve_type} curve for {currency} on the {snapshot} snapshot"
            )
        return curve

    @property
    def currencies(self) -> tuple[str, ...]:
        return tuple(sorted({k.currency for k in self.curves}))

    def __len__(self) -> int:
        return len(self.curves)

    def __iter__(self) -> Iterator[ZeroCurve]:
        return iter(self.curves.values())

    def __contains__(self, key: object) -> bool:
        return key in self.curves

    # -- the vectorised workhorse ------------------------------------------- #

    def rates_for(
        self,
        currencies: np.ndarray,
        years: np.ndarray,
        curve_type: str,
        snapshot: str,
    ) -> np.ndarray:
        """One interpolated rate per position, in PERCENT.

        Positions are grouped by currency and each group interpolated in a
        single `np.interp` call, so the cost is one pass per currency rather
        than one lookup per position.  A currency with no such curve comes back
        NaN across the board, which propagates into a blank spread and a
        missing-data status - never into a zero that reads as "no move".
        """
        ccy = np.asarray(currencies, dtype=object)
        y = np.asarray(years, dtype=np.float64)
        out = np.full(y.shape, np.nan, dtype=np.float64)

        if y.size == 0:
            return out

        normalised = np.array(
            [str(c).strip().upper() if c is not None else "" for c in ccy.ravel()],
            dtype=object,
        ).reshape(ccy.shape)

        for code in {c for c in normalised.tolist() if c}:
            curve = self.get(code, curve_type, snapshot)
            if curve is None:
                continue
            mask = normalised == code
            out[mask] = curve.rate(y[mask])

        return out

    def describe(self) -> list[dict[str, object]]:
        """Flat rows for the report's curve sheet."""
        rows: list[dict[str, object]] = []
        for key in sorted(self.curves, key=lambda k: (k.currency, k.curve_type, k.snapshot)):
            curve = self.curves[key]
            rows.append(
                {
                    "currency": key.currency,
                    "curve_type": key.curve_type,
                    "snapshot": key.snapshot,
                    "nodes": curve.n_nodes,
                    "min_tenor": float(curve.tenors[0]),
                    "max_tenor": float(curve.tenors[-1]),
                    "min_rate_pct": float(curve.rates.min()),
                    "max_rate_pct": float(curve.rates.max()),
                }
            )
        return rows
