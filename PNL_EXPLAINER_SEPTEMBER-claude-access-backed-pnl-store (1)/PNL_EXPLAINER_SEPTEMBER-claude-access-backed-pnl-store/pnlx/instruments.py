"""Position types and their column-array books.

Two views of the same data, on purpose:

    BondPosition / FuturePosition / SwapPosition
        one position, as a frozen dataclass.  Typed, documented, and what you
        get back when you index a book.  Used for readable construction, for
        tests, and for anything that legitimately deals with one instrument.

    BondBook / FutureBook / SwapBook
        the same positions as parallel numpy arrays.  This is what the engines
        actually run on: the attribution is a handful of array expressions over
        a whole book, not a loop over positions.

The books own the arrays, so a column is allocated once and every engine reads
the same buffer.  Building a book from a list of dataclasses costs one pass;
everything after that is vectorised.

SIGN CONVENTION - one rule, applied once
----------------------------------------
`notional` and `contracts` are SIGNED.  A short is a negative notional.  There
is no separate side flag anywhere in this package, and nothing multiplies by a
sign a second time.

The legacy workbook carried a `PositionSide` column in three places and the
comments in it are a record of what that costs: DirtyMV and DV01_EUR were
already signed by it, so any formula that also multiplied by PositionSide
silently flipped every short position - and then quietly corrupted every DV01
comparison downstream of it.  Signed notional and nothing else removes the
whole class of error.

UNITS - stated once, enforced by name
-------------------------------------
    prices           per 100 nominal
    yields, curves   PERCENT (3.25 means 3.25%)
    spreads          BASIS POINTS
    coupon_pct       PERCENT on the position record, decimal inside the maths
    fx               units of BASE currency per 1 unit of position currency
    notional         nominal in position currency, signed
    all PnL, MV, DV01  BASE currency
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, fields
from typing import Any, Iterable, Sequence

import numpy as np

from . import daycount as dc

__all__ = [
    "BondPosition",
    "FuturePosition",
    "SwapPosition",
    "BondBook",
    "FutureBook",
    "SwapBook",
    "SWAP_PLAIN",
    "SWAP_SYNTHETIC",
]

SWAP_PLAIN = "PLAIN"
SWAP_SYNTHETIC = "SYNTHETIC"

_DAY = np.dtype("datetime64[D]")


def _f(value: Any) -> float:
    """Anything -> float, with blanks and junk becoming NaN.

    Missing market data must stay missing.  Coercing it to zero is how a gap
    turns into a confident report of "no move", which is worse than a blank
    because nothing downstream can tell the difference.
    """
    if value is None:
        return float("nan")
    if isinstance(value, (int, float, np.floating, np.integer)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.upper() in {"NA", "N/A", "#N/A", "NAN", "NULL", "-", "#VALUE!", "#REF!"}:
        return float("nan")
    try:
        return float(text)
    except ValueError:
        return float("nan")


def _s(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.upper() in {"NAN", "NONE", "NULL"} else text


def _column(rows: Sequence[Any], name: str, kind: str) -> np.ndarray:
    """Pull one attribute off every position into a typed array."""
    raw = [getattr(r, name) for r in rows]
    if kind == "f":
        return np.array([_f(v) for v in raw], dtype=np.float64)
    if kind == "d":
        return dc.as_date_array(raw)
    return np.array([_s(v) for v in raw], dtype=object)


# --------------------------------------------------------------------------- #
# bonds
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class BondPosition:
    """One bond position, with both snapshots of its market data."""

    isin: str
    name: str = ""
    currency: str = "EUR"
    portfolio: str = ""
    acctg_cat: str = ""

    notional: float = float("nan")          # signed nominal
    coupon_pct: float = float("nan")        # percent, e.g. 3.75
    coupon_freq: float = float("nan")       # payments per year
    maturity: _dt.date | None = None
    book_value: float = float("nan")

    # prior (T-1) marks
    clean_px_prior: float = float("nan")
    dirty_px_prior: float = float("nan")
    ytm_prior: float = float("nan")         # percent
    zspread_prior: float = float("nan")     # bp
    asw_prior: float = float("nan")         # bp
    oas_prior: float = float("nan")         # bp
    mod_duration_prior: float = float("nan")

    # current (T0) marks
    clean_px_current: float = float("nan")
    dirty_px_current: float = float("nan")
    ytm_current: float = float("nan")       # percent
    zspread_current: float = float("nan")   # bp
    asw_current: float = float("nan")       # bp
    oas_current: float = float("nan")       # bp
    mod_duration_current: float = float("nan")
    oas_mod_duration_current: float = float("nan")
    convexity: float = float("nan")

    fx_prior: float = float("nan")
    fx_current: float = float("nan")
    funding_rate_prior: float = float("nan")   # percent
    funding_rate_current: float = float("nan")  # percent

    day_count_desc: str = ""
    day_count_code: float = float("nan")
    pricing_source: str = ""
    benchmark_bond: str = ""

    @property
    def coupon_decimal(self) -> float:
        return self.coupon_pct / 100.0


@dataclass(slots=True)
class BondBook:
    """A whole bond book as parallel arrays."""

    isin: np.ndarray
    name: np.ndarray
    currency: np.ndarray
    portfolio: np.ndarray
    acctg_cat: np.ndarray

    notional: np.ndarray
    coupon_pct: np.ndarray
    coupon_freq: np.ndarray
    maturity: np.ndarray
    book_value: np.ndarray

    clean_px_prior: np.ndarray
    dirty_px_prior: np.ndarray
    ytm_prior: np.ndarray
    zspread_prior: np.ndarray
    asw_prior: np.ndarray
    oas_prior: np.ndarray
    mod_duration_prior: np.ndarray

    clean_px_current: np.ndarray
    dirty_px_current: np.ndarray
    ytm_current: np.ndarray
    zspread_current: np.ndarray
    asw_current: np.ndarray
    oas_current: np.ndarray
    mod_duration_current: np.ndarray
    oas_mod_duration_current: np.ndarray
    convexity: np.ndarray

    fx_prior: np.ndarray
    fx_current: np.ndarray
    funding_rate_prior: np.ndarray
    funding_rate_current: np.ndarray

    day_count_desc: np.ndarray
    day_count_code: np.ndarray
    pricing_source: np.ndarray
    benchmark_bond: np.ndarray

    positions: tuple[BondPosition, ...] = ()

    _SPEC = (
        ("isin", "s"), ("name", "s"), ("currency", "s"), ("portfolio", "s"),
        ("acctg_cat", "s"),
        ("notional", "f"), ("coupon_pct", "f"), ("coupon_freq", "f"),
        ("maturity", "d"), ("book_value", "f"),
        ("clean_px_prior", "f"), ("dirty_px_prior", "f"), ("ytm_prior", "f"),
        ("zspread_prior", "f"), ("asw_prior", "f"), ("oas_prior", "f"),
        ("mod_duration_prior", "f"),
        ("clean_px_current", "f"), ("dirty_px_current", "f"), ("ytm_current", "f"),
        ("zspread_current", "f"), ("asw_current", "f"), ("oas_current", "f"),
        ("mod_duration_current", "f"), ("oas_mod_duration_current", "f"),
        ("convexity", "f"),
        ("fx_prior", "f"), ("fx_current", "f"),
        ("funding_rate_prior", "f"), ("funding_rate_current", "f"),
        ("day_count_desc", "s"), ("day_count_code", "f"),
        ("pricing_source", "s"), ("benchmark_bond", "s"),
    )

    @classmethod
    def from_positions(cls, positions: Iterable[BondPosition]) -> "BondBook":
        rows = tuple(positions)
        kwargs = {name: _column(rows, name, kind) for name, kind in cls._SPEC}

        # Normalise identifiers once, here, so nothing downstream has to guess
        # whether an ISIN has been upper-cased yet.  Hedge matching is an exact
        # string join; a stray space would silently orphan a hedge.
        kwargs["isin"] = np.array(
            [s.strip().upper() for s in kwargs["isin"].tolist()], dtype=object
        )
        kwargs["currency"] = np.array(
            [s.strip().upper() for s in kwargs["currency"].tolist()], dtype=object
        )

        # Day count: an explicit code wins; otherwise map the description.  An
        # unmappable description stays NaN rather than defaulting to ACT/ACT -
        # a bond accruing on the wrong basis looks right and is not.
        code = kwargs["day_count_code"]
        from_desc = np.array(
            [dc.day_count_from_description(d) for d in kwargs["day_count_desc"].tolist()],
            dtype=np.float64,
        )
        kwargs["day_count_code"] = np.where(np.isfinite(code), code, from_desc)

        # Coupon frequency likewise accepts either a number or a description.
        freq = kwargs["coupon_freq"]
        bad = ~np.isin(freq, np.array(dc.VALID_FREQUENCIES, dtype=np.float64))
        if bad.any():
            recovered = np.array(
                [
                    dc.coupon_frequency_from_description(getattr(r, "coupon_freq"))
                    for r in rows
                ],
                dtype=np.float64,
            )
            kwargs["coupon_freq"] = np.where(bad, recovered, freq)

        return cls(positions=rows, **kwargs)

    # -- sequence protocol -------------------------------------------------- #

    def __len__(self) -> int:
        return int(self.isin.size)

    def __getitem__(self, index: int) -> BondPosition:
        return self.positions[index]

    def __iter__(self):
        return iter(self.positions)

    # -- derived ------------------------------------------------------------ #

    @property
    def coupon_decimal(self) -> np.ndarray:
        return self.coupon_pct / 100.0

    @property
    def is_base_currency_flag(self) -> np.ndarray:
        return np.array([c == "EUR" for c in self.currency.tolist()])

    def years_to_maturity(self, as_of: np.datetime64, basis: float = 365.0) -> np.ndarray:
        """Tenor used to pick the point on each curve."""
        days = (
            self.maturity.astype("datetime64[D]").astype("int64")
            - np.datetime64(as_of, "D").astype("int64")
        ).astype(np.float64)
        return np.where(np.isnat(self.maturity), np.nan, np.maximum(days, 0.0) / basis)

    def duplicate_isin_mask(self) -> np.ndarray:
        """Rows whose ISIN appears more than once in the book.

        Hedges attach to bonds by ISIN, so a bond held in two portfolios makes
        BOTH rows claim the full hedge PnL and DV01 of that ISIN.  This is not
        hypothetical - it is the reason the legacy dashboard carried a "Rows
        Sharing an ISIN" tile.  Flagging it per row lets the report say which
        positions are overstated rather than only how many.
        """
        values, counts = np.unique(self.isin, return_counts=True)
        repeated = set(values[counts > 1].tolist())
        return np.array([i in repeated for i in self.isin.tolist()])


# --------------------------------------------------------------------------- #
# futures
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class FuturePosition:
    """One bond-future position and its cheapest-to-deliver reference."""

    contract_code: str
    exchange: str = ""
    currency: str = "EUR"
    portfolio: str = ""
    linked_isin: str = ""
    hedge_type: str = ""

    contracts: float = float("nan")       # signed
    face_value: float = float("nan")
    deliv_date: _dt.date | None = None

    ctd_isin: str = ""
    ctd_cf: float = float("nan")          # conversion factor
    ctd_dirty_px_current: float = float("nan")

    avg_entry_px: float = float("nan")
    fut_px_prior: float = float("nan")
    fut_px_current: float = float("nan")
    fut_val_pt: float = float("nan")      # cash per full price point per contract
    fut_px_val_bp: float = float("nan")   # see FuturesConfig.dv01_includes_point_value

    fx_prior: float = float("nan")
    fx_current: float = float("nan")

    implied_repo_bbg: float = float("nan")
    net_basis_bbg: float = float("nan")
    gross_basis_bbg: float = float("nan")


@dataclass(slots=True)
class FutureBook:
    contract_code: np.ndarray
    exchange: np.ndarray
    currency: np.ndarray
    portfolio: np.ndarray
    linked_isin: np.ndarray
    hedge_type: np.ndarray

    contracts: np.ndarray
    face_value: np.ndarray
    deliv_date: np.ndarray

    ctd_isin: np.ndarray
    ctd_cf: np.ndarray
    ctd_dirty_px_current: np.ndarray

    avg_entry_px: np.ndarray
    fut_px_prior: np.ndarray
    fut_px_current: np.ndarray
    fut_val_pt: np.ndarray
    fut_px_val_bp: np.ndarray

    fx_prior: np.ndarray
    fx_current: np.ndarray

    implied_repo_bbg: np.ndarray
    net_basis_bbg: np.ndarray
    gross_basis_bbg: np.ndarray

    positions: tuple[FuturePosition, ...] = ()

    _SPEC = (
        ("contract_code", "s"), ("exchange", "s"), ("currency", "s"),
        ("portfolio", "s"), ("linked_isin", "s"), ("hedge_type", "s"),
        ("contracts", "f"), ("face_value", "f"), ("deliv_date", "d"),
        ("ctd_isin", "s"), ("ctd_cf", "f"), ("ctd_dirty_px_current", "f"),
        ("avg_entry_px", "f"), ("fut_px_prior", "f"), ("fut_px_current", "f"),
        ("fut_val_pt", "f"), ("fut_px_val_bp", "f"),
        ("fx_prior", "f"), ("fx_current", "f"),
        ("implied_repo_bbg", "f"), ("net_basis_bbg", "f"), ("gross_basis_bbg", "f"),
    )

    @classmethod
    def from_positions(cls, positions: Iterable[FuturePosition]) -> "FutureBook":
        rows = tuple(positions)
        kwargs = {name: _column(rows, name, kind) for name, kind in cls._SPEC}
        for key in ("linked_isin", "ctd_isin", "currency"):
            kwargs[key] = np.array(
                [s.strip().upper() for s in kwargs[key].tolist()], dtype=object
            )
        return cls(positions=rows, **kwargs)

    def __len__(self) -> int:
        return int(self.contract_code.size)

    def __getitem__(self, index: int) -> FuturePosition:
        return self.positions[index]

    def __iter__(self):
        return iter(self.positions)

    @property
    def is_linked(self) -> np.ndarray:
        """A future with no LinkedISIN sits on no bond row.

        Its PnL is therefore absent from every bond-level total, which is the
        single largest thing the legacy attribution could not see.  It is
        measured here and reported as its own reconciliation line.
        """
        return np.array([bool(s) for s in self.linked_isin.tolist()])


# --------------------------------------------------------------------------- #
# swaps
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class SwapPosition:
    """One interest-rate swap.

    `swap_id_source` separates the two economically different populations the
    legacy sheet also distinguished:

        PLAIN      a swap actually in the book.  Its risk and its PnL are real.
        SYNTHETIC  the hedge the coverage relationship says SHOULD be on.  It is
                   a target, not a position: it must never contribute risk or
                   PnL to a total, only to the hedge-efficiency comparison.
    """

    deal_id: str
    currency: str = "EUR"
    portfolio: str = ""
    linked_isin: str = ""
    counterparty: str = ""
    swap_id_source: str = SWAP_PLAIN

    notional: float = float("nan")        # signed by pay_fixed, see PayFixed note
    fixed_rate: float = float("nan")      # percent
    float_index: str = ""
    float_spread: float = float("nan")    # percent
    start_date: _dt.date | None = None
    end_date: _dt.date | None = None
    pay_fixed: str = ""                   # "Y" -> pays fixed, receives float

    float_curve_type: str = ""            # ESTR / SOFR / EURIBOR / UNKNOWN
    dv01_supplied: float = float("nan")   # base ccy per bp, SIGNED
    npv_prior: float = float("nan")       # base ccy
    npv_current: float = float("nan")     # base ccy
    fx_prior: float = float("nan")
    fx_current: float = float("nan")

    @property
    def pays_fixed(self) -> bool:
        return self.pay_fixed.strip().upper().startswith("Y")


@dataclass(slots=True)
class SwapBook:
    deal_id: np.ndarray
    currency: np.ndarray
    portfolio: np.ndarray
    linked_isin: np.ndarray
    counterparty: np.ndarray
    swap_id_source: np.ndarray

    notional: np.ndarray
    fixed_rate: np.ndarray
    float_index: np.ndarray
    float_spread: np.ndarray
    start_date: np.ndarray
    end_date: np.ndarray
    pay_fixed: np.ndarray

    float_curve_type: np.ndarray
    dv01_supplied: np.ndarray
    npv_prior: np.ndarray
    npv_current: np.ndarray
    fx_prior: np.ndarray
    fx_current: np.ndarray

    positions: tuple[SwapPosition, ...] = ()

    _SPEC = (
        ("deal_id", "s"), ("currency", "s"), ("portfolio", "s"),
        ("linked_isin", "s"), ("counterparty", "s"), ("swap_id_source", "s"),
        ("notional", "f"), ("fixed_rate", "f"), ("float_index", "s"),
        ("float_spread", "f"), ("start_date", "d"), ("end_date", "d"),
        ("pay_fixed", "s"),
        ("float_curve_type", "s"), ("dv01_supplied", "f"),
        ("npv_prior", "f"), ("npv_current", "f"),
        ("fx_prior", "f"), ("fx_current", "f"),
    )

    @classmethod
    def from_positions(cls, positions: Iterable[SwapPosition]) -> "SwapBook":
        rows = tuple(positions)
        kwargs = {name: _column(rows, name, kind) for name, kind in cls._SPEC}
        for key in ("linked_isin", "currency", "swap_id_source", "float_curve_type"):
            kwargs[key] = np.array(
                [s.strip().upper() for s in kwargs[key].tolist()], dtype=object
            )
        kwargs["swap_id_source"] = np.array(
            [s if s in (SWAP_PLAIN, SWAP_SYNTHETIC) else SWAP_PLAIN
             for s in kwargs["swap_id_source"].tolist()],
            dtype=object,
        )
        # Classify the floating leg where the source file did not.
        family = kwargs["float_curve_type"]
        need = np.array([f in ("", "UNKNOWN") for f in family.tolist()])
        if need.any():
            derived = np.array(
                [
                    classify_float_family(idx, ccy)
                    for idx, ccy in zip(
                        kwargs["float_index"].tolist(), kwargs["currency"].tolist()
                    )
                ],
                dtype=object,
            )
            kwargs["float_curve_type"] = np.where(need, derived, family)
        return cls(positions=rows, **kwargs)

    def __len__(self) -> int:
        return int(self.deal_id.size)

    def __getitem__(self, index: int) -> SwapPosition:
        return self.positions[index]

    def __iter__(self):
        return iter(self.positions)

    @property
    def is_plain(self) -> np.ndarray:
        return self.swap_id_source == SWAP_PLAIN

    @property
    def is_synthetic(self) -> np.ndarray:
        return self.swap_id_source == SWAP_SYNTHETIC

    @property
    def is_linked(self) -> np.ndarray:
        return np.array([bool(s) for s in self.linked_isin.tolist()])

    @property
    def pay_fixed_sign(self) -> np.ndarray:
        """-1 when the swap pays fixed, +1 when it receives.

        A payer swap is short duration - it loses when rates fall - so its risk
        must carry the opposite sign to a long bond position for the two to net.
        """
        return np.array(
            [-1.0 if str(v).strip().upper().startswith("Y") else 1.0
             for v in self.pay_fixed.tolist()],
            dtype=np.float64,
        )

    def year_fraction_to_end(self, as_of: np.datetime64, basis: float = 365.0) -> np.ndarray:
        days = (
            self.end_date.astype("datetime64[D]").astype("int64")
            - np.datetime64(as_of, "D").astype("int64")
        ).astype(np.float64)
        return np.where(np.isnat(self.end_date), np.nan, np.maximum(days, 0.0) / basis)


# --------------------------------------------------------------------------- #
# floating-index classification
# --------------------------------------------------------------------------- #


def classify_float_family(rate_index: object, currency: object) -> str:
    """Map a floating-rate index onto the curve that projects it.

    Which curve a swap's floating leg references decides which rate move its
    hedge PnL is measured against, so getting this wrong attributes an ESTR
    swap's PnL to the EURIBOR curve and books the OIS/IBOR basis as model error.

        ESTR / SOFR  overnight-index swaps -> the OIS curve
        EURIBOR      IBOR-projection swaps -> the swap curve
        UNKNOWN      classification failed; the row is flagged, not guessed

    UNKNOWN is a real answer.  Defaulting an unclassifiable index to the OIS
    curve would produce a plausible-looking hedge PnL from an assumption nobody
    made.
    """
    s = str(rate_index or "").strip().upper()
    for ch in (" ", "-", "_"):
        s = s.replace(ch, "")
    ccy = str(currency or "").strip().upper()

    if not s:
        return "UNKNOWN"

    if ccy == "USD" and any(t in s for t in ("SOFR", "OIS")):
        return "SOFR"
    if ccy == "GBP" and any(t in s for t in ("SONIA", "SONIO", "OIS")):
        return "SONIA"
    if ccy == "EUR":
        if any(t in s for t in ("ESTR", "€STR", "EONIA", "OIS")):
            return "ESTR"
        if any(t in s for t in ("EURIBOR", "EURIB", "EUR003M", "EUR006M", "EUR012M")):
            return "EURIBOR"
    return "UNKNOWN"


def float_family_status(rate_index: object, currency: object, family: str) -> str:
    """Flag a floating index that does not belong to its swap's currency.

    A EUR swap referencing SOFR is a data error, not an exotic trade.  It is
    reported rather than classified, because whichever curve it were assigned to
    would be the wrong one.
    """
    s = str(rate_index or "").strip().upper()
    for ch in (" ", "-", "_"):
        s = s.replace(ch, "")
    ccy = str(currency or "").strip().upper()

    if not s:
        return "UNKNOWN_SWAP_FAMILY"
    if ccy == "EUR" and "SOFR" in s:
        return "SWAP_CCY_INDEX_MISMATCH"
    if ccy == "USD" and any(
        t in s for t in ("EURIBOR", "EURIB", "EUR003M", "EUR006M", "EUR012M", "ESTR", "EONIA")
    ):
        return "SWAP_CCY_INDEX_MISMATCH"
    if not family or family == "UNKNOWN":
        return "UNKNOWN_SWAP_FAMILY"
    return "OK"
