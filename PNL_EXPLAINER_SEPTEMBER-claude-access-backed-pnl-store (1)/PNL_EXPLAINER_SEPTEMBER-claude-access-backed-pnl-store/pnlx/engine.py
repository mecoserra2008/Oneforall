"""The attribution engines: bonds, futures, swaps.

THE BRIDGE
==========
Two totals are computed for every position, and the whole report is the gap
between them.

    PRACTICAL PnL   what the book actually made, from marks and cash:

        practical = change in dirty market value (base ccy)
                  + coupon cash actually received in the period
                  + actual hedge PnL (futures variation margin + swap NPV move)

    THEORETICAL PnL what the risk model says it should have made:

        theoretical = duration      (framework chain, ties to -DV01 * dy)
                    + convexity     (second order in the same yield move)
                    + carry         (coupon accrual + pull to par)
                    + FX            (translation of the opening position)
                    + hedge model   (hedge DV01 against the curve it hedges)

    EXPLAINED = theoretical + hedge basis
    RESIDUAL  = practical - explained

`hedge basis` is `actual hedge PnL - model hedge PnL`.  Adding it back means
the hedge leg of `explained` telescopes to the ACTUAL hedge PnL, which is what
`practical` contains - so the bridge closes by construction and the residual
isolates BOND-LEG model error.  That is what "unexplained" should mean.

Leaving the basis out - which the legacy sheet did before it was corrected -
guarantees that the futures/CTD basis and the swap spread mismatch land in the
residual every single day, however good the model is.

The basis is still reported in its own right, because a desk running a
cross-instrument hedge wants to watch it.

WHAT WAS CORRECTED FROM THE WORKBOOK
====================================
Each of these changes a number.  The reasoning is in docs/MODEL.md.

1.  RISK IS STRUCK AT THE START OF THE PERIOD.
    The workbook built DV01 from the T0 dirty price, so the risk used to
    explain a move was the risk AFTER that move.  Its own comments measured the
    cost: on a self-consistent test book a 10bp day left +288 EUR of residual
    on -66,170 EUR of duration PnL, of which +278 (96.5%) was this convention
    alone.  Attribution DV01 is now struck at T-1; `dv01_timing_bias` reports
    what the old convention would have added, so the change is auditable rather
    than merely asserted.
    Hedging metrics still use CURRENT risk - that is the risk to be hedged
    tomorrow, and the hedge ratio rightly measures it.

2.  FX IS CONSISTENT END TO END.
    The workbook translated rate legs at FX_T0 and the FX leg at FX_T-1, which
    opens a standing residual on every non-base position.  Every model leg here
    is struck at the OPENING fix, the FX leg is the translation of the opening
    position, and the second-order term (local move x FX move) is booked
    explicitly instead of falling into the residual.

        d(MV_base) = MV_local_prior * dFX          <- pnl_fx
                   + d(MV_local) * FX_prior        <- the rate/spread/carry legs
                   + d(MV_local) * dFX             <- pnl_fx_cross

    an exact identity, so nothing is left over.

3.  COUPON PAYMENTS ARE HANDLED.
    The workbook accrued coupon smoothly and had no term for a coupon actually
    PAID inside the period.  On an ex-coupon date the dirty price drops by the
    coupon and nothing on the explained side put the cash back, so that bond
    showed a one-day residual the size of its coupon.  Carry is now
    `accrued(T0) - accrued(T-1) + coupons paid`, and the cash appears on the
    practical side too, so the bridge closes across a payment date.

4.  ONE SIGN DRIVER.
    Signed notional, and nothing multiplies by a side flag a second time.

SIGN AND UNIT CONVENTIONS
=========================
    DV01 is signed by the position.  A long bond has positive DV01 and loses
    when yields rise, so every first-order leg is `-DV01 * delta_bp`.
    A payer swap and a short future carry negative DV01 and therefore net
    against the bond without any special-casing.

    Deltas are in BASIS POINTS.  Curves and yields are in PERCENT.  DV01 is
    base currency per basis point.  So `-DV01 * delta_bp` is base currency.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from . import bondmath, daycount as dc
from .columns import Column, ColumnFrame, Kind
from .config import AppConfig
from .curves import CurveSet
from .frameworks import CHAINS, FrameworkResolution, curve_for_framework, resolve_frameworks
from .instruments import BondBook, FutureBook, SwapBook, float_family_status
from .loaders import InputBundle

__all__ = [
    "AttributionEngine",
    "AttributionResult",
    "HedgeLinkage",
    "BOND_SCHEMA",
    "FUTURES_SCHEMA",
    "SWAPS_SCHEMA",
]

_DAY = np.dtype("datetime64[D]")
_OIS_FAMILIES = ("ESTR", "SOFR", "SONIA")
_IBOR_FAMILIES = ("EURIBOR",)


def _nan(n: int) -> np.ndarray:
    return np.full(n, np.nan, dtype=np.float64)


def _both(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    """Elementwise 'both are real numbers'."""
    return np.isfinite(a) & np.isfinite(b)


def _diff_bp(current: np.ndarray, prior: np.ndarray, scale: float = 1.0) -> np.ndarray:
    """Change between two snapshots, NaN unless BOTH ends are numbers.

    Requiring both ends is the whole point.  Treating a missing prior mark as
    zero manufactures a move the size of the current level, which then flows
    into a spread PnL that looks real.
    """
    return np.where(_both(current, prior), (current - prior) * scale, np.nan)


def _first_order(dv01: np.ndarray, delta_bp: np.ndarray) -> np.ndarray:
    """-DV01 * delta_bp, blank unless both are known."""
    return np.where(_both(dv01, delta_bp), -dv01 * delta_bp, np.nan)


def _sum_by_key(
    keys: np.ndarray, values: np.ndarray, index: dict[str, int], size: int
) -> np.ndarray:
    """Scatter-add `values` into `size` buckets chosen by `keys`.

    This replaces the workbook's per-row SUMIFS over the hedge sheets: one
    `np.add.at` for the whole book instead of one range scan per bond per
    hedge column.  NaN contributions are skipped, so one unpriced hedge does
    not blank a bond's whole hedge total - the count columns say how many
    hedges were found, so a partial sum is visible rather than implied.
    """
    out = np.zeros(size, dtype=np.float64)
    if keys.size == 0:
        return out
    positions = np.array([index.get(str(k), -1) for k in keys.tolist()], dtype=np.int64)
    vals = np.asarray(values, dtype=np.float64)
    ok = (positions >= 0) & np.isfinite(vals)
    if ok.any():
        np.add.at(out, positions[ok], vals[ok])
    return out


def _count_by_key(keys: np.ndarray, index: dict[str, int], size: int) -> np.ndarray:
    out = np.zeros(size, dtype=np.float64)
    if keys.size == 0:
        return out
    positions = np.array([index.get(str(k), -1) for k in keys.tolist()], dtype=np.int64)
    ok = positions >= 0
    if ok.any():
        np.add.at(out, positions[ok], 1.0)
    return out


# --------------------------------------------------------------------------- #
# hedge linkage
# --------------------------------------------------------------------------- #


@dataclass(slots=True)
class HedgeLinkage:
    """Hedge risk and PnL aggregated onto the bonds they are linked to.

    Hedges attach to bonds by ISIN and nothing else, which has two consequences
    the report must state rather than hide:

      * a hedge with a blank LinkedISIN sits on NO bond row, so its PnL is
        absent from every bond-level total;
      * if the same ISIN appears on two bond rows - the same bond held in two
        portfolios - BOTH rows claim the full hedge PnL and DV01 of that ISIN.

    `unlinked_*` measures the first and `BondBook.duplicate_isin_mask` the
    second.  Neither is fixed here, because neither can be fixed by arithmetic:
    they are data questions.  They are quantified so a total is never presented
    as complete when it is not.
    """

    futures_dv01: np.ndarray
    futures_dv01_prior: np.ndarray
    futures_pnl: np.ndarray
    futures_count: np.ndarray

    plain_swap_dv01: np.ndarray
    plain_swap_dv01_ois: np.ndarray
    plain_swap_dv01_ibor: np.ndarray
    plain_swap_dv01_unknown: np.ndarray
    plain_swap_pnl: np.ndarray
    plain_swap_count: np.ndarray

    synthetic_swap_dv01: np.ndarray
    synthetic_swap_count: np.ndarray

    unlinked_futures_pnl: float = 0.0
    unlinked_futures_dv01: float = 0.0
    unlinked_futures_count: int = 0
    unlinked_swap_pnl: float = 0.0
    unlinked_swap_dv01: float = 0.0
    unlinked_swap_count: int = 0
    orphan_futures_pnl: float = 0.0
    orphan_futures_count: int = 0
    orphan_swap_pnl: float = 0.0
    orphan_swap_count: int = 0


# --------------------------------------------------------------------------- #
# result
# --------------------------------------------------------------------------- #


@dataclass(slots=True)
class AttributionResult:
    """Everything one run produced."""

    bonds: ColumnFrame
    futures: ColumnFrame
    swaps: ColumnFrame
    linkage: HedgeLinkage
    frameworks: FrameworkResolution
    config: AppConfig
    curves: CurveSet
    diagnostics: dict[str, object] = field(default_factory=dict)

    @property
    def inputs_curves_described(self) -> list[dict[str, object]]:
        """Flat rows describing every curve the run used, for the report."""
        return self.curves.describe()


# --------------------------------------------------------------------------- #
# the engine
# --------------------------------------------------------------------------- #


class AttributionEngine:
    """Runs the whole attribution over a loaded input bundle."""

    def __init__(self, config: AppConfig, inputs: InputBundle) -> None:
        self.config = config
        self.inputs = inputs
        self.prior = np.datetime64(config.run.prior, "D")
        self.current = np.datetime64(config.run.as_of, "D")
        self.year_fraction = config.run.year_fraction
        self.days = config.run.days

    # -- public ------------------------------------------------------------- #

    def run(self) -> AttributionResult:
        bonds = self.inputs.bonds
        n = len(bonds)

        fx_prior, fx_current = self._fx(bonds)
        market = self._bond_market(bonds, fx_prior, fx_current)
        deltas = self._bond_deltas(bonds, market)
        linkage = self._link_hedges(bonds, market, deltas)
        frameworks = self._resolve_frameworks(bonds, deltas, linkage)
        legs = self._bond_legs(bonds, market, deltas, linkage, frameworks)

        bond_frame = self._bond_frame(bonds, market, deltas, linkage, frameworks, legs)
        futures_frame = self._futures_frame(bonds, market)
        swaps_frame = self._swaps_frame()

        diagnostics = self._diagnostics(bonds, bond_frame, linkage)

        return AttributionResult(
            bonds=bond_frame,
            futures=futures_frame,
            swaps=swaps_frame,
            linkage=linkage,
            frameworks=frameworks,
            config=self.config,
            curves=self.inputs.curves,
            diagnostics=diagnostics,
        )

    # -- FX ------------------------------------------------------------------ #

    def _fx(self, bonds: BondBook) -> tuple[np.ndarray, np.ndarray]:
        """Per-position FX, preferring the position file and falling back to the
        FX table.  The base currency is pinned at 1.0 on both dates."""
        base = self.config.run.base_ccy
        table = self.inputs.fx

        prior = np.asarray(bonds.fx_prior, dtype=np.float64).copy()
        current = np.asarray(bonds.fx_current, dtype=np.float64).copy()

        for i, ccy in enumerate(bonds.currency.tolist()):
            if ccy == base:
                prior[i] = 1.0
                current[i] = 1.0
                continue
            pair = table.get(ccy)
            if pair is None:
                continue
            if not np.isfinite(prior[i]):
                prior[i] = pair[0]
            if not np.isfinite(current[i]):
                current[i] = pair[1]
        return prior, current

    # -- bond market values and risk ----------------------------------------- #

    def _bond_market(
        self, bonds: BondBook, fx_prior: np.ndarray, fx_current: np.ndarray
    ) -> dict[str, np.ndarray]:
        n = len(bonds)
        cfg = self.config
        notional = np.asarray(bonds.notional, dtype=np.float64)
        coupon = bonds.coupon_decimal
        freq = np.asarray(bonds.coupon_freq, dtype=np.float64)
        code = np.asarray(bonds.day_count_code, dtype=np.float64)
        maturity = bonds.maturity

        prior_d = np.full(n, self.prior, dtype=_DAY)
        current_d = np.full(n, self.current, dtype=_DAY)

        # --- market values -------------------------------------------------- #
        mv_prior_local = notional * bonds.dirty_px_prior / 100.0
        mv_current_local = notional * bonds.dirty_px_current / 100.0
        mv_prior_base = mv_prior_local * fx_prior
        mv_current_base = mv_current_local * fx_current
        delta_mv_base = np.where(
            _both(mv_current_base, mv_prior_base), mv_current_base - mv_prior_base, np.nan
        )

        # --- duration, filled in from the model where the feed is silent ---- #
        model_dur_prior, model_convexity_prior = bondmath.yield_risk(
            prior_d, maturity, coupon, bonds.ytm_prior / 100.0, freq, code,
            bump_bp=cfg.model.convexity_bump_bp,
            max_flows=cfg.model.pull_to_par.max_cashflows,
        )
        model_dur_current, _ = bondmath.yield_risk(
            current_d, maturity, coupon, bonds.ytm_current / 100.0, freq, code,
            bump_bp=cfg.model.convexity_bump_bp,
            max_flows=cfg.model.pull_to_par.max_cashflows,
        )

        # Spread duration: prefer the OAS duration where it is quoted, since a
        # spread move is what it measures the sensitivity to.  Otherwise the
        # ordinary modified duration, and only then the model's own.
        dur_current = _coalesce(
            np.abs(bonds.oas_mod_duration_current),
            np.abs(bonds.mod_duration_current),
            np.abs(model_dur_current),
        )
        dur_prior = _coalesce(
            np.abs(bonds.mod_duration_prior),
            np.abs(model_dur_prior),
            dur_current,
        )
        duration_source = np.where(
            np.isfinite(bonds.mod_duration_prior),
            "feed T-1",
            np.where(np.isfinite(model_dur_prior), "model T-1", "fallback T0"),
        ).astype(object)

        # --- convexity ------------------------------------------------------ #
        if cfg.model.convexity_source == "supplied":
            convexity = np.asarray(bonds.convexity, dtype=np.float64)
        elif cfg.model.convexity_source == "analytic":
            _, convexity = bondmath.yield_risk(
                prior_d, maturity, coupon, bonds.ytm_prior / 100.0, freq, code,
                bump_bp=0.01, max_flows=cfg.model.pull_to_par.max_cashflows,
            )
        else:
            convexity = model_convexity_prior
        convexity = _coalesce(convexity, np.asarray(bonds.convexity, dtype=np.float64))

        # --- DV01 ----------------------------------------------------------- #
        # DV01 per unit nominal = duration x dirty price / 100 x fx x 1bp.
        # Struck twice, on purpose: the PRIOR figure explains the period, the
        # CURRENT figure is the risk to hedge.
        dv01_unit_prior = (
            dur_prior * np.abs(bonds.dirty_px_prior) / 100.0 * np.abs(fx_prior) * 1e-4
        )
        dv01_unit_current = (
            dur_current * np.abs(bonds.dirty_px_current) / 100.0 * np.abs(fx_current) * 1e-4
        )
        dv01_prior = notional * dv01_unit_prior
        dv01_current = notional * dv01_unit_current

        attribution_dv01 = (
            dv01_prior if cfg.model.attribution_risk_date == "prior" else dv01_current
        )

        # --- accrued and coupon cash ---------------------------------------- #
        accrued_prior = dc.accrued_interest(prior_d, maturity, coupon, freq, code)
        accrued_current = dc.accrued_interest(current_d, maturity, coupon, freq, code)
        coupons_paid = dc.coupons_paid_between(
            prior_d, current_d, maturity, coupon, freq,
            max_flows=cfg.model.pull_to_par.max_cashflows,
        )

        return {
            "notional": notional,
            "fx_prior": fx_prior,
            "fx_current": fx_current,
            "mv_prior_local": mv_prior_local,
            "mv_current_local": mv_current_local,
            "mv_prior_base": mv_prior_base,
            "mv_current_base": mv_current_base,
            "delta_mv_base": delta_mv_base,
            "duration_prior": dur_prior,
            "duration_current": dur_current,
            "duration_source": duration_source,
            "model_duration_prior": model_dur_prior,
            "convexity": convexity,
            "dv01_unit_prior": dv01_unit_prior,
            "dv01_unit_current": dv01_unit_current,
            "dv01_prior": dv01_prior,
            "dv01_current": dv01_current,
            "attribution_dv01": attribution_dv01,
            "accrued_prior": accrued_prior,
            "accrued_current": accrued_current,
            "coupons_paid": coupons_paid,
            "years_prior": bonds.years_to_maturity(self.prior),
            "years_current": bonds.years_to_maturity(self.current),
        }

    # -- curve deltas -------------------------------------------------------- #

    def _bond_deltas(self, bonds: BondBook, market: dict[str, np.ndarray]) -> dict[str, np.ndarray]:
        """Every rate and spread move the attribution needs, in basis points.

        The curve legs are read at the bond's own tenor on the PRIOR snapshot's
        tenor, so both ends of each difference are the same point on the curve
        moving - not two different points on two different days, which would
        book roll-down as a curve move.
        """
        curves: CurveSet = self.inputs.curves
        years = market["years_prior"]
        ccy = bonds.currency

        ois_prior = curves.rates_for(ccy, years, "OIS", "prior")
        ois_current = curves.rates_for(ccy, years, "OIS", "current")
        gov_prior = curves.rates_for(ccy, years, "GOV", "prior")
        gov_current = curves.rates_for(ccy, years, "GOV", "current")
        swap_prior = curves.rates_for(ccy, years, "SWAP", "prior")
        swap_current = curves.rates_for(ccy, years, "SWAP", "current")

        # y = r + g + q + i, every term a difference of two quoted rates.
        g_prior = np.where(_both(gov_prior, ois_prior), gov_prior - ois_prior, np.nan)
        g_current = np.where(_both(gov_current, ois_current), gov_current - ois_current, np.nan)
        q_prior = np.where(_both(swap_prior, gov_prior), swap_prior - gov_prior, np.nan)
        q_current = np.where(_both(swap_current, gov_current), swap_current - gov_current, np.nan)

        ispread_prior = _diff_bp(bonds.ytm_prior, swap_prior, 100.0)
        ispread_current = _diff_bp(bonds.ytm_current, swap_current, 100.0)
        gspread_prior = _diff_bp(bonds.ytm_prior, gov_prior, 100.0)
        gspread_current = _diff_bp(bonds.ytm_current, gov_current, 100.0)

        return {
            "ois_prior": ois_prior,
            "ois_current": ois_current,
            "gov_prior": gov_prior,
            "gov_current": gov_current,
            "swap_prior": swap_prior,
            "swap_current": swap_current,
            "g_prior": g_prior,
            "g_current": g_current,
            "q_prior": q_prior,
            "q_current": q_current,
            "ispread_prior": ispread_prior,
            "ispread_current": ispread_current,
            "gspread_prior": gspread_prior,
            "gspread_current": gspread_current,
            "delta_y_bp": _diff_bp(bonds.ytm_current, bonds.ytm_prior, 100.0),
            "delta_r_bp": _diff_bp(ois_current, ois_prior, 100.0),
            "delta_gov_bp": _diff_bp(gov_current, gov_prior, 100.0),
            "delta_swap_bp": _diff_bp(swap_current, swap_prior, 100.0),
            "delta_g_bp": _diff_bp(g_current, g_prior, 100.0),
            "delta_q_bp": _diff_bp(q_current, q_prior, 100.0),
            "delta_ispread_bp": _diff_bp(ispread_current, ispread_prior),
            "delta_gspread_bp": _diff_bp(gspread_current, gspread_prior),
            "delta_zspread_bp": _diff_bp(bonds.zspread_current, bonds.zspread_prior),
            "delta_asw_bp": _diff_bp(bonds.asw_current, bonds.asw_prior),
            "delta_oas_bp": _diff_bp(bonds.oas_current, bonds.oas_prior),
        }

    # -- hedges -------------------------------------------------------------- #

    def _link_hedges(
        self,
        bonds: BondBook,
        market: dict[str, np.ndarray],
        deltas: dict[str, np.ndarray],
    ) -> HedgeLinkage:
        n = len(bonds)
        futures = self.inputs.futures
        swaps = self.inputs.swaps

        # First occurrence wins for the index; a repeated ISIN is reported
        # separately rather than silently splitting its hedges.
        index: dict[str, int] = {}
        for i, key in enumerate(bonds.isin.tolist()):
            index.setdefault(str(key), i)

        # ---- futures ------------------------------------------------------- #
        f_dv01_cur, f_dv01_prior = self._futures_dv01(futures)
        f_pnl = self._futures_actual_pnl(futures)

        fut_keys = futures.linked_isin if len(futures) else np.array([], dtype=object)
        futures_dv01 = _sum_by_key(fut_keys, f_dv01_cur, index, n)
        futures_dv01_prior = _sum_by_key(fut_keys, f_dv01_prior, index, n)
        futures_pnl = _sum_by_key(fut_keys, f_pnl, index, n)
        futures_count = _count_by_key(fut_keys, index, n)

        # ---- swaps --------------------------------------------------------- #
        s_dv01 = self._swap_dv01(swaps)
        s_pnl = self._swap_actual_pnl(swaps)

        if len(swaps):
            plain = swaps.is_plain
            synthetic = swaps.is_synthetic
            family = np.array([str(f) for f in swaps.float_curve_type.tolist()], dtype=object)
            is_ois = np.isin(family, np.array(_OIS_FAMILIES, dtype=object))
            is_ibor = np.isin(family, np.array(_IBOR_FAMILIES, dtype=object))
            is_unknown = ~(is_ois | is_ibor)

            swap_keys = swaps.linked_isin
            plain_swap_dv01 = _sum_by_key(swap_keys, np.where(plain, s_dv01, np.nan), index, n)
            plain_swap_dv01_ois = _sum_by_key(
                swap_keys, np.where(plain & is_ois, s_dv01, np.nan), index, n
            )
            plain_swap_dv01_ibor = _sum_by_key(
                swap_keys, np.where(plain & is_ibor, s_dv01, np.nan), index, n
            )
            plain_swap_dv01_unknown = _sum_by_key(
                swap_keys, np.where(plain & is_unknown, s_dv01, np.nan), index, n
            )
            plain_swap_pnl = _sum_by_key(swap_keys, np.where(plain, s_pnl, np.nan), index, n)
            plain_swap_count = _count_by_key(swap_keys[plain], index, n)

            synthetic_swap_dv01 = _sum_by_key(
                swap_keys, np.where(synthetic, s_dv01, np.nan), index, n
            )
            synthetic_swap_count = _count_by_key(swap_keys[synthetic], index, n)
        else:
            zeros = np.zeros(n)
            plain_swap_dv01 = zeros.copy()
            plain_swap_dv01_ois = zeros.copy()
            plain_swap_dv01_ibor = zeros.copy()
            plain_swap_dv01_unknown = zeros.copy()
            plain_swap_pnl = zeros.copy()
            plain_swap_count = zeros.copy()
            synthetic_swap_dv01 = zeros.copy()
            synthetic_swap_count = zeros.copy()
            plain = synthetic = np.array([], dtype=bool)

        # ---- what the bond rows cannot see --------------------------------- #
        known = set(index)

        def _leak(book, values, dv01_values, linked_mask):
            unlinked = ~linked_mask
            orphan = np.array(
                [bool(k) and str(k) not in known for k in book.linked_isin.tolist()]
            ) if len(book) else np.array([], dtype=bool)
            return (
                float(np.nansum(values[unlinked])) if len(book) else 0.0,
                float(np.nansum(dv01_values[unlinked])) if len(book) else 0.0,
                int(unlinked.sum()) if len(book) else 0,
                float(np.nansum(values[orphan])) if len(book) else 0.0,
                int(orphan.sum()) if len(book) else 0,
            )

        (
            unlinked_fut_pnl, unlinked_fut_dv01, unlinked_fut_n,
            orphan_fut_pnl, orphan_fut_n,
        ) = _leak(futures, f_pnl, f_dv01_cur, futures.is_linked if len(futures) else np.array([], bool))

        if len(swaps):
            plain_only = swaps.is_plain
            (
                unlinked_sw_pnl, unlinked_sw_dv01, unlinked_sw_n,
                orphan_sw_pnl, orphan_sw_n,
            ) = _leak(
                swaps,
                np.where(plain_only, s_pnl, np.nan),
                np.where(plain_only, s_dv01, np.nan),
                swaps.is_linked,
            )
        else:
            unlinked_sw_pnl = unlinked_sw_dv01 = orphan_sw_pnl = 0.0
            unlinked_sw_n = orphan_sw_n = 0

        return HedgeLinkage(
            futures_dv01=futures_dv01,
            futures_dv01_prior=futures_dv01_prior,
            futures_pnl=futures_pnl,
            futures_count=futures_count,
            plain_swap_dv01=plain_swap_dv01,
            plain_swap_dv01_ois=plain_swap_dv01_ois,
            plain_swap_dv01_ibor=plain_swap_dv01_ibor,
            plain_swap_dv01_unknown=plain_swap_dv01_unknown,
            plain_swap_pnl=plain_swap_pnl,
            plain_swap_count=plain_swap_count,
            synthetic_swap_dv01=synthetic_swap_dv01,
            synthetic_swap_count=synthetic_swap_count,
            unlinked_futures_pnl=unlinked_fut_pnl,
            unlinked_futures_dv01=unlinked_fut_dv01,
            unlinked_futures_count=unlinked_fut_n,
            unlinked_swap_pnl=unlinked_sw_pnl,
            unlinked_swap_dv01=unlinked_sw_dv01,
            unlinked_swap_count=unlinked_sw_n,
            orphan_futures_pnl=orphan_fut_pnl,
            orphan_futures_count=orphan_fut_n,
            orphan_swap_pnl=orphan_sw_pnl,
            orphan_swap_count=orphan_sw_n,
        )

    def _futures_dv01(self, futures: FutureBook) -> tuple[np.ndarray, np.ndarray]:
        """Signed cash DV01 per futures position, at both FX fixes.

        Bloomberg's FUT_PX_VAL_BP is quoted in PRICE POINTS per basis point
        (about 0.0605 for a Bund), so the cash figure needs a further multiply
        by FUT_VAL_PT - EUR 1000 per full point - giving roughly EUR 60 per
        contract per bp.  A feed that already reports cash is handled by the
        `futures.dv01_includes_point_value` switch.

        Desk check: `futures_dv01 / contracts` should read 60-90 for a Bund.
        A figure near 60,000 means the switch is set the wrong way, and the
        symptom is every futures-hedged bond reading "Over-hedged".
        """
        if len(futures) == 0:
            empty = np.array([], dtype=np.float64)
            return empty, empty
        contracts = np.asarray(futures.contracts, dtype=np.float64)
        unit = np.asarray(futures.fut_px_val_bp, dtype=np.float64)
        point = np.asarray(futures.fut_val_pt, dtype=np.float64)
        scale = 1.0 if self.config.futures.dv01_includes_point_value else point
        base = contracts * unit * scale
        return base * futures.fx_current, base * futures.fx_prior

    def _futures_actual_pnl(self, futures: FutureBook) -> np.ndarray:
        """Variation margin: contracts x point value x price move x FX.

        A future has no market value to revalue - it is margined daily - so its
        PnL IS the price move, and there is no separate FX translation of a
        position that does not exist.
        """
        if len(futures) == 0:
            return np.array([], dtype=np.float64)
        move = _diff_bp(futures.fut_px_current, futures.fut_px_prior)
        return (
            np.asarray(futures.contracts, dtype=np.float64)
            * np.asarray(futures.fut_val_pt, dtype=np.float64)
            * move
            * np.asarray(futures.fx_current, dtype=np.float64)
        )

    def _swap_dv01(self, swaps: SwapBook) -> np.ndarray:
        """Signed swap DV01 in base currency per basis point.

        The supplied figure is used where the swap system publishes one; it
        prices the real payment schedule, which the fallback below does not.

        The fallback is a single-annuity approximation:

            DV01 = sign(pay fixed) * |notional| * 1bp * annuity * FX
            annuity = (1 - DF) / z          z in DECIMAL

        The legacy workbook divided by the rate in PERCENTAGE POINTS here,
        which made the model swap DV01 a hundred times too small.  It went
        unnoticed because the sheet used the Bloomberg BPV instead and the
        model figure was never read.  It is fixed rather than reproduced,
        because a fallback that is only ever wrong is worse than no fallback.
        """
        if len(swaps) == 0:
            return np.array([], dtype=np.float64)

        supplied = np.asarray(swaps.dv01_supplied, dtype=np.float64)
        if self.config.swaps.dv01_source == "supplied":
            model = self._swap_model_dv01(swaps)
            return np.where(np.isfinite(supplied), supplied, model)
        return self._swap_model_dv01(swaps)

    def _swap_model_dv01(self, swaps: SwapBook) -> np.ndarray:
        years = swaps.year_fraction_to_end(self.current)
        curves = self.inputs.curves
        rate_pct = curves.rates_for(swaps.currency, years, "OIS", "current")
        z = rate_pct / 100.0

        with np.errstate(divide="ignore", invalid="ignore"):
            df = np.where(1.0 + z > 0, (1.0 + z) ** (-years), np.nan)
            # (1 - DF) / z is the level of a unit annuity; at z -> 0 it tends to
            # the tenor itself, which is the right limit and avoids a blow-up.
            annuity = np.where(np.abs(z) > 1e-9, (1.0 - df) / z, years)

        fx = _coalesce(
            np.asarray(swaps.fx_current, dtype=np.float64),
            np.ones(len(swaps)),
        )
        notional = np.abs(np.asarray(swaps.notional, dtype=np.float64))
        return swaps.pay_fixed_sign * notional * 1e-4 * annuity * fx

    def _swap_actual_pnl(self, swaps: SwapBook) -> np.ndarray:
        """NPV move, in base currency.  Already a PnL; nothing to translate."""
        if len(swaps) == 0:
            return np.array([], dtype=np.float64)
        return _diff_bp(swaps.npv_current, swaps.npv_prior)

    # -- frameworks ---------------------------------------------------------- #

    def _resolve_frameworks(
        self,
        bonds: BondBook,
        deltas: dict[str, np.ndarray],
        linkage: HedgeLinkage,
    ) -> FrameworkResolution:
        available = {
            "G": np.isfinite(deltas["delta_gspread_bp"]),
            "I": np.isfinite(deltas["delta_ispread_bp"]),
            "ASW": np.isfinite(deltas["delta_asw_bp"]),
            "Z": np.isfinite(deltas["delta_zspread_bp"]),
            "OAS": np.isfinite(deltas["delta_oas_bp"]),
        }
        return resolve_frameworks(
            bonds.isin,
            linkage.futures_dv01,
            linkage.plain_swap_dv01,
            available,
            global_override=self.config.model.spread_framework_override,
            per_bond_override=self.inputs.override_codes,
            override_reasons=self.inputs.override_notes,
        )

    # -- the legs ------------------------------------------------------------ #

    def _bond_legs(
        self,
        bonds: BondBook,
        market: dict[str, np.ndarray],
        deltas: dict[str, np.ndarray],
        linkage: HedgeLinkage,
        frameworks: FrameworkResolution,
    ) -> dict[str, np.ndarray]:
        cfg = self.config
        n = len(bonds)
        d = market["attribution_dv01"]
        legs: dict[str, np.ndarray] = {}

        # ---- first-order curve and spread legs ----------------------------- #
        legs["pnl_ois"] = _first_order(d, deltas["delta_r_bp"])
        legs["pnl_gov_basis"] = _first_order(d, deltas["delta_g_bp"])
        legs["pnl_swap_gov_basis"] = _first_order(d, deltas["delta_q_bp"])
        legs["pnl_ispread"] = _first_order(d, deltas["delta_ispread_bp"])
        legs["pnl_gspread"] = _first_order(d, deltas["delta_gspread_bp"])
        legs["pnl_zspread"] = _first_order(d, deltas["delta_zspread_bp"])
        legs["pnl_asw"] = _first_order(d, deltas["delta_asw_bp"])
        legs["pnl_oas"] = _first_order(d, deltas["delta_oas_bp"])
        legs["pnl_yield_only"] = _first_order(d, deltas["delta_y_bp"])

        # For OIS/SOFR there is no explicit spread leg, so the spread is imputed
        # as the part of the yield move the OIS curve did not explain.
        spread_over_ois = np.where(
            np.isfinite(d) & _both(deltas["delta_y_bp"], deltas["delta_r_bp"]),
            -d * (deltas["delta_y_bp"] - deltas["delta_r_bp"]),
            np.nan,
        )
        legs["pnl_spread_over_ois"] = spread_over_ois

        # ---- chain sums, one framework at a time --------------------------- #
        duration_total = np.full(n, np.nan)
        spread_used = np.full(n, np.nan)
        code = frameworks.code

        for name, chain in CHAINS.items():
            mask = code == name
            if not mask.any():
                continue
            if name == "MIXED":
                # DV01-weighted blend: a bond hedged half by futures and half by
                # swaps is attributed half on the government chain and half on
                # the swap chain.  There is no honest per-leg split for such a
                # row, which is why the report shows its duration total in a
                # "not separately split" line rather than pretending otherwise.
                g_total = _chain_sum(legs, CHAINS["G"].legs, n)
                i_total = _chain_sum(legs, CHAINS["I"].legs, n)
                w_f = frameworks.futures_weight
                w_s = frameworks.swap_weight
                blended = np.where(
                    _both(g_total, i_total) & ((w_f + w_s) > 0),
                    w_f * g_total + w_s * i_total,
                    np.nan,
                )
                duration_total = np.where(mask, blended, duration_total)
                spread_used = np.where(
                    mask,
                    np.where(
                        _both(legs["pnl_gspread"], legs["pnl_ispread"]),
                        w_f * legs["pnl_gspread"] + w_s * legs["pnl_ispread"],
                        np.nan,
                    ),
                    spread_used,
                )
            elif name == "REVIEW":
                duration_total = np.where(mask, np.nan, duration_total)
                spread_used = np.where(mask, np.nan, spread_used)
            else:
                duration_total = np.where(mask, _chain_sum(legs, chain.legs, n), duration_total)
                if chain.spread_leg:
                    spread_used = np.where(mask, legs[chain.spread_leg], spread_used)

        legs["pnl_duration_total"] = duration_total
        legs["spread_pnl_used"] = spread_used

        # The chain must reproduce -DV01 * dy.  This column is the difference.
        # Read it as data quality, not model quality: G and I tie exactly by
        # construction, so anything non-zero there means a curve leg is stale
        # or T0 and T-1 came from different snapshots.  ASW/Z/OAS carry a small
        # standing difference because they are quoted on their own conventions.
        legs["duration_identity_check"] = np.where(
            _both(duration_total, legs["pnl_yield_only"]),
            duration_total - legs["pnl_yield_only"],
            np.nan,
        )

        # ---- convexity ------------------------------------------------------ #
        # Second order in the SAME yield move, struck against the opening market
        # value so it is consistent with the DV01 above.
        legs["pnl_convexity"] = np.where(
            np.isfinite(market["mv_prior_base"])
            & np.isfinite(market["convexity"])
            & np.isfinite(deltas["delta_y_bp"]),
            0.5
            * market["mv_prior_base"]
            * market["convexity"]
            * (deltas["delta_y_bp"] / 10_000.0) ** 2,
            np.nan,
        )

        # ---- carry ---------------------------------------------------------- #
        notional = market["notional"]
        fx_prior = market["fx_prior"]

        if cfg.model.coupon_carry == "exact":
            # accrued(T0) - accrued(T-1) + coupons actually paid.  Survives an
            # ex-coupon date: the accrued term goes sharply negative and the
            # paid term puts the cash back, leaving one period's accrual.
            per_100 = (
                market["accrued_current"] - market["accrued_prior"] + market["coupons_paid"]
            )
            carry_coupon = notional / 100.0 * per_100 * fx_prior
        else:
            carry_coupon = notional * bonds.coupon_decimal * self.year_fraction * fx_prior

        coupon_cash_base = notional / 100.0 * market["coupons_paid"] * fx_prior

        if cfg.model.pull_to_par.enabled:
            base_curve = curve_for_framework(frameworks.code)
            prior_spread = self._prior_spread_bp(bonds, deltas, frameworks)
            ptp = bondmath.pull_to_par(
                np.full(n, self.prior, dtype=_DAY),
                np.full(n, self.current, dtype=_DAY),
                bonds.maturity,
                bonds.coupon_decimal,
                np.asarray(bonds.coupon_freq, dtype=np.float64),
                np.asarray(bonds.day_count_code, dtype=np.float64),
                bonds.currency,
                base_curve,
                prior_spread,
                self.inputs.curves,
                day_basis=cfg.model.pull_to_par.day_basis,
                max_flows=cfg.model.pull_to_par.max_cashflows,
            )
            carry_roll = np.where(
                np.isfinite(ptp.clean_change) & np.isfinite(notional) & np.isfinite(fx_prior),
                notional * fx_prior * ptp.clean_change / 100.0,
                np.nan,
            )
            legs["pull_to_par_price"] = ptp.clean_change
            legs["pull_to_par_usable"] = ptp.usable
        else:
            carry_roll = np.zeros(n)
            legs["pull_to_par_price"] = np.zeros(n)
            legs["pull_to_par_usable"] = np.ones(n, dtype=bool)

        legs["carry_coupon"] = carry_coupon
        legs["carry_roll_to_par"] = carry_roll
        legs["coupon_cash"] = coupon_cash_base
        legs["carry_total"] = np.where(
            _both(carry_coupon, carry_roll), carry_coupon + carry_roll, np.nan
        )

        # Funding is ECONOMIC carry, not mark-to-market PnL.  Practical PnL here
        # is a mark-to-market figure with no financing leg, so putting funding
        # in the bridge would open a residual of exactly the funding cost.  It
        # is reported as a memo unless the config says the feed is funded.
        funding = _coalesce(
            _mean_rate(bonds.funding_rate_prior, bonds.funding_rate_current),
        )
        funding_decimal = np.where(np.abs(funding) > 1.0, funding / 100.0, funding)
        legs["funding_carry_memo"] = np.where(
            np.isfinite(market["mv_prior_base"]) & np.isfinite(funding_decimal),
            -market["mv_prior_base"] * funding_decimal * self.year_fraction,
            np.nan,
        )

        # ---- FX ------------------------------------------------------------- #
        # d(MV_base) = MV_local_prior * dFX + d(MV_local) * FX_prior
        #            + d(MV_local) * dFX          <- exact, nothing left over
        d_fx = np.where(
            _both(market["fx_current"], market["fx_prior"]),
            market["fx_current"] - market["fx_prior"],
            np.nan,
        )
        d_mv_local = np.where(
            _both(market["mv_current_local"], market["mv_prior_local"]),
            market["mv_current_local"] - market["mv_prior_local"],
            np.nan,
        )
        is_base = np.array([c == cfg.run.base_ccy for c in bonds.currency.tolist()])

        pnl_fx = np.where(is_base, 0.0, market["mv_prior_local"] * d_fx)
        pnl_fx_cross = (
            np.where(is_base, 0.0, d_mv_local * d_fx)
            if cfg.model.fx_cross_term
            else np.zeros(n)
        )
        legs["pnl_fx"] = pnl_fx
        legs["pnl_fx_cross"] = pnl_fx_cross

        # ---- hedge legs ------------------------------------------------------ #
        fut_dv01_attr = (
            linkage.futures_dv01_prior
            if cfg.model.attribution_risk_date == "prior"
            else linkage.futures_dv01
        )
        # Futures hedge the deliverable GOVERNMENT curve through the CTD, so the
        # move they respond to is the government curve, not the bond's yield.
        model_futures = np.where(
            linkage.futures_count == 0,
            0.0,
            _first_order(fut_dv01_attr, deltas["delta_gov_bp"]),
        )

        # Swaps hedge the curve their floating leg projects off.  Splitting the
        # DV01 by family and applying each against its own curve move is what
        # stops an ESTR swap's PnL being measured against EURIBOR - which would
        # book the OIS/IBOR basis as model error every day.
        ois_leg = _first_order(linkage.plain_swap_dv01_ois, deltas["delta_r_bp"])
        ibor_leg = _first_order(linkage.plain_swap_dv01_ibor, deltas["delta_swap_bp"])
        model_swap = np.where(
            linkage.plain_swap_count == 0,
            0.0,
            np.where(
                linkage.plain_swap_dv01_unknown != 0.0,
                np.nan,  # an unclassified floating leg has no curve to measure against
                np.nan_to_num(ois_leg, nan=0.0) + np.nan_to_num(ibor_leg, nan=0.0),
            ),
        )
        model_swap = np.where(
            (linkage.plain_swap_count > 0)
            & (linkage.plain_swap_dv01_unknown == 0.0)
            & ~(np.isfinite(ois_leg) | np.isfinite(ibor_leg)),
            np.nan,
            model_swap,
        )

        actual_futures = np.where(linkage.futures_count == 0, 0.0, linkage.futures_pnl)
        actual_swap = np.where(linkage.plain_swap_count == 0, 0.0, linkage.plain_swap_pnl)

        legs["model_futures_pnl"] = model_futures
        legs["model_swap_pnl"] = model_swap
        legs["actual_futures_pnl"] = actual_futures
        legs["actual_plain_swap_pnl"] = actual_swap

        legs["futures_basis_pnl"] = np.where(
            _both(actual_futures, model_futures), actual_futures - model_futures, np.nan
        )
        legs["swap_basis_pnl"] = np.where(
            _both(actual_swap, model_swap), actual_swap - model_swap, np.nan
        )
        legs["model_hedge_pnl"] = np.where(
            _both(model_futures, model_swap), model_futures + model_swap, np.nan
        )
        legs["actual_hedge_pnl"] = np.where(
            _both(actual_futures, actual_swap), actual_futures + actual_swap, np.nan
        )
        legs["hedge_basis_pnl"] = np.where(
            _both(legs["actual_hedge_pnl"], legs["model_hedge_pnl"]),
            legs["actual_hedge_pnl"] - legs["model_hedge_pnl"],
            np.nan,
        )

        # ---- the two totals and the residual --------------------------------- #
        theoretical_parts = [
            legs["pnl_duration_total"],
            legs["pnl_convexity"],
            legs["carry_total"],
            legs["pnl_fx"],
            legs["pnl_fx_cross"],
            legs["model_hedge_pnl"],
        ]
        if cfg.model.include_funding_in_bridge:
            theoretical_parts.append(legs["funding_carry_memo"])

        theoretical = _all_or_nothing(theoretical_parts)
        legs["theoretical_pnl"] = theoretical

        legs["total_explained"] = np.where(
            _both(theoretical, legs["hedge_basis_pnl"]),
            theoretical + legs["hedge_basis_pnl"],
            np.nan,
        )

        practical = _all_or_nothing(
            [market["delta_mv_base"], legs["actual_hedge_pnl"], coupon_cash_base]
        )
        legs["practical_pnl"] = practical

        legs["residual_pnl"] = np.where(
            _both(practical, legs["total_explained"]),
            practical - legs["total_explained"],
            np.nan,
        )
        legs["model_residual_pnl"] = np.where(
            _both(practical, theoretical), practical - theoretical, np.nan
        )
        with np.errstate(divide="ignore", invalid="ignore"):
            legs["residual_pct"] = np.where(
                np.isfinite(legs["residual_pnl"]) & np.isfinite(practical) & (practical != 0),
                legs["residual_pnl"] / np.abs(practical),
                np.nan,
            )

        # What the legacy timing convention would have added: the same duration
        # leg struck on CURRENT risk, less the one struck on opening risk.
        legs["dv01_timing_bias"] = np.where(
            _both(market["dv01_current"], market["dv01_prior"])
            & np.isfinite(deltas["delta_y_bp"]),
            -(market["dv01_current"] - market["dv01_prior"]) * deltas["delta_y_bp"],
            np.nan,
        )

        return legs

    def _prior_spread_bp(
        self,
        bonds: BondBook,
        deltas: dict[str, np.ndarray],
        frameworks: FrameworkResolution,
    ) -> np.ndarray:
        """The T-1 spread the pull-to-par calculation is calibrated on.

        Chosen by the bond's own framework, so the price is rebuilt on the same
        base curve plus the same spread the attribution measures it against.
        Mixing them - repricing on the swap curve with a G-spread, say - puts
        the swap/government basis straight into carry.
        """
        n = len(bonds)
        out = np.full(n, np.nan)
        source = {
            "G": deltas["gspread_prior"],
            "I": deltas["ispread_prior"],
            "ASW": np.asarray(bonds.asw_prior, dtype=np.float64),
            "Z": np.asarray(bonds.zspread_prior, dtype=np.float64),
            "OAS": np.asarray(bonds.oas_prior, dtype=np.float64),
            "OIS": deltas["gspread_prior"],
            "SOFR": deltas["gspread_prior"],
            "MIXED": deltas["ispread_prior"],
        }
        for code, values in source.items():
            mask = frameworks.code == code
            if mask.any():
                out = np.where(mask, values, out)
        return out

    # ---------------------------------------------------------------------- #
    # frames
    # ---------------------------------------------------------------------- #

    def _bond_frame(
        self,
        bonds: BondBook,
        market: dict[str, np.ndarray],
        deltas: dict[str, np.ndarray],
        linkage: HedgeLinkage,
        frameworks: FrameworkResolution,
        legs: dict[str, np.ndarray],
    ) -> ColumnFrame:
        n = len(bonds)
        tol = self.config.tolerances

        actual_hedge_dv01 = np.where(
            _both(linkage.futures_dv01, linkage.plain_swap_dv01),
            linkage.futures_dv01 + linkage.plain_swap_dv01,
            np.nan,
        )
        # The hedge a bond SHOULD have: replicate the synthetic where the
        # coverage relationship names one, otherwise be flat.  The rule lives
        # here and only here - the legacy sheet applied one rule and the
        # dashboard another, so the same bond scored differently depending on
        # which you read.
        target_hedge_dv01 = np.where(
            np.isfinite(linkage.synthetic_swap_dv01) & (linkage.synthetic_swap_dv01 != 0.0),
            linkage.synthetic_swap_dv01,
            np.where(
                np.isfinite(market["dv01_current"]) & (market["dv01_current"] != 0.0),
                -market["dv01_current"],
                np.nan,
            ),
        )
        hedge_gap = np.where(
            _both(actual_hedge_dv01, target_hedge_dv01),
            actual_hedge_dv01 - target_hedge_dv01,
            np.nan,
        )
        with np.errstate(divide="ignore", invalid="ignore"):
            hedge_ratio = np.where(
                np.isfinite(actual_hedge_dv01)
                & np.isfinite(market["dv01_current"])
                & (market["dv01_current"] != 0),
                -actual_hedge_dv01 / market["dv01_current"],
                np.nan,
            )
            # An unhedged bond has a hedge ratio of zero, not of negative zero.
            hedge_ratio = hedge_ratio + 0.0
            # NOT clamped to [0, 1] on purpose.  Clamping reported a wrong-way
            # hedge and a triple-sized hedge both as 0%, hiding exactly the
            # positions worth looking at.  -1 means the hedge is as wrong as it
            # could be at that size, and that is worth seeing.
            hedge_efficiency = np.where(
                np.isfinite(hedge_gap) & np.isfinite(target_hedge_dv01) & (target_hedge_dv01 != 0),
                1.0 - np.abs(hedge_gap) / np.abs(target_hedge_dv01),
                np.nan,
            )
        residual_dv01 = np.where(
            _both(market["dv01_current"], actual_hedge_dv01),
            market["dv01_current"] + actual_hedge_dv01,
            np.nan,
        )

        status = self._attribution_status(
            bonds, market, deltas, linkage, frameworks, legs,
            hedge_ratio, actual_hedge_dv01, tol,
        )

        data: dict[str, np.ndarray] = {
            "isin": bonds.isin,
            "name": bonds.name,
            "currency": bonds.currency,
            "portfolio": bonds.portfolio,
            "acctg_cat": bonds.acctg_cat,
            "maturity": bonds.maturity,
            "notional": market["notional"],
            "days": np.full(n, float(self.days)),
            "year_fraction": np.full(n, self.year_fraction),
            "years_to_maturity": market["years_prior"],

            "fx_prior": market["fx_prior"],
            "fx_current": market["fx_current"],
            "clean_px_prior": bonds.clean_px_prior,
            "clean_px_current": bonds.clean_px_current,
            "dirty_px_prior": bonds.dirty_px_prior,
            "dirty_px_current": bonds.dirty_px_current,
            "accrued_prior": market["accrued_prior"],
            "accrued_current": market["accrued_current"],
            "coupons_paid_per_100": market["coupons_paid"],
            "mv_prior_base": market["mv_prior_base"],
            "mv_current_base": market["mv_current_base"],
            "delta_mv_base": market["delta_mv_base"],

            "duration_prior": market["duration_prior"],
            "duration_current": market["duration_current"],
            "duration_source": market["duration_source"],
            "convexity": market["convexity"],
            "dv01_unit_prior": market["dv01_unit_prior"],
            "dv01_prior": market["dv01_prior"],
            "dv01_current": market["dv01_current"],
            "attribution_dv01": market["attribution_dv01"],
            "dv01_timing_bias": legs["dv01_timing_bias"],

            "ytm_prior": bonds.ytm_prior,
            "ytm_current": bonds.ytm_current,
            "ois_prior": deltas["ois_prior"],
            "ois_current": deltas["ois_current"],
            "gov_prior": deltas["gov_prior"],
            "gov_current": deltas["gov_current"],
            "swap_prior": deltas["swap_prior"],
            "swap_current": deltas["swap_current"],
            "ispread_prior": deltas["ispread_prior"],
            "ispread_current": deltas["ispread_current"],
            "gspread_prior": deltas["gspread_prior"],
            "gspread_current": deltas["gspread_current"],

            "delta_y_bp": deltas["delta_y_bp"],
            "delta_r_bp": deltas["delta_r_bp"],
            "delta_gov_bp": deltas["delta_gov_bp"],
            "delta_swap_bp": deltas["delta_swap_bp"],
            "delta_g_bp": deltas["delta_g_bp"],
            "delta_q_bp": deltas["delta_q_bp"],
            "delta_ispread_bp": deltas["delta_ispread_bp"],
            "delta_gspread_bp": deltas["delta_gspread_bp"],
            "delta_zspread_bp": deltas["delta_zspread_bp"],
            "delta_asw_bp": deltas["delta_asw_bp"],
            "delta_oas_bp": deltas["delta_oas_bp"],

            "spread_framework": frameworks.code,
            "spread_framework_reason": frameworks.reason,
            "framework_futures_weight": frameworks.futures_weight,
            "framework_swap_weight": frameworks.swap_weight,

            "pnl_ois": legs["pnl_ois"],
            "pnl_gov_basis": legs["pnl_gov_basis"],
            "pnl_swap_gov_basis": legs["pnl_swap_gov_basis"],
            "pnl_ispread": legs["pnl_ispread"],
            "pnl_gspread": legs["pnl_gspread"],
            "pnl_zspread": legs["pnl_zspread"],
            "pnl_asw": legs["pnl_asw"],
            "pnl_oas": legs["pnl_oas"],
            "pnl_yield_only": legs["pnl_yield_only"],
            "spread_pnl_used": legs["spread_pnl_used"],
            "pnl_duration_total": legs["pnl_duration_total"],
            "duration_identity_check": legs["duration_identity_check"],
            "pnl_convexity": legs["pnl_convexity"],

            "carry_coupon": legs["carry_coupon"],
            "carry_roll_to_par": legs["carry_roll_to_par"],
            "pull_to_par_price": legs["pull_to_par_price"],
            "carry_total": legs["carry_total"],
            "coupon_cash": legs["coupon_cash"],
            "funding_carry_memo": legs["funding_carry_memo"],

            "pnl_fx": legs["pnl_fx"],
            "pnl_fx_cross": legs["pnl_fx_cross"],

            "futures_count": linkage.futures_count,
            "plain_swap_count": linkage.plain_swap_count,
            "synthetic_swap_count": linkage.synthetic_swap_count,
            "futures_dv01": linkage.futures_dv01,
            "plain_swap_dv01": linkage.plain_swap_dv01,
            "plain_swap_dv01_ois": linkage.plain_swap_dv01_ois,
            "plain_swap_dv01_ibor": linkage.plain_swap_dv01_ibor,
            "plain_swap_dv01_unknown": linkage.plain_swap_dv01_unknown,
            "synthetic_swap_dv01": linkage.synthetic_swap_dv01,
            "actual_hedge_dv01": actual_hedge_dv01,
            "target_hedge_dv01": target_hedge_dv01,
            "hedge_dv01_gap": hedge_gap,
            "residual_dv01": residual_dv01,
            "hedge_ratio": hedge_ratio,
            "hedge_efficiency": hedge_efficiency,

            "model_futures_pnl": legs["model_futures_pnl"],
            "actual_futures_pnl": legs["actual_futures_pnl"],
            "futures_basis_pnl": legs["futures_basis_pnl"],
            "model_swap_pnl": legs["model_swap_pnl"],
            "actual_plain_swap_pnl": legs["actual_plain_swap_pnl"],
            "swap_basis_pnl": legs["swap_basis_pnl"],
            "model_hedge_pnl": legs["model_hedge_pnl"],
            "actual_hedge_pnl": legs["actual_hedge_pnl"],
            "hedge_basis_pnl": legs["hedge_basis_pnl"],

            "theoretical_pnl": legs["theoretical_pnl"],
            "total_explained": legs["total_explained"],
            "practical_pnl": legs["practical_pnl"],
            "residual_pnl": legs["residual_pnl"],
            "model_residual_pnl": legs["model_residual_pnl"],
            "residual_pct": legs["residual_pct"],

            "shared_isin": bonds.duplicate_isin_mask(),
            "attribution_status": status,
        }
        return ColumnFrame.build(BOND_SCHEMA, data)

    def _attribution_status(
        self,
        bonds: BondBook,
        market: dict[str, np.ndarray],
        deltas: dict[str, np.ndarray],
        linkage: HedgeLinkage,
        frameworks: FrameworkResolution,
        legs: dict[str, np.ndarray],
        hedge_ratio: np.ndarray,
        actual_hedge_dv01: np.ndarray,
        tol,
    ) -> np.ndarray:
        """First failing check wins, ordered cause before symptom.

        Bad market data is reported ahead of a missing framework total, and a
        missing framework total ahead of a high residual - because each of those
        causes the next, and reporting the symptom sends the reader to the wrong
        place.
        """
        n = len(bonds)
        status = np.full(n, "OK", dtype=object)

        def mark(condition: np.ndarray, text: str) -> None:
            nonlocal status
            status = np.where((status == "OK") & condition, text, status)

        missing_marks = ~(
            np.isfinite(bonds.dirty_px_prior)
            & np.isfinite(bonds.dirty_px_current)
            & np.isfinite(bonds.ytm_prior)
            & np.isfinite(bonds.ytm_current)
        )
        mark(missing_marks, "Missing market data")
        mark(~np.isfinite(np.asarray(bonds.day_count_code, dtype=np.float64)),
             "Unknown day-count convention")
        mark(~np.isin(np.asarray(bonds.coupon_freq, dtype=np.float64),
                      np.array(dc.VALID_FREQUENCIES, dtype=np.float64)),
             "Unknown coupon frequency")
        mark(~np.isfinite(deltas["delta_r_bp"]), "Missing OIS curve")
        mark(~np.isfinite(deltas["delta_gov_bp"]), "Missing government curve")
        mark(frameworks.code == "REVIEW", "Framework review required")
        mark(linkage.plain_swap_dv01_unknown != 0.0, "Unclassified swap floating leg")
        mark(
            (linkage.plain_swap_count > 0) & ~np.isfinite(legs["actual_plain_swap_pnl"]),
            "Missing actual swap PnL",
        )
        mark(
            (linkage.futures_count > 0) & ~np.isfinite(legs["actual_futures_pnl"]),
            "Missing actual futures PnL",
        )
        mark(np.isfinite(hedge_ratio) & (hedge_ratio < 0), "Wrong-direction hedge")
        mark(
            np.isfinite(hedge_ratio) & (hedge_ratio > 1.0 + tol.hedge_ratio),
            "Over-hedged",
        )
        mark(~np.isfinite(legs["pnl_duration_total"]), "Missing selected framework data")
        mark(
            np.isfinite(legs["duration_identity_check"])
            & (
                np.abs(legs["duration_identity_check"])
                > np.maximum(
                    tol.identity_eur,
                    tol.identity_pct * np.abs(np.nan_to_num(legs["pnl_duration_total"])),
                )
            ),
            "Duration chain does not tie",
        )
        mark(~np.isfinite(legs["practical_pnl"]), "Missing practical PnL")
        mark(bonds.duplicate_isin_mask(), "Hedge shared with another row (same ISIN)")
        # Both tests, not either: a residual has to be a large SHARE and a real
        # AMOUNT.  On its own the percentage flags a 24 EUR break on a 65 EUR
        # position and buries the ones that matter underneath it.
        mark(
            np.isfinite(legs["residual_pct"])
            & (np.abs(legs["residual_pct"]) > tol.residual_pct)
            & (np.abs(np.nan_to_num(legs["residual_pnl"])) > tol.residual_eur),
            "High residual",
        )
        return status

    # -- futures frame -------------------------------------------------------- #

    def _futures_frame(self, bonds: BondBook, market: dict[str, np.ndarray]) -> ColumnFrame:
        futures = self.inputs.futures
        n = len(futures)
        if n == 0:
            return ColumnFrame.build(
                FUTURES_SCHEMA, {c.name: np.array([]) for c in FUTURES_SCHEMA}
            )

        dv01_current, dv01_prior = self._futures_dv01(futures)
        actual = self._futures_actual_pnl(futures)

        # Tenor for the curve move a future responds to: its CTD's, if the CTD
        # is in the bond book; otherwise the bond it hedges.  Neither available
        # means no model PnL - a tenor guessed from the delivery month would
        # produce a plausible number from an assumption nobody made.
        tenor_by_isin = {
            str(k): float(v)
            for k, v in zip(bonds.isin.tolist(), market["years_prior"].tolist())
        }
        tenor = np.array(
            [
                tenor_by_isin.get(str(c), tenor_by_isin.get(str(l), np.nan))
                for c, l in zip(futures.ctd_isin.tolist(), futures.linked_isin.tolist())
            ],
            dtype=np.float64,
        )
        tenor_source = np.array(
            [
                "CTD" if str(c) in tenor_by_isin
                else ("linked bond" if str(l) in tenor_by_isin else "unresolved")
                for c, l in zip(futures.ctd_isin.tolist(), futures.linked_isin.tolist())
            ],
            dtype=object,
        )

        curves = self.inputs.curves
        gov_prior = curves.rates_for(futures.currency, tenor, "GOV", "prior")
        gov_current = curves.rates_for(futures.currency, tenor, "GOV", "current")
        delta_gov = _diff_bp(gov_current, gov_prior, 100.0)

        dv01_attr = (
            dv01_prior if self.config.model.attribution_risk_date == "prior" else dv01_current
        )
        model = _first_order(dv01_attr, delta_gov)
        basis = np.where(_both(actual, model), actual - model, np.nan)

        notional_value = (
            np.asarray(futures.contracts, dtype=np.float64)
            * np.asarray(futures.fut_val_pt, dtype=np.float64)
            * np.asarray(futures.fut_px_current, dtype=np.float64)
            * np.asarray(futures.fx_current, dtype=np.float64)
        )

        # Gross basis is a CLEAN-price concept: CTD clean price less the futures
        # price times the conversion factor.  Computed from the dirty price it
        # is overstated by the CTD's accrued interest, which is not basis.
        # The CTD's accrued is available whenever the CTD is in the bond book.
        accrued_by_isin = {
            str(k): float(v)
            for k, v in zip(bonds.isin.tolist(), market["accrued_current"].tolist())
        }
        ctd_accrued = np.array(
            [accrued_by_isin.get(str(c), np.nan) for c in futures.ctd_isin.tolist()],
            dtype=np.float64,
        )
        ctd_clean = np.asarray(futures.ctd_dirty_px_current, dtype=np.float64) - ctd_accrued
        gross_basis = np.where(
            np.isfinite(ctd_clean)
            & np.isfinite(futures.fut_px_current)
            & np.isfinite(futures.ctd_cf),
            ctd_clean - futures.fut_px_current * futures.ctd_cf,
            np.nan,
        )
        basis_quality = np.where(
            np.isfinite(ctd_accrued), "clean (CTD accrued known)", "unavailable (CTD not in book)"
        ).astype(object)

        status = np.full(n, "OK", dtype=object)

        def mark(condition, text):
            nonlocal status
            status = np.where((status == "OK") & condition, text, status)

        mark(~futures.is_linked, "Unlinked - not on any bond row")
        mark(~np.isfinite(np.asarray(futures.fut_px_prior, dtype=np.float64)),
             "Missing prior futures price")
        mark(~np.isfinite(np.asarray(futures.fut_px_current, dtype=np.float64)),
             "Missing current futures price")
        mark(~np.isfinite(np.asarray(futures.fut_val_pt, dtype=np.float64)),
             "Missing point value")
        mark(~np.isfinite(np.asarray(futures.fut_px_val_bp, dtype=np.float64)),
             "Missing unit DV01")
        mark(np.array([not bool(s) for s in futures.ctd_isin.tolist()]), "Missing CTD")
        mark(tenor_source == "unresolved", "No tenor for the curve move")

        data = {
            "contract_code": futures.contract_code,
            "exchange": futures.exchange,
            "currency": futures.currency,
            "portfolio": futures.portfolio,
            "linked_isin": futures.linked_isin,
            "hedge_type": futures.hedge_type,
            "ctd_isin": futures.ctd_isin,
            "ctd_cf": futures.ctd_cf,
            "deliv_date": futures.deliv_date,
            "contracts": futures.contracts,
            "fut_px_prior": futures.fut_px_prior,
            "fut_px_current": futures.fut_px_current,
            "fut_val_pt": futures.fut_val_pt,
            "fut_px_val_bp": futures.fut_px_val_bp,
            "fx_prior": futures.fx_prior,
            "fx_current": futures.fx_current,
            "notional_value_base": notional_value,
            "tenor_years": tenor,
            "tenor_source": tenor_source,
            "gov_prior": gov_prior,
            "gov_current": gov_current,
            "delta_gov_bp": delta_gov,
            "dv01_prior": dv01_prior,
            "dv01_current": dv01_current,
            "gross_basis": gross_basis,
            "gross_basis_quality": basis_quality,
            "gross_basis_bbg": futures.gross_basis_bbg,
            "net_basis_bbg": futures.net_basis_bbg,
            "implied_repo_bbg": futures.implied_repo_bbg,
            "theoretical_pnl": model,
            "practical_pnl": actual,
            "basis_pnl": basis,
            "is_linked": futures.is_linked,
            "status": status,
        }
        return ColumnFrame.build(FUTURES_SCHEMA, data)

    # -- swaps frame ---------------------------------------------------------- #

    def _swaps_frame(self) -> ColumnFrame:
        swaps = self.inputs.swaps
        n = len(swaps)
        if n == 0:
            return ColumnFrame.build(
                SWAPS_SCHEMA, {c.name: np.array([]) for c in SWAPS_SCHEMA}
            )

        curves = self.inputs.curves
        years = swaps.year_fraction_to_end(self.current)

        ois_prior = curves.rates_for(swaps.currency, years, "OIS", "prior")
        ois_current = curves.rates_for(swaps.currency, years, "OIS", "current")
        swap_prior = curves.rates_for(swaps.currency, years, "SWAP", "prior")
        swap_current = curves.rates_for(swaps.currency, years, "SWAP", "current")

        family = np.array([str(f) for f in swaps.float_curve_type.tolist()], dtype=object)
        is_ois = np.isin(family, np.array(_OIS_FAMILIES, dtype=object))
        is_ibor = np.isin(family, np.array(_IBOR_FAMILIES, dtype=object))

        float_prior = np.where(is_ois, ois_prior, np.where(is_ibor, swap_prior, np.nan))
        float_current = np.where(is_ois, ois_current, np.where(is_ibor, swap_current, np.nan))
        delta_float_bp = _diff_bp(float_current, float_prior, 100.0)

        # Model spread: the fixed rate plus the contractual float spread, less
        # the curve the floating leg actually projects off.  Its move is the
        # part of the swap's value change the curve does not explain.
        fixed = np.asarray(swaps.fixed_rate, dtype=np.float64)
        fspread = np.nan_to_num(np.asarray(swaps.float_spread, dtype=np.float64), nan=0.0)
        model_spread_prior = np.where(
            np.isfinite(fixed) & np.isfinite(float_prior), fixed + fspread - float_prior, np.nan
        )
        model_spread_current = np.where(
            np.isfinite(fixed) & np.isfinite(float_current), fixed + fspread - float_current, np.nan
        )
        delta_model_spread_bp = _diff_bp(model_spread_current, model_spread_prior, 100.0)

        dv01 = self._swap_dv01(swaps)
        dv01_model = self._swap_model_dv01(swaps)
        dv01_source = np.where(
            np.isfinite(np.asarray(swaps.dv01_supplied, dtype=np.float64))
            & (self.config.swaps.dv01_source == "supplied"),
            "supplied",
            "annuity model",
        ).astype(object)

        actual = self._swap_actual_pnl(swaps)
        model = _first_order(dv01, delta_float_bp)
        basis = np.where(_both(actual, model), actual - model, np.nan)

        # A synthetic is a TARGET, not a position.  Its PnL must never reach a
        # total, so it is reported as NaN here and excluded everywhere.
        synthetic = swaps.is_synthetic
        actual = np.where(synthetic, np.nan, actual)
        model = np.where(synthetic, np.nan, model)
        basis = np.where(synthetic, np.nan, basis)

        family_status = np.array(
            [
                float_family_status(idx, ccy, fam)
                for idx, ccy, fam in zip(
                    swaps.float_index.tolist(), swaps.currency.tolist(), family.tolist()
                )
            ],
            dtype=object,
        )

        status = np.full(n, "OK", dtype=object)

        def mark(condition, text):
            nonlocal status
            status = np.where((status == "OK") & condition, text, status)

        mark(synthetic, "Synthetic target - excluded from all totals")
        mark(family_status != "OK", family_status)
        mark(~swaps.is_linked, "Unlinked - not on any bond row")
        mark(~np.isfinite(dv01), "Missing DV01")
        mark(~np.isfinite(actual) & ~synthetic, "Missing NPV move")
        mark(~np.isfinite(delta_float_bp), "Missing floating curve")

        data = {
            "deal_id": swaps.deal_id,
            "currency": swaps.currency,
            "portfolio": swaps.portfolio,
            "counterparty": swaps.counterparty,
            "linked_isin": swaps.linked_isin,
            "swap_id_source": swaps.swap_id_source,
            "notional": swaps.notional,
            "pay_fixed": swaps.pay_fixed,
            "fixed_rate": swaps.fixed_rate,
            "float_index": swaps.float_index,
            "float_spread": swaps.float_spread,
            "float_curve_type": swaps.float_curve_type,
            "float_family_status": family_status,
            "start_date": swaps.start_date,
            "end_date": swaps.end_date,
            "year_fraction": years,
            "ois_prior": ois_prior,
            "ois_current": ois_current,
            "float_curve_prior": float_prior,
            "float_curve_current": float_current,
            "delta_float_curve_bp": delta_float_bp,
            "model_spread_prior": model_spread_prior,
            "model_spread_current": model_spread_current,
            "delta_model_spread_bp": delta_model_spread_bp,
            "dv01": dv01,
            "dv01_model": dv01_model,
            "dv01_source": dv01_source,
            "npv_prior": swaps.npv_prior,
            "npv_current": swaps.npv_current,
            "theoretical_pnl": model,
            "practical_pnl": actual,
            "basis_pnl": basis,
            "is_linked": swaps.is_linked,
            "is_plain": swaps.is_plain,
            "status": status,
        }
        return ColumnFrame.build(SWAPS_SCHEMA, data)

    # -- diagnostics ---------------------------------------------------------- #

    def _diagnostics(
        self, bonds: BondBook, frame: ColumnFrame, linkage: HedgeLinkage
    ) -> dict[str, object]:
        n = len(bonds)
        status = frame["attribution_status"]
        values, counts = np.unique(status, return_counts=True)
        return {
            "bond_count": n,
            "status_counts": {str(v): int(c) for v, c in zip(values.tolist(), counts.tolist())},
            "rows_ok": int((status == "OK").sum()),
            "rows_attributed": frame.count_finite("total_explained"),
            "rows_with_practical": frame.count_finite("practical_pnl"),
            "shared_isin_rows": int(np.asarray(frame["shared_isin"], dtype=bool).sum()),
            "unlinked_futures_pnl": linkage.unlinked_futures_pnl,
            "unlinked_futures_count": linkage.unlinked_futures_count,
            "unlinked_swap_pnl": linkage.unlinked_swap_pnl,
            "unlinked_swap_count": linkage.unlinked_swap_count,
            "orphan_futures_pnl": linkage.orphan_futures_pnl,
            "orphan_futures_count": linkage.orphan_futures_count,
            "orphan_swap_pnl": linkage.orphan_swap_pnl,
            "orphan_swap_count": linkage.orphan_swap_count,
        }


# --------------------------------------------------------------------------- #
# small numeric helpers
# --------------------------------------------------------------------------- #


def _coalesce(*arrays: np.ndarray) -> np.ndarray:
    """First finite value across the arrays, elementwise."""
    out = np.asarray(arrays[0], dtype=np.float64).copy()
    for other in arrays[1:]:
        other = np.asarray(other, dtype=np.float64)
        need = ~np.isfinite(out)
        if not need.any():
            break
        out = np.where(need, other, out)
    return out


def _mean_rate(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    """Average of two rates, tolerating one of them being absent."""
    a = np.asarray(a, dtype=np.float64)
    b = np.asarray(b, dtype=np.float64)
    both = _both(a, b)
    return np.where(both, (a + b) / 2.0, np.where(np.isfinite(a), a, b))


def _chain_sum(legs: dict[str, np.ndarray], names: tuple[str, ...], n: int) -> np.ndarray:
    """Sum a framework chain, blank unless EVERY leg is a number.

    All-or-nothing on purpose.  A chain missing one leg is not a smaller
    attribution, it is an incomplete one, and a partial total that looks
    complete is worse than a blank: it ties to nothing and nobody can see why.
    """
    if not names:
        return np.full(n, np.nan)
    stack = np.vstack([np.asarray(legs[name], dtype=np.float64) for name in names])
    complete = np.isfinite(stack).all(axis=0)
    return np.where(complete, stack.sum(axis=0), np.nan)


def _all_or_nothing(parts: list[np.ndarray]) -> np.ndarray:
    stack = np.vstack([np.asarray(p, dtype=np.float64) for p in parts])
    complete = np.isfinite(stack).all(axis=0)
    return np.where(complete, stack.sum(axis=0), np.nan)


# --------------------------------------------------------------------------- #
# schemas - the single source of truth for names, units and meanings
# --------------------------------------------------------------------------- #

_C = Column

BOND_SCHEMA: tuple[Column, ...] = (
    # identity
    _C("isin", "ISIN", Kind.TEXT, "Instrument identifier; also the hedge join key.", group="Identity"),
    _C("name", "Name", Kind.TEXT, "Instrument description.", group="Identity"),
    _C("currency", "CCY", Kind.TEXT, "Position currency.", group="Identity"),
    _C("portfolio", "Portfolio", Kind.TEXT, "Owning portfolio.", group="Identity"),
    _C("acctg_cat", "Acctg Cat", Kind.TEXT, "Accounting classification (HTC, AFS, trading).", group="Identity"),
    _C("maturity", "Maturity", Kind.DATE, "Redemption date.", group="Identity"),
    _C("notional", "Notional", Kind.NOTIONAL, "Signed nominal. Negative is a short; the only sign driver in the model.", group="Identity"),
    _C("days", "Days", Kind.COUNT, "Calendar days in the period being explained.", "as_of - prior", group="Identity"),
    _C("year_fraction", "Year Frac", Kind.RATIO, "Period length on the ACT/365F carry axis.", "days / 365", group="Identity"),
    _C("years_to_maturity", "Years to Mat", Kind.RATIO, "Tenor at which every curve is read for this bond.", group="Identity"),

    # marks
    _C("fx_prior", "FX T-1", Kind.RATIO, "Base currency per unit of position currency at the open.", group="Marks"),
    _C("fx_current", "FX T0", Kind.RATIO, "Base currency per unit of position currency at the close.", group="Marks"),
    _C("clean_px_prior", "Clean Px T-1", Kind.PRICE, "Quoted clean price at the open.", group="Marks"),
    _C("clean_px_current", "Clean Px T0", Kind.PRICE, "Quoted clean price at the close.", group="Marks"),
    _C("dirty_px_prior", "Dirty Px T-1", Kind.PRICE, "Clean price plus accrued at the open; what the position is marked at.", group="Marks"),
    _C("dirty_px_current", "Dirty Px T0", Kind.PRICE, "Clean price plus accrued at the close.", group="Marks"),
    _C("accrued_prior", "Accrued T-1", Kind.PRICE, "Accrued interest per 100 at the open, on the bond's own day count.", group="Marks"),
    _C("accrued_current", "Accrued T0", Kind.PRICE, "Accrued interest per 100 at the close.", group="Marks"),
    _C("coupons_paid_per_100", "Coupons Paid /100", Kind.PRICE, "Coupon cash per 100 nominal actually paid inside the period.", "sum of coupons in (T-1, T0]", group="Marks"),
    _C("mv_prior_base", "MV T-1", Kind.EUR, "Dirty market value at the open, in base currency.", "notional * dirty_px_prior/100 * fx_prior", group="Marks"),
    _C("mv_current_base", "MV T0", Kind.EUR, "Dirty market value at the close, in base currency.", "notional * dirty_px_current/100 * fx_current", group="Marks"),
    _C("delta_mv_base", "Delta MV", Kind.EUR, "Change in dirty market value; the mark-to-market part of practical PnL.", "MV T0 - MV T-1", group="Marks"),

    # risk
    _C("duration_prior", "ModDur T-1", Kind.RATIO, "Modified duration at the open; the risk used to explain the period.", group="Risk"),
    _C("duration_current", "ModDur T0", Kind.RATIO, "Modified duration at the close; the risk to be hedged next.", group="Risk"),
    _C("duration_source", "Duration Source", Kind.TEXT, "Whether the opening duration came from the feed or was modelled.", group="Risk"),
    _C("convexity", "Convexity", Kind.RATIO, "Second derivative of price to yield, per unit price.", "(P+ + P- - 2P0) / (P0 * dy^2)", group="Risk"),
    _C("dv01_unit_prior", "DV01 Unit T-1", Kind.RATIO, "Base-currency DV01 per unit of nominal, at the open.", "duration * dirty_px/100 * fx * 1bp", group="Risk"),
    _C("dv01_prior", "DV01 T-1", Kind.EUR, "Position DV01 at the open. Signed by notional.", "notional * dv01_unit_prior", group="Risk"),
    _C("dv01_current", "DV01 T0", Kind.EUR, "Position DV01 at the close; used for every hedging metric.", "notional * dv01_unit_current", group="Risk"),
    _C("attribution_dv01", "Attribution DV01", Kind.EUR, "The DV01 every first-order leg is struck on, per model.attribution_risk_date.", group="Risk"),
    _C("dv01_timing_bias", "DV01 Timing Bias", Kind.EUR, "What striking risk at the close instead of the open would have added. The legacy convention's cost, measured.", "-(DV01_T0 - DV01_T-1) * delta_y", group="Risk"),

    # levels
    _C("ytm_prior", "YTM T-1", Kind.RATE, "Yield to maturity at the open.", group="Levels"),
    _C("ytm_current", "YTM T0", Kind.RATE, "Yield to maturity at the close.", group="Levels"),
    _C("ois_prior", "OIS T-1", Kind.RATE, "Risk-free OIS zero rate at the bond's tenor, at the open. The 'r' leg.", group="Levels"),
    _C("ois_current", "OIS T0", Kind.RATE, "OIS zero rate at the close.", group="Levels"),
    _C("gov_prior", "Gov T-1", Kind.RATE, "Government benchmark rate at the bond's tenor, at the open.", group="Levels"),
    _C("gov_current", "Gov T0", Kind.RATE, "Government benchmark rate at the close.", group="Levels"),
    _C("swap_prior", "Swap T-1", Kind.RATE, "IBOR swap rate at the bond's tenor, at the open.", group="Levels"),
    _C("swap_current", "Swap T0", Kind.RATE, "IBOR swap rate at the close.", group="Levels"),
    _C("ispread_prior", "I-Spread T-1", Kind.BP, "Yield over the swap curve at the open.", "(ytm - swap) * 100", group="Levels"),
    _C("ispread_current", "I-Spread T0", Kind.BP, "Yield over the swap curve at the close.", group="Levels"),
    _C("gspread_prior", "G-Spread T-1", Kind.BP, "Yield over the government curve at the open.", "(ytm - gov) * 100", group="Levels"),
    _C("gspread_current", "G-Spread T0", Kind.BP, "Yield over the government curve at the close.", group="Levels"),

    # moves
    _C("delta_y_bp", "d Yield", Kind.BP, "Total yield move. Every framework chain must reproduce -DV01 times this.", group="Moves"),
    _C("delta_r_bp", "d OIS (r)", Kind.BP, "Risk-free rate move.", group="Moves"),
    _C("delta_gov_bp", "d Gov", Kind.BP, "Government curve move; what a bond future hedges.", group="Moves"),
    _C("delta_swap_bp", "d Swap", Kind.BP, "Swap curve move; what an IBOR swap hedges.", group="Moves"),
    _C("delta_g_bp", "d g (Gov-OIS)", Kind.BP, "Sovereign / collateral basis move.", "(Gov - OIS) T0 less T-1", group="Moves"),
    _C("delta_q_bp", "d q (Swap-Gov)", Kind.BP, "Swap / government basis move.", "(Swap - Gov) T0 less T-1", group="Moves"),
    _C("delta_ispread_bp", "d I-Spread", Kind.BP, "Credit move measured over swaps.", group="Moves"),
    _C("delta_gspread_bp", "d G-Spread", Kind.BP, "Credit move measured over governments.", group="Moves"),
    _C("delta_zspread_bp", "d Z-Spread", Kind.BP, "Z-spread move, as quoted.", group="Moves"),
    _C("delta_asw_bp", "d ASW", Kind.BP, "Asset-swap spread move, as quoted.", group="Moves"),
    _C("delta_oas_bp", "d OAS", Kind.BP, "Option-adjusted spread move, as quoted.", group="Moves"),

    # framework
    _C("spread_framework", "Framework", Kind.TEXT, "Which spread the credit leg is measured against. Changes the split, never the total.", group="Framework"),
    _C("spread_framework_reason", "Framework Reason", Kind.TEXT, "Why that framework was chosen, and by which precedence rule.", group="Framework"),
    _C("framework_futures_weight", "Futures Weight", Kind.PCT, "Share of hedge DV01 that is futures. Drives the automatic choice and the MIXED blend.", group="Framework"),
    _C("framework_swap_weight", "Swap Weight", Kind.PCT, "Share of hedge DV01 that is plain swaps.", group="Framework"),

    # duration legs
    _C("pnl_ois", "PnL OIS", Kind.EUR, "PnL from the risk-free rate move.", "-DV01 * d r", group="Duration"),
    _C("pnl_gov_basis", "PnL Gov Basis", Kind.EUR, "PnL from the sovereign basis move.", "-DV01 * d g", group="Duration"),
    _C("pnl_swap_gov_basis", "PnL Swap-Gov Basis", Kind.EUR, "PnL from the swap/government basis move.", "-DV01 * d q", group="Duration"),
    _C("pnl_ispread", "PnL I-Spread", Kind.EUR, "Credit PnL measured over swaps.", "-DV01 * d i", group="Duration"),
    _C("pnl_gspread", "PnL G-Spread", Kind.EUR, "Credit PnL measured over governments.", "-DV01 * d G-spread", group="Duration"),
    _C("pnl_zspread", "PnL Z-Spread", Kind.EUR, "Credit PnL on the Z-spread convention.", "-DV01 * d Z", group="Duration"),
    _C("pnl_asw", "PnL ASW", Kind.EUR, "Credit PnL on the asset-swap convention.", "-DV01 * d ASW", group="Duration"),
    _C("pnl_oas", "PnL OAS", Kind.EUR, "Credit PnL on the option-adjusted convention.", "-DV01 * d OAS", group="Duration"),
    _C("pnl_yield_only", "PnL Yield Only", Kind.EUR, "The whole yield move, undecomposed. The benchmark every chain must reproduce.", "-DV01 * d y", group="Duration"),
    _C("spread_pnl_used", "Spread PnL Used", Kind.EUR, "The credit leg of the SELECTED chain. Reporting only - already inside the duration total.", group="Duration"),
    _C("pnl_duration_total", "PnL Duration Total", Kind.EUR, "Sum of the selected framework's chain. Enters the bridge.", group="Duration"),
    _C("duration_identity_check", "Identity Check", Kind.EUR, "Chain total less -DV01*dy. A data-quality signal: G and I tie exactly, ASW/Z/OAS approximately.", group="Duration"),
    _C("pnl_convexity", "PnL Convexity", Kind.EUR, "Second-order PnL in the same yield move.", "0.5 * MV_T-1 * convexity * (dy/10000)^2", group="Duration"),

    # carry
    _C("carry_coupon", "Carry Coupon", Kind.EUR, "Coupon earned in the period, accrual plus any coupon actually paid.", "notional/100 * (AI_T0 - AI_T-1 + coupons paid) * fx_prior", group="Carry"),
    _C("carry_roll_to_par", "Carry Roll to Par", Kind.EUR, "Passage of time on the prior curve, priced forward - not a spot reprice at a shorter maturity.", "notional * fx_prior * pull_to_par_price / 100", group="Carry"),
    _C("pull_to_par_price", "Pull to Par /100", Kind.PRICE, "Clean price change per 100 from pure passage of time.", group="Carry"),
    _C("carry_total", "Carry Total", Kind.EUR, "Coupon plus roll to par. Funding is deliberately excluded - see the memo.", group="Carry"),
    _C("coupon_cash", "Coupon Cash", Kind.EUR, "Coupon cash received in the period. On the practical side, matching the dirty-price drop.", group="Carry"),
    _C("funding_carry_memo", "Funding Carry (memo)", Kind.EUR, "Cost of financing the position. Economic carry, not mark-to-market - outside the bridge by default.", "-MV_T-1 * funding rate * year fraction", group="Carry"),

    # fx
    _C("pnl_fx", "PnL FX", Kind.EUR, "Translation of the OPENING position at the new fix.", "MV_local_T-1 * (fx_T0 - fx_T-1)", group="FX"),
    _C("pnl_fx_cross", "PnL FX Cross", Kind.EUR, "Second-order term: the local move revalued at the FX move. Booked explicitly so the FX identity is exact.", "d(MV_local) * d fx", group="FX"),

    # hedges
    _C("futures_count", "Futures #", Kind.COUNT, "Futures linked to this ISIN.", group="Hedge"),
    _C("plain_swap_count", "Plain Swaps #", Kind.COUNT, "Plain swaps linked to this ISIN.", group="Hedge"),
    _C("synthetic_swap_count", "Synthetic #", Kind.COUNT, "Synthetic swap targets linked to this ISIN. Not positions.", group="Hedge"),
    _C("futures_dv01", "Futures DV01", Kind.EUR, "Futures risk attached to this bond, at the close.", "contracts * FUT_PX_VAL_BP * point value * fx", group="Hedge"),
    _C("plain_swap_dv01", "Plain Swap DV01", Kind.EUR, "Plain swap risk attached to this bond.", group="Hedge"),
    _C("plain_swap_dv01_ois", "Swap DV01 (OIS leg)", Kind.EUR, "Swap risk whose floating leg projects off the OIS curve.", group="Hedge"),
    _C("plain_swap_dv01_ibor", "Swap DV01 (IBOR leg)", Kind.EUR, "Swap risk whose floating leg projects off the IBOR swap curve.", group="Hedge"),
    _C("plain_swap_dv01_unknown", "Swap DV01 (unclassified)", Kind.EUR, "Swap risk whose floating index could not be classified. Suppresses the model swap leg rather than guessing a curve.", group="Hedge"),
    _C("synthetic_swap_dv01", "Synthetic DV01", Kind.EUR, "The hedge the coverage relationship says should be on. A target, never a position.", group="Hedge"),
    _C("actual_hedge_dv01", "Actual Hedge DV01", Kind.EUR, "Futures plus plain swaps: the hedge risk really in the book.", group="Hedge"),
    _C("target_hedge_dv01", "Target Hedge DV01", Kind.EUR, "Synthetic where one exists, otherwise flat (-bond DV01).", group="Hedge"),
    _C("hedge_dv01_gap", "Hedge DV01 Gap", Kind.EUR, "Actual less target hedge risk.", group="Hedge"),
    _C("residual_dv01", "Residual DV01", Kind.EUR, "Net risk after hedging. Zero means flat.", "bond DV01 + hedge DV01", group="Hedge"),
    _C("hedge_ratio", "Hedge Ratio", Kind.RATIO, "Signed: 1.0 is fully hedged, negative is wrong-way, above 1 is over-hedged.", "-hedge DV01 / bond DV01", group="Hedge"),
    _C("hedge_efficiency", "Hedge Efficiency", Kind.PCT, "1 means the hedge matches its target. Deliberately not clamped: a negative value is meaningful.", "1 - |gap| / |target|", group="Hedge"),

    _C("model_futures_pnl", "Futures PnL (model)", Kind.EUR, "What the futures hedge should have made from the government curve move.", "-futures DV01 * d Gov", group="Hedge PnL"),
    _C("actual_futures_pnl", "Futures PnL (actual)", Kind.EUR, "Variation margin actually received.", group="Hedge PnL"),
    _C("futures_basis_pnl", "Futures Basis", Kind.EUR, "Actual less model: the CTD/delivery-option basis.", group="Hedge PnL"),
    _C("model_swap_pnl", "Swap PnL (model)", Kind.EUR, "What the swap hedge should have made from its own curve move.", "-swap DV01 * d (its curve)", group="Hedge PnL"),
    _C("actual_plain_swap_pnl", "Swap PnL (actual)", Kind.EUR, "NPV move actually booked.", group="Hedge PnL"),
    _C("swap_basis_pnl", "Swap Basis", Kind.EUR, "Actual less model: swap spread and convention mismatch.", group="Hedge PnL"),
    _C("model_hedge_pnl", "Hedge PnL (model)", Kind.EUR, "Futures plus swap model PnL. Part of theoretical PnL.", group="Hedge PnL"),
    _C("actual_hedge_pnl", "Hedge PnL (actual)", Kind.EUR, "Futures plus swap actual PnL. Part of practical PnL.", group="Hedge PnL"),
    _C("hedge_basis_pnl", "Hedge Basis", Kind.EUR, "Actual less model hedge PnL. Added to explained so the hedge leg telescopes to actual and the residual is bond-leg error only.", group="Hedge PnL"),

    # the bridge
    _C("theoretical_pnl", "Theoretical PnL", Kind.EUR, "What the risk model says the position should have made.", "duration + convexity + carry + FX + model hedge", group="Bridge"),
    _C("total_explained", "Total Explained", Kind.EUR, "Theoretical plus hedge basis: the full explained figure.", "theoretical + hedge basis", group="Bridge"),
    _C("practical_pnl", "Practical PnL", Kind.EUR, "What the position actually made, from marks and cash.", "d MV + actual hedge PnL + coupon cash", group="Bridge"),
    _C("residual_pnl", "Residual", Kind.EUR, "Practical less explained. Bond-leg model error only.", group="Bridge"),
    _C("model_residual_pnl", "Model Residual", Kind.EUR, "Practical less theoretical: includes the hedge basis, so it measures the pure model.", group="Bridge"),
    _C("residual_pct", "Residual %", Kind.PCT, "Residual as a share of absolute practical PnL.", group="Bridge"),

    # quality
    _C("shared_isin", "Shared ISIN", Kind.FLAG, "This ISIN appears on more than one row, so both rows claim its full hedge PnL and DV01.", group="Quality"),
    _C("attribution_status", "Status", Kind.TEXT, "First failing check, ordered cause before symptom.", group="Quality"),
)


FUTURES_SCHEMA: tuple[Column, ...] = (
    _C("contract_code", "Contract", Kind.TEXT, "Exchange contract code.", group="Identity"),
    _C("exchange", "Exchange", Kind.TEXT, "Listing venue.", group="Identity"),
    _C("currency", "CCY", Kind.TEXT, "Contract currency.", group="Identity"),
    _C("portfolio", "Portfolio", Kind.TEXT, "Owning portfolio.", group="Identity"),
    _C("linked_isin", "Linked ISIN", Kind.TEXT, "Bond this future hedges. Blank means it appears on no bond row.", group="Identity"),
    _C("hedge_type", "Hedge Type", Kind.TEXT, "Coverage relationship type.", group="Identity"),
    _C("ctd_isin", "CTD ISIN", Kind.TEXT, "Cheapest-to-deliver bond; what the contract actually tracks.", group="Identity"),
    _C("ctd_cf", "CTD CF", Kind.RATIO, "Exchange conversion factor for the CTD.", group="Identity"),
    _C("deliv_date", "Delivery", Kind.DATE, "First delivery date.", group="Identity"),
    _C("contracts", "Contracts", Kind.NOTIONAL, "Signed contract count. Negative is short.", group="Position"),
    _C("fut_px_prior", "Fut Px T-1", Kind.PRICE, "Futures price at the open.", group="Position"),
    _C("fut_px_current", "Fut Px T0", Kind.PRICE, "Futures price at the close.", group="Position"),
    _C("fut_val_pt", "Value / Point", Kind.RATIO, "Cash per full price point per contract.", group="Position"),
    _C("fut_px_val_bp", "Px Val / bp", Kind.RATIO, "Price points per basis point. Needs the point value to become cash - see config.", group="Position"),
    _C("fx_prior", "FX T-1", Kind.RATIO, "Base per contract currency at the open.", group="Position"),
    _C("fx_current", "FX T0", Kind.RATIO, "Base per contract currency at the close.", group="Position"),
    _C("notional_value_base", "Exposure", Kind.EUR, "Market exposure, not a market value: a future is margined daily and has none.", "contracts * point value * price * fx", group="Position"),
    _C("tenor_years", "Tenor", Kind.RATIO, "Curve point the contract responds to, taken from the CTD.", group="Risk"),
    _C("tenor_source", "Tenor Source", Kind.TEXT, "Whether the tenor came from the CTD, the linked bond, or could not be resolved.", group="Risk"),
    _C("gov_prior", "Gov T-1", Kind.RATE, "Government rate at that tenor, at the open.", group="Risk"),
    _C("gov_current", "Gov T0", Kind.RATE, "Government rate at that tenor, at the close.", group="Risk"),
    _C("delta_gov_bp", "d Gov", Kind.BP, "Government curve move the contract hedged.", group="Risk"),
    _C("dv01_prior", "DV01 T-1", Kind.EUR, "Contract risk at the open.", group="Risk"),
    _C("dv01_current", "DV01 T0", Kind.EUR, "Contract risk at the close.", group="Risk"),
    _C("gross_basis", "Gross Basis", Kind.PRICE, "CTD clean price less futures price times conversion factor. A clean-price concept.", "CTD clean - futures * CF", group="Basis"),
    _C("gross_basis_quality", "Basis Quality", Kind.TEXT, "Whether the CTD's accrued was known, so the basis is a clean-price figure.", group="Basis"),
    _C("gross_basis_bbg", "Gross Basis (feed)", Kind.PRICE, "Gross basis as published by the market-data feed.", group="Basis"),
    _C("net_basis_bbg", "Net Basis (feed)", Kind.PRICE, "Basis after carry, as published.", group="Basis"),
    _C("implied_repo_bbg", "Implied Repo (feed)", Kind.PCT, "Repo rate implied by buying the CTD and selling the future.", group="Basis"),
    _C("theoretical_pnl", "Theoretical PnL", Kind.EUR, "What the contract should have made from the government curve move.", "-DV01 * d Gov", group="Bridge"),
    _C("practical_pnl", "Practical PnL", Kind.EUR, "Variation margin actually received.", "contracts * point value * price move * fx", group="Bridge"),
    _C("basis_pnl", "Basis PnL", Kind.EUR, "Actual less model: what the CTD switch and delivery option were worth.", group="Bridge"),
    _C("is_linked", "Linked", Kind.FLAG, "Whether this contract reaches a bond row at all.", group="Quality"),
    _C("status", "Status", Kind.TEXT, "First failing check.", group="Quality"),
)


SWAPS_SCHEMA: tuple[Column, ...] = (
    _C("deal_id", "Deal", Kind.TEXT, "Swap identifier.", group="Identity"),
    _C("currency", "CCY", Kind.TEXT, "Deal currency.", group="Identity"),
    _C("portfolio", "Portfolio", Kind.TEXT, "Owning portfolio.", group="Identity"),
    _C("counterparty", "Counterparty", Kind.TEXT, "Trade counterparty.", group="Identity"),
    _C("linked_isin", "Linked ISIN", Kind.TEXT, "Bond this swap hedges. Blank means it appears on no bond row.", group="Identity"),
    _C("swap_id_source", "Class", Kind.TEXT, "PLAIN is a real position. SYNTHETIC is a hedge target and is excluded from every total.", group="Identity"),
    _C("notional", "Notional", Kind.NOTIONAL, "Deal notional.", group="Position"),
    _C("pay_fixed", "Pay Fixed", Kind.TEXT, "Y means the desk pays fixed, so the deal is short duration.", group="Position"),
    _C("fixed_rate", "Fixed Rate", Kind.RATE, "Contractual fixed rate.", group="Position"),
    _C("float_index", "Float Index", Kind.TEXT, "Floating reference index as traded.", group="Position"),
    _C("float_spread", "Float Spread", Kind.RATE, "Contractual spread on the floating leg.", group="Position"),
    _C("float_curve_type", "Float Family", Kind.TEXT, "Curve the floating leg projects off: ESTR/SOFR/SONIA use OIS, EURIBOR uses the swap curve.", group="Position"),
    _C("float_family_status", "Family Status", Kind.TEXT, "Flags an index that does not belong to the deal's currency.", group="Position"),
    _C("start_date", "Start", Kind.DATE, "Effective date.", group="Position"),
    _C("end_date", "End", Kind.DATE, "Maturity date.", group="Position"),
    _C("year_fraction", "Tenor", Kind.RATIO, "Remaining life in years; the curve point the deal is read at.", group="Risk"),
    _C("ois_prior", "OIS T-1", Kind.RATE, "OIS rate at the deal tenor, at the open.", group="Risk"),
    _C("ois_current", "OIS T0", Kind.RATE, "OIS rate at the deal tenor, at the close.", group="Risk"),
    _C("float_curve_prior", "Float Curve T-1", Kind.RATE, "The curve this deal's floating leg actually references, at the open.", group="Risk"),
    _C("float_curve_current", "Float Curve T0", Kind.RATE, "The same curve at the close.", group="Risk"),
    _C("delta_float_curve_bp", "d Float Curve", Kind.BP, "Move in the curve the deal is exposed to.", group="Risk"),
    _C("model_spread_prior", "Model Spread T-1", Kind.RATE, "Fixed rate plus float spread less the projection curve, at the open.", group="Risk"),
    _C("model_spread_current", "Model Spread T0", Kind.RATE, "The same at the close.", group="Risk"),
    _C("delta_model_spread_bp", "d Model Spread", Kind.BP, "The part of the deal's value change the curve does not explain.", group="Risk"),
    _C("dv01", "DV01", Kind.EUR, "Signed risk in base currency per basis point.", group="Risk"),
    _C("dv01_model", "DV01 (annuity model)", Kind.EUR, "Internal annuity approximation, for comparison with the supplied figure.", "sign * |notional| * 1bp * (1-DF)/z * fx", group="Risk"),
    _C("dv01_source", "DV01 Source", Kind.TEXT, "Whether the supplied risk or the internal model was used.", group="Risk"),
    _C("npv_prior", "NPV T-1", Kind.EUR, "Deal present value at the open.", group="Bridge"),
    _C("npv_current", "NPV T0", Kind.EUR, "Deal present value at the close.", group="Bridge"),
    _C("theoretical_pnl", "Theoretical PnL", Kind.EUR, "What the deal should have made from its own curve move.", "-DV01 * d float curve", group="Bridge"),
    _C("practical_pnl", "Practical PnL", Kind.EUR, "NPV move actually booked.", "NPV T0 - NPV T-1", group="Bridge"),
    _C("basis_pnl", "Basis PnL", Kind.EUR, "Actual less model: spread and convention mismatch.", group="Bridge"),
    _C("is_linked", "Linked", Kind.FLAG, "Whether this deal reaches a bond row at all.", group="Quality"),
    _C("is_plain", "Plain", Kind.FLAG, "A real position rather than a synthetic target.", group="Quality"),
    _C("status", "Status", Kind.TEXT, "First failing check.", group="Quality"),
)
