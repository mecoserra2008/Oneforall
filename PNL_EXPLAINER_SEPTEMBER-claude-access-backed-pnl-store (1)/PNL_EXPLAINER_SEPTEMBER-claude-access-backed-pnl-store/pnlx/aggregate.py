"""Portfolio aggregation: the bridge, the asset-class comparison, data quality.

Three things are built here, and each answers a different question.

THE BRIDGE           "where did the day's PnL come from?"
    Reads top to bottom from the model to the marks.  Only the duration TOTAL
    enters the sum; the OIS / Gov-OIS / Swap-Gov / credit lines beneath it are
    "of which" memo lines and are deliberately excluded.

    Summing those legs directly is double counting, and it is worth being
    precise about why: WHICH legs make up a row's duration total depends on
    that row's framework.  A G-framework row has no Swap-Gov leg in its chain
    at all, so summing the Swap-Gov column across every row inflates the bridge
    by exactly the G rows' share of it.  The memo lines are therefore summed
    only over the rows whose chain actually contains them, and a balancing line
    carries the OIS/SOFR and MIXED rows that have no per-leg split.

    Two tie-out rows prove the arithmetic rather than asserting it.

ASSET-CLASS COMPARISON   "which instrument type made it, and did each behave?"
    Bonds, futures and swaps each get their own theoretical, practical and
    residual, computed on that class's own terms, plus what is linked to a bond
    and what is not.  This is the view the legacy workbook could not produce:
    everything there was a bond row, so a hedge with no LinkedISIN simply did
    not appear.

DATA QUALITY         "how much of the book is this actually covering?"
    A total that quietly covers 80% of the book is the failure mode worth
    guarding against, so coverage is reported next to every headline figure.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Sequence

import numpy as np

from .columns import ColumnFrame
from .config import AppConfig
from .engine import AttributionResult
from .frameworks import CHAINS

__all__ = [
    "BridgeLine",
    "Bridge",
    "AssetClassSummary",
    "PortfolioSummary",
    "build_portfolio",
]

#: Frameworks whose chain contains each memo leg.  MIXED is in none of them:
#: its duration total is a DV01-weighted blend, so its legs do not enter at
#: weight 1 and attributing them in full would overstate the memo.  A blended
#: row has no honest per-leg split, so its whole total falls into the balancing
#: line, alongside the undecomposed OIS/SOFR rows.
_CURVE_FRAMEWORKS = ("G", "I", "ASW", "Z", "OAS")
_SWAPGOV_FRAMEWORKS = ("I", "ASW", "Z", "OAS")
_SPREAD_FRAMEWORKS = ("G", "I", "ASW", "Z", "OAS", "OIS", "SOFR")


@dataclass(frozen=True, slots=True)
class BridgeLine:
    """One line of the PnL bridge."""

    label: str
    amount: float
    kind: str = "component"      # component | memo | subtotal | total | check
    note: str = ""

    @property
    def is_memo(self) -> bool:
        return self.kind == "memo"

    @property
    def in_sum(self) -> bool:
        return self.kind == "component"


@dataclass(slots=True)
class Bridge:
    """The portfolio bridge, from the model to the marks."""

    lines: list[BridgeLine] = field(default_factory=list)

    def add(self, label: str, amount: float, kind: str = "component", note: str = "") -> None:
        self.lines.append(BridgeLine(label, amount, kind, note))

    @property
    def components(self) -> list[BridgeLine]:
        return [line for line in self.lines if line.in_sum]

    @property
    def memos(self) -> list[BridgeLine]:
        return [line for line in self.lines if line.is_memo]

    def value(self, label: str) -> float:
        for line in self.lines:
            if line.label == label:
                return line.amount
        return float("nan")

    def waterfall(self) -> list[tuple[str, float, float, float]]:
        """`(label, start, end, delta)` per component, for a waterfall chart.

        Cumulative from zero through every component, so the bars sit on top of
        one another and the last one lands on the explained total.
        """
        out: list[tuple[str, float, float, float]] = []
        running = 0.0
        for line in self.components:
            amount = line.amount if np.isfinite(line.amount) else 0.0
            out.append((line.label, running, running + amount, amount))
            running += amount
        return out

    def to_records(self) -> list[dict[str, object]]:
        return [
            {
                "label": line.label,
                "amount": line.amount,
                "kind": line.kind,
                "note": line.note,
            }
            for line in self.lines
        ]


@dataclass(frozen=True, slots=True)
class AssetClassSummary:
    """One instrument type, on its own terms."""

    name: str
    count: int
    linked_count: int
    unlinked_count: int
    exposure: float
    dv01: float
    theoretical: float
    practical: float
    basis: float
    residual: float
    unlinked_practical: float
    coverage: float
    note: str = ""

    def to_record(self) -> dict[str, object]:
        return {
            "asset_class": self.name,
            "positions": self.count,
            "linked": self.linked_count,
            "unlinked": self.unlinked_count,
            "exposure": self.exposure,
            "dv01": self.dv01,
            "theoretical_pnl": self.theoretical,
            "practical_pnl": self.practical,
            "basis_pnl": self.basis,
            "residual_pnl": self.residual,
            "unlinked_practical_pnl": self.unlinked_practical,
            "coverage": self.coverage,
            "note": self.note,
        }


@dataclass(slots=True)
class PortfolioSummary:
    """Everything the report needs above position level."""

    bridge: Bridge
    asset_classes: list[AssetClassSummary]
    headline: dict[str, float]
    quality: dict[str, object]
    framework_mix: list[dict[str, object]]
    by_portfolio: list[dict[str, object]]
    by_currency: list[dict[str, object]]
    config: AppConfig

    def asset_class_records(self) -> list[dict[str, object]]:
        return [ac.to_record() for ac in self.asset_classes]


# --------------------------------------------------------------------------- #


def _sum_where(frame: ColumnFrame, column: str, mask: np.ndarray) -> float:
    if len(frame) == 0:
        return 0.0
    values = np.asarray(frame.get(column), dtype=np.float64)
    return float(np.nansum(values[mask])) if mask.any() else 0.0


def build_portfolio(result: AttributionResult) -> PortfolioSummary:
    """Aggregate an attribution result into everything above position level."""
    cfg = result.config
    bonds = result.bonds
    futures = result.futures
    swaps = result.swaps

    n_bonds = len(bonds)
    framework = bonds.get("spread_framework", np.array([], dtype=object))

    # Rows that produced a complete explained figure.  Every bridge component
    # must be summed over the SAME rows as the total it is compared against; a
    # component summed over all rows against a total built only from rows that
    # fully attributed cannot add up, and the difference looks like a broken
    # model when it is really missing data.
    attributed = np.isfinite(np.asarray(bonds.get("total_explained"), dtype=np.float64))
    has_practical = np.isfinite(np.asarray(bonds.get("practical_pnl"), dtype=np.float64))

    def att(column: str) -> float:
        return _sum_where(bonds, column, attributed)

    bridge = Bridge()

    # ---- the model side ---------------------------------------------------- #
    duration_total = att("pnl_duration_total")
    bridge.add(
        "Rates and spread (duration)",
        duration_total,
        note="Sum of the selected framework's chain. The lines below are 'of which'.",
    )

    memo_legs = [
        ("of which OIS rate", "pnl_ois", _CURVE_FRAMEWORKS),
        ("of which Gov-OIS basis", "pnl_gov_basis", _CURVE_FRAMEWORKS),
        ("of which Swap-Gov basis", "pnl_swap_gov_basis", _SWAPGOV_FRAMEWORKS),
        ("of which credit / spread", "spread_pnl_used", _SPREAD_FRAMEWORKS),
    ]
    memo_total = 0.0
    for label, column, codes in memo_legs:
        mask = attributed & np.isin(framework, np.array(codes, dtype=object))
        amount = _sum_where(bonds, column, mask)
        memo_total += amount
        bridge.add(label, amount, kind="memo")

    # OIS/SOFR rows are undecomposed and MIXED rows are a weighted blend, so the
    # four lines above cannot tie to the duration total on their own.  This is
    # the remainder, which makes the memo block tie exactly rather than nearly.
    bridge.add(
        "of which not separately split (OIS/SOFR/MIXED)",
        duration_total - memo_total,
        kind="memo",
    )

    bridge.add("Convexity", att("pnl_convexity"),
               note="Second order in the same yield move.")
    bridge.add("Carry (coupon + roll to par)", att("carry_total"),
               note="Funding is an economic-carry memo, not part of a mark-to-market bridge.")
    bridge.add("FX translation", att("pnl_fx"),
               note="Revaluation of the opening position at the new fix.")
    bridge.add("FX cross term", att("pnl_fx_cross"),
               note="Local move revalued at the FX move. Booked so the FX identity is exact.")
    bridge.add("Hedge PnL (model)", att("model_hedge_pnl"),
               note="Hedge DV01 against the curve each hedge actually neutralises.")

    theoretical = att("theoretical_pnl")
    bridge.add("Theoretical PnL", theoretical, kind="subtotal",
               note="What the risk model says the book should have made.")

    bridge.add("Hedge basis (actual - model)", att("hedge_basis_pnl"),
               note="CTD/delivery basis and swap spread mismatch. In the bridge so the "
                    "hedge leg telescopes to actual and the residual is bond-leg error only.")

    explained = att("total_explained")
    bridge.add("Total explained", explained, kind="subtotal")

    bridge.add("Unexplained residual", att("residual_pnl"),
               note="Practical less explained: bond-leg model error.")

    # A row with no explained figure still HAS a practical one.  It appears in
    # the practical total and in neither explained nor residual, and without
    # this line the portfolio bridge fails to add up even though every
    # individual row's bridge is correct.
    not_attributed = _sum_where(bonds, "practical_pnl", has_practical & ~attributed)
    bridge.add("Not attributed (no explained PnL)", not_attributed,
               note="Rows whose model could not be completed. Their practical PnL is real.")

    practical = _sum_where(bonds, "practical_pnl", has_practical)
    bridge.add("Practical PnL (bonds + linked hedges)", practical, kind="total",
               note="Change in dirty market value, plus coupon cash, plus actual hedge PnL.")

    # ---- tie-outs ---------------------------------------------------------- #
    component_sum = sum(
        line.amount for line in bridge.components if np.isfinite(line.amount)
    )
    bridge.add(
        "Check: components less Total explained (want 0)",
        component_sum - explained - att("residual_pnl") - not_attributed,
        kind="check",
    )
    bridge.add(
        "Check: Explained + Residual + Not attributed less Practical (want 0)",
        explained + att("residual_pnl") + not_attributed - practical,
        kind="check",
    )
    bridge.add(
        "Check: memo lines less duration total (want 0)",
        sum(line.amount for line in bridge.memos if np.isfinite(line.amount)) - duration_total,
        kind="check",
    )

    # ---- outside the bridge ------------------------------------------------- #
    bridge.add("Memo: funding carry (not in bridge)", att("funding_carry_memo"), kind="memo",
               note="Cost of financing. Practical PnL is a mark-to-market figure with no "
                    "financing leg, so adding this would open a residual equal to it.")
    bridge.add("Memo: unlinked hedge PnL (on no bond row)",
               float(result.linkage.unlinked_futures_pnl + result.linkage.unlinked_swap_pnl),
               kind="memo",
               note="Hedges with no LinkedISIN. Real PnL the bond-level totals cannot see.")
    bridge.add("Memo: DV01 timing bias avoided", att("dv01_timing_bias"), kind="memo",
               note="What striking risk at the close instead of the open would have added.")

    # ---- asset classes ------------------------------------------------------ #
    asset_classes = _asset_classes(result, attributed, has_practical)

    # ---- headline ----------------------------------------------------------- #
    total_practical = sum(ac.practical for ac in asset_classes)
    total_theoretical = sum(ac.theoretical for ac in asset_classes)

    headline = {
        "practical_pnl": practical,
        "theoretical_pnl": theoretical,
        "total_explained": explained,
        "residual_pnl": att("residual_pnl"),
        "not_attributed": not_attributed,
        "residual_pct": (
            att("residual_pnl") / abs(practical) if practical else float("nan")
        ),
        "duration_total": duration_total,
        "convexity": att("pnl_convexity"),
        "carry_total": att("carry_total"),
        "carry_coupon": att("carry_coupon"),
        "carry_roll_to_par": att("carry_roll_to_par"),
        "funding_carry_memo": att("funding_carry_memo"),
        "pnl_fx": att("pnl_fx") + att("pnl_fx_cross"),
        "model_hedge_pnl": att("model_hedge_pnl"),
        "actual_hedge_pnl": att("actual_hedge_pnl"),
        "hedge_basis_pnl": att("hedge_basis_pnl"),
        "bond_dv01": bonds.total("dv01_current"),
        "hedge_dv01": bonds.total("actual_hedge_dv01"),
        "residual_dv01": bonds.total("residual_dv01"),
        "mv_prior": bonds.total("mv_prior_base"),
        "mv_current": bonds.total("mv_current_base"),
        "all_class_practical": total_practical,
        "all_class_theoretical": total_theoretical,
        "unlinked_hedge_pnl": float(
            result.linkage.unlinked_futures_pnl + result.linkage.unlinked_swap_pnl
        ),
        "dv01_timing_bias": att("dv01_timing_bias"),
        "portfolio_hedge_efficiency": _portfolio_hedge_efficiency(bonds),
        "median_hedge_efficiency": _median(bonds.get("hedge_efficiency")),
    }

    # ---- quality ------------------------------------------------------------ #
    quality = dict(result.diagnostics)
    quality.update(
        {
            "coverage_explained": float(attributed.sum()) / n_bonds if n_bonds else 0.0,
            "coverage_practical": float(has_practical.sum()) / n_bonds if n_bonds else 0.0,
            "bridge_tie_out": explained + att("residual_pnl") + not_attributed - practical,
            "memo_tie_out": (
                sum(line.amount for line in bridge.memos[:5] if np.isfinite(line.amount))
                - duration_total
            ),
            "identity_break_rows": int(
                (
                    np.abs(np.nan_to_num(np.asarray(bonds.get("duration_identity_check"), float)))
                    > np.maximum(
                        cfg.tolerances.identity_eur,
                        cfg.tolerances.identity_pct
                        * np.abs(np.nan_to_num(np.asarray(bonds.get("pnl_duration_total"), float))),
                    )
                ).sum()
            ),
        }
    )

    return PortfolioSummary(
        bridge=bridge,
        asset_classes=asset_classes,
        headline=headline,
        quality=quality,
        framework_mix=_framework_mix(bonds, result),
        by_portfolio=_group_by(bonds, "portfolio"),
        by_currency=_group_by(bonds, "currency"),
        config=cfg,
    )


# --------------------------------------------------------------------------- #


def _asset_classes(
    result: AttributionResult, attributed: np.ndarray, has_practical: np.ndarray
) -> list[AssetClassSummary]:
    """Bonds, futures and swaps each measured on their own terms.

    The bond line is the BOND LEG only - its own theoretical and practical PnL
    with the hedge legs removed - so the three lines add up to the book without
    counting the hedges twice.  A hedge appears once, in its own class.
    """
    bonds = result.bonds
    futures = result.futures
    swaps = result.swaps
    out: list[AssetClassSummary] = []

    # ---- bonds (the cash leg alone) ---------------------------------------- #
    #
    # BOTH sides are summed over the SAME rows - the ones that produced a
    # complete explained figure.  Mixing the masks (theoretical over attributed
    # rows, practical over all rows with marks) makes this residual disagree
    # with the bridge's by exactly the unattributed rows' market-value move,
    # which reads as a model failure and is only a population mismatch.
    #
    # Because the hedge legs telescope in the bridge, summing on one mask makes
    # this line's residual EQUAL the bridge's unexplained residual, which is a
    # checkable statement rather than a coincidence.
    n = len(bonds)
    if n:
        bond_theoretical = (
            _sum_where(bonds, "pnl_duration_total", attributed)
            + _sum_where(bonds, "pnl_convexity", attributed)
            + _sum_where(bonds, "carry_total", attributed)
            + _sum_where(bonds, "pnl_fx", attributed)
            + _sum_where(bonds, "pnl_fx_cross", attributed)
        )
        bond_practical = (
            _sum_where(bonds, "delta_mv_base", attributed)
            + _sum_where(bonds, "coupon_cash", attributed)
        )
        not_attributed_practical = _sum_where(
            bonds, "practical_pnl", has_practical & ~attributed
        )
        out.append(
            AssetClassSummary(
                name="Bonds (cash leg)",
                count=n,
                linked_count=int(attributed.sum()),
                unlinked_count=int((has_practical & ~attributed).sum()),
                exposure=bonds.total("mv_current_base"),
                dv01=bonds.total("dv01_current"),
                theoretical=bond_theoretical,
                practical=bond_practical,
                basis=0.0,
                residual=bond_practical - bond_theoretical,
                unlinked_practical=not_attributed_practical,
                coverage=float(attributed.sum()) / n,
                note="Market value move plus coupon cash, against duration + convexity "
                     "+ carry + FX. Hedges are excluded and appear on their own lines, "
                     "so this residual equals the bridge's unexplained residual. "
                     "'Unlinked' here counts rows that could not be attributed.",
            )
        )

    # ---- futures ------------------------------------------------------------ #
    nf = len(futures)
    if nf:
        linked = np.asarray(futures.get("is_linked"), dtype=bool)
        prac = np.asarray(futures.get("practical_pnl"), dtype=np.float64)
        theo = np.asarray(futures.get("theoretical_pnl"), dtype=np.float64)
        covered = np.isfinite(theo)
        out.append(
            AssetClassSummary(
                name="Bond futures",
                count=nf,
                linked_count=int(linked.sum()),
                unlinked_count=int((~linked).sum()),
                exposure=futures.total("notional_value_base"),
                dv01=futures.total("dv01_current"),
                theoretical=float(np.nansum(theo)),
                practical=float(np.nansum(prac)),
                basis=futures.total("basis_pnl"),
                residual=futures.total("basis_pnl"),
                unlinked_practical=float(np.nansum(prac[~linked])) if nf else 0.0,
                coverage=float(covered.sum()) / nf,
                note="Variation margin against the government curve move at the CTD's "
                     "tenor. The gap is the CTD switch and delivery option.",
            )
        )

    # ---- swaps -------------------------------------------------------------- #
    ns = len(swaps)
    if ns:
        plain = np.asarray(swaps.get("is_plain"), dtype=bool)
        linked = np.asarray(swaps.get("is_linked"), dtype=bool) & plain
        prac = np.asarray(swaps.get("practical_pnl"), dtype=np.float64)
        theo = np.asarray(swaps.get("theoretical_pnl"), dtype=np.float64)
        covered = np.isfinite(theo) & plain
        out.append(
            AssetClassSummary(
                name="Interest-rate swaps",
                count=int(plain.sum()),
                linked_count=int(linked.sum()),
                unlinked_count=int((plain & ~linked).sum()),
                exposure=_sum_where(swaps, "notional", plain),
                dv01=_sum_where(swaps, "dv01", plain),
                theoretical=_sum_where(swaps, "theoretical_pnl", plain),
                practical=_sum_where(swaps, "practical_pnl", plain),
                basis=_sum_where(swaps, "basis_pnl", plain),
                residual=_sum_where(swaps, "basis_pnl", plain),
                unlinked_practical=_sum_where(swaps, "practical_pnl", plain & ~linked),
                coverage=float(covered.sum()) / max(int(plain.sum()), 1),
                note="NPV move against the curve each deal's floating leg projects off. "
                     "Synthetic targets are excluded from every figure here.",
            )
        )

        synthetic = np.asarray(swaps.get("is_plain"), dtype=bool) == False  # noqa: E712
        if synthetic.any():
            out.append(
                AssetClassSummary(
                    name="Synthetic swap targets",
                    count=int(synthetic.sum()),
                    linked_count=int((synthetic & np.asarray(swaps.get("is_linked"), bool)).sum()),
                    unlinked_count=0,
                    exposure=_sum_where(swaps, "notional", synthetic),
                    dv01=_sum_where(swaps, "dv01", synthetic),
                    theoretical=0.0,
                    practical=0.0,
                    basis=0.0,
                    residual=0.0,
                    unlinked_practical=0.0,
                    coverage=1.0,
                    note="Not positions. The hedge the coverage relationship says should be "
                         "on, used only as the target in hedge efficiency. Carries risk in "
                         "this table and PnL nowhere.",
                )
            )

    return out


def _framework_mix(bonds: ColumnFrame, result: AttributionResult) -> list[dict[str, object]]:
    """How the book is being measured, weighted by risk not by row count.

    Weighting by absolute DV01 is the point: ten small G-framework lines and one
    enormous I-framework line is a book measured against swaps, and a headcount
    would say the opposite.
    """
    if len(bonds) == 0:
        return []
    codes = np.asarray(bonds.get("spread_framework"), dtype=object)
    dv01 = np.abs(np.nan_to_num(np.asarray(bonds.get("dv01_current"), dtype=np.float64)))
    duration = np.asarray(bonds.get("pnl_duration_total"), dtype=np.float64)
    total_dv01 = dv01.sum()

    rows: list[dict[str, object]] = []
    for code in sorted(set(codes.tolist())):
        mask = codes == code
        chain = CHAINS.get(str(code))
        rows.append(
            {
                "framework": code,
                "positions": int(mask.sum()),
                "dv01": float(dv01[mask].sum()),
                "dv01_share": float(dv01[mask].sum() / total_dv01) if total_dv01 else 0.0,
                "duration_pnl": float(np.nansum(duration[mask])),
                "ties_exactly": bool(chain.exact) if chain else False,
                "chain": " + ".join(chain.legs) if chain and chain.legs else "-",
                "meaning": chain.description if chain else "",
            }
        )
    return rows


def _group_by(bonds: ColumnFrame, key: str) -> list[dict[str, object]]:
    """Headline figures split by portfolio or currency."""
    if len(bonds) == 0:
        return []
    keys = np.asarray(bonds.get(key), dtype=object)
    rows: list[dict[str, object]] = []
    for value in sorted(set(str(k) for k in keys.tolist())):
        mask = keys == value
        practical = _sum_where(bonds, "practical_pnl", mask)
        theoretical = _sum_where(bonds, "theoretical_pnl", mask)
        explained = _sum_where(bonds, "total_explained", mask)
        rows.append(
            {
                key: value,
                "positions": int(mask.sum()),
                "mv_current": _sum_where(bonds, "mv_current_base", mask),
                "dv01": _sum_where(bonds, "dv01_current", mask),
                "hedge_dv01": _sum_where(bonds, "actual_hedge_dv01", mask),
                "residual_dv01": _sum_where(bonds, "residual_dv01", mask),
                "duration_pnl": _sum_where(bonds, "pnl_duration_total", mask),
                "carry": _sum_where(bonds, "carry_total", mask),
                "convexity": _sum_where(bonds, "pnl_convexity", mask),
                "fx": _sum_where(bonds, "pnl_fx", mask) + _sum_where(bonds, "pnl_fx_cross", mask),
                "hedge_pnl": _sum_where(bonds, "actual_hedge_pnl", mask),
                "theoretical_pnl": theoretical,
                "total_explained": explained,
                "practical_pnl": practical,
                "residual_pnl": _sum_where(bonds, "residual_pnl", mask),
                "residual_pct": (
                    _sum_where(bonds, "residual_pnl", mask) / abs(practical)
                    if practical
                    else float("nan")
                ),
            }
        )
    return rows


def _portfolio_hedge_efficiency(bonds: ColumnFrame) -> float:
    """1 - |actual - target| / |target|, on the NETTED book totals.

    Netted, so offsetting per-bond errors cancel and this always reads better
    than the per-bond distribution.  Both are reported deliberately: the gap
    between them IS the information - a book that is flat in aggregate while
    every line is individually mis-hedged has a real problem that the netted
    number alone would hide.
    """
    if len(bonds) == 0:
        return float("nan")
    actual = bonds.total("actual_hedge_dv01")
    target = bonds.total("target_hedge_dv01")
    if target == 0:
        return float("nan")
    return 1.0 - abs(actual - target) / abs(target)


def _median(values: np.ndarray) -> float:
    arr = np.asarray(values, dtype=np.float64)
    arr = arr[np.isfinite(arr)]
    return float(np.median(arr)) if arr.size else float("nan")
