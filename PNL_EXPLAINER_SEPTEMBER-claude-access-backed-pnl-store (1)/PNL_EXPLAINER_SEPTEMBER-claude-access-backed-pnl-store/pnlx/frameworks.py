"""Spread frameworks: which spread the credit leg is measured against.

WHAT A FRAMEWORK IS
-------------------
A framework is a way of SPLITTING a bond's yield move into a curve part and a
spread part.  It never changes the total.  Every chain below is an exact
telescoping identity in the underlying quoted rates:

    y = r + g + q + i          the decomposition the whole model rests on

        r = OIS zero rate at the bond's tenor
        g = Gov  - OIS         sovereign / collateral basis
        q = Swap - Gov         swap / government basis
        i = y    - Swap        the bond's I-spread

so, writing D for the attribution DV01 and every delta in basis points:

    G     -D*dr + -D*dg + -D*d(G-spread)              == -D*dy   EXACTLY
    I     -D*dr + -D*dg + -D*dq + -D*d(I-spread)      == -D*dy   EXACTLY
    ASW   -D*dr + -D*dg + -D*dq + -D*d(ASW)           ~= -D*dy
    Z     -D*dr + -D*dg + -D*dq + -D*d(Z)             ~= -D*dy
    OAS   -D*dr + -D*dg + -D*dq + -D*d(OAS)           ~= -D*dy
    OIS   -D*dy                                        == -D*dy   trivially
    SOFR  -D*dy                                        == -D*dy   trivially
    MIXED DV01-weighted blend of the G and I chains
    REVIEW  attribution suppressed

G and I tie exactly because G-spread and I-spread are DEFINED as differences
from the same yield.  ASW, Z and OAS are quoted on their own conventions -
asset-swap spread against a floating leg, Z-spread against the whole zero
curve, OAS after stripping optionality - so their chains tie approximately, and
`duration_identity_check` is where that difference shows up.  A large value
there means an input is stale or quoted on an unreconcilable basis; it does not
mean the framework is wrong.

HOW A BOND'S FRAMEWORK IS CHOSEN
--------------------------------
Precedence, highest first:

    1. per-bond override      the overrides file, keyed on ISIN
    2. global override        `model.spread_framework_override` in the YAML
    3. automatic              from the mix of hedges actually attached

The automatic rule follows what the hedge actually neutralises:

    futures-dominated  ->  G.  Bond futures hedge the deliverable GOVERNMENT
                           curve through the CTD, so what is left unhedged is
                           the bond's spread over governments.
    swap-dominated     ->  I.  A swap hedges the SWAP curve, so what is left
                           unhedged is the bond's spread over swaps.
    unhedged           ->  preference order I, G, ASW, Z, OAS, by availability.

Dominance is measured on ABSOLUTE DV01, so it does not depend on the sign
convention of either hedge leg.

Synthetic swaps deliberately do NOT vote.  A synthetic is the hedge the desk
COULD have put on, not the one whose risk is in the book, so letting it choose
the framework would measure the bond against a curve it is not actually
exposed to.

An unrecognised override code resolves to REVIEW rather than being ignored, so
a typo suppresses attribution loudly instead of quietly changing the answer.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Sequence

import numpy as np

__all__ = [
    "FRAMEWORKS",
    "FrameworkChain",
    "CHAINS",
    "FrameworkResolution",
    "resolve_frameworks",
    "curve_for_framework",
    "spread_column_for_framework",
]

FRAMEWORKS: tuple[str, ...] = ("G", "I", "ASW", "Z", "OAS", "OIS", "SOFR", "MIXED", "REVIEW")


@dataclass(frozen=True, slots=True)
class FrameworkChain:
    """The ordered legs whose sum is one framework's duration PnL."""

    code: str
    legs: tuple[str, ...]
    spread_leg: str
    base_curve: str
    exact: bool
    description: str

    def __str__(self) -> str:  # pragma: no cover - display only
        return f"{self.code}: {' + '.join(self.legs)}"


#: Leg names are the attribution column names produced by `engine.py`.
CHAINS: Mapping[str, FrameworkChain] = {
    "G": FrameworkChain(
        code="G",
        legs=("pnl_ois", "pnl_gov_basis", "pnl_gspread"),
        spread_leg="pnl_gspread",
        base_curve="GOV",
        exact=True,
        description="measured vs governments - futures/govie hedge, or G-spread is the strongest available quote",
    ),
    "I": FrameworkChain(
        code="I",
        legs=("pnl_ois", "pnl_gov_basis", "pnl_swap_gov_basis", "pnl_ispread"),
        spread_leg="pnl_ispread",
        base_curve="SWAP",
        exact=True,
        description="measured vs swaps - swap hedge, or the default credit framework",
    ),
    "ASW": FrameworkChain(
        code="ASW",
        legs=("pnl_ois", "pnl_gov_basis", "pnl_swap_gov_basis", "pnl_asw"),
        spread_leg="pnl_asw",
        base_curve="SWAP",
        exact=False,
        description="asset-swap spread - quoted against a floating leg, so the chain ties approximately",
    ),
    "Z": FrameworkChain(
        code="Z",
        legs=("pnl_ois", "pnl_gov_basis", "pnl_swap_gov_basis", "pnl_zspread"),
        spread_leg="pnl_zspread",
        base_curve="SWAP",
        exact=False,
        description="Z-spread - quoted against the whole zero curve, so the chain ties approximately",
    ),
    "OAS": FrameworkChain(
        code="OAS",
        legs=("pnl_ois", "pnl_gov_basis", "pnl_swap_gov_basis", "pnl_oas"),
        spread_leg="pnl_oas",
        base_curve="SWAP",
        exact=False,
        description="option-adjusted spread - optionality stripped out, so the chain ties approximately",
    ),
    "OIS": FrameworkChain(
        code="OIS",
        legs=("pnl_yield_only",),
        spread_leg="pnl_spread_over_ois",
        base_curve="OIS",
        exact=True,
        description="no decomposition - the whole yield move is taken as -DV01 x dy",
    ),
    "SOFR": FrameworkChain(
        code="SOFR",
        legs=("pnl_yield_only",),
        spread_leg="pnl_spread_over_ois",
        base_curve="OIS",
        exact=True,
        description="no decomposition - the whole yield move is taken as -DV01 x dy",
    ),
    "MIXED": FrameworkChain(
        code="MIXED",
        legs=(),
        spread_leg="",
        base_curve="SWAP",
        exact=False,
        description="DV01-weighted blend of the government and swap chains",
    ),
    "REVIEW": FrameworkChain(
        code="REVIEW",
        legs=(),
        spread_leg="",
        base_curve="",
        exact=False,
        description="attribution suppressed - unrecognised override code, or flagged for manual review",
    ),
}

#: Which spread-delta column each framework's credit leg is built from.
_SPREAD_DELTA = {
    "G": "delta_gspread_bp",
    "I": "delta_ispread_bp",
    "ASW": "delta_asw_bp",
    "Z": "delta_zspread_bp",
    "OAS": "delta_oas_bp",
}

#: Fallback order when nothing is hedged: the most reliably quoted first.
_UNHEDGED_PREFERENCE = ("I", "G", "ASW", "Z", "OAS")
_FUTURES_PREFERENCE = ("G", "I", "ASW", "Z", "OAS")
_SWAP_PREFERENCE = ("I", "ASW", "G", "Z", "OAS")


def curve_for_framework(codes: np.ndarray) -> np.ndarray:
    """Base curve each framework's spread is quoted over.

    Used by pull to par: a G-framework bond is repriced on the government
    curve plus its G-spread, an I/ASW/Z/OAS bond on the swap curve plus its
    spread, and an OIS/SOFR bond on the OIS curve.  Repricing on the wrong base
    curve puts the whole government/swap basis into carry.
    """
    out = np.empty(np.asarray(codes).shape, dtype=object)
    for i, code in enumerate(np.asarray(codes).tolist()):
        chain = CHAINS.get(str(code).strip().upper())
        out[i] = chain.base_curve if chain else ""
    return out


def spread_column_for_framework(code: str) -> str:
    """The spread-delta column feeding a framework's credit leg ('' if none)."""
    return _SPREAD_DELTA.get(str(code).strip().upper(), "")


@dataclass(slots=True)
class FrameworkResolution:
    """Per-bond framework choice, with the reason it was made."""

    code: np.ndarray          # object array of framework codes
    source: np.ndarray        # 'per-bond override' / 'global override' / 'automatic'
    reason: np.ndarray        # human-readable sentence
    futures_weight: np.ndarray   # share of hedge DV01 that is futures
    swap_weight: np.ndarray      # share of hedge DV01 that is plain swaps

    def counts(self) -> dict[str, int]:
        values, counts = np.unique(self.code, return_counts=True)
        return {str(v): int(c) for v, c in zip(values.tolist(), counts.tolist())}


def resolve_frameworks(
    isin: np.ndarray,
    futures_dv01: np.ndarray,
    plain_swap_dv01: np.ndarray,
    available: Mapping[str, np.ndarray],
    *,
    global_override: str = "",
    per_bond_override: Mapping[str, str] | None = None,
    override_reasons: Mapping[str, str] | None = None,
) -> FrameworkResolution:
    """Choose a spread framework for every bond in the book.

    `available[code]` is a boolean array saying whether that framework's spread
    delta is a usable number for each bond.  A framework whose spread is missing
    is never selected automatically - picking it would produce a blank duration
    total and no attribution at all.

    An explicit override IS honoured even when its spread is missing.  That is
    deliberate: an override is a human decision, and silently overruling it
    would hide the fact that the data it needs is absent.  The row then reports
    a missing-framework status, which is the honest outcome.
    """
    isin = np.asarray(isin, dtype=object)
    n = isin.size

    fut = np.abs(np.nan_to_num(np.asarray(futures_dv01, dtype=np.float64), nan=0.0))
    swp = np.abs(np.nan_to_num(np.asarray(plain_swap_dv01, dtype=np.float64), nan=0.0))
    total = fut + swp

    with np.errstate(divide="ignore", invalid="ignore"):
        fut_w = np.where(total > 0, fut / total, 0.0)
        swp_w = np.where(total > 0, swp / total, 0.0)

    have = {code: np.asarray(available.get(code, np.zeros(n, bool)), dtype=bool)
            for code in _SPREAD_DELTA}

    # ---- automatic choice, vectorised over the whole book ------------------ #
    def first_available(order: Sequence[str]) -> np.ndarray:
        """First framework in `order` whose spread is usable, else 'REVIEW'.

        'REVIEW' rather than a hopeful default: if not one of the five spreads
        is quoted for a bond, there is nothing to measure its credit leg
        against, and inventing one produces a number with no source.
        """
        chosen = np.full(n, "REVIEW", dtype=object)
        assigned = np.zeros(n, dtype=bool)
        for code in order:
            take = have[code] & ~assigned
            chosen = np.where(take, code, chosen)
            assigned |= take
        return chosen

    futures_led = first_available(_FUTURES_PREFERENCE)
    swap_led = first_available(_SWAP_PREFERENCE)
    unhedged = first_available(_UNHEDGED_PREFERENCE)

    futures_dominant = (fut > 0) & (fut >= swp)
    swap_dominant = (swp > 0) & ~futures_dominant

    auto = np.where(futures_dominant, futures_led, np.where(swap_dominant, swap_led, unhedged))

    auto_reason = np.where(
        futures_dominant,
        "automatic - futures-dominated hedge, so the unhedged risk is spread over governments",
        np.where(
            swap_dominant,
            "automatic - swap-dominated hedge, so the unhedged risk is spread over swaps",
            "automatic - unhedged, so the best available spread quote is used",
        ),
    ).astype(object)

    # ---- overrides --------------------------------------------------------- #
    code = auto.copy()
    source = np.full(n, "automatic from hedge DV01 mix", dtype=object)
    reason = auto_reason.copy()

    gbl = str(global_override or "").strip().upper()
    if gbl:
        resolved = gbl if gbl in FRAMEWORKS else "REVIEW"
        code = np.full(n, resolved, dtype=object)
        source = np.full(n, "global override (config model.spread_framework_override)", dtype=object)
        reason = np.full(
            n,
            CHAINS[resolved].description
            if resolved in CHAINS
            else "unrecognised global override code",
            dtype=object,
        )

    if per_bond_override:
        lookup = {str(k).strip().upper(): str(v).strip().upper()
                  for k, v in per_bond_override.items()}
        notes = {str(k).strip().upper(): str(v)
                 for k, v in (override_reasons or {}).items()}
        for i, key in enumerate(isin.tolist()):
            picked = lookup.get(str(key).strip().upper())
            if picked is None:
                continue
            resolved = picked if picked in FRAMEWORKS else "REVIEW"
            code[i] = resolved
            source[i] = "per-bond override (overrides file)"
            note = notes.get(str(key).strip().upper(), "")
            base = (
                CHAINS[resolved].description
                if resolved in CHAINS
                else f"unrecognised override code {picked!r}"
            )
            reason[i] = f"{base} [{note}]" if note else base

    # Blend weights only mean anything under MIXED; keep the raw shares anyway,
    # because the report uses them to show how the BOOK is being measured.
    return FrameworkResolution(
        code=code,
        source=source,
        reason=np.array(
            [f"{r} [{s}]" for r, s in zip(reason.tolist(), source.tolist())], dtype=object
        ),
        futures_weight=fut_w,
        swap_weight=swp_w,
    )
