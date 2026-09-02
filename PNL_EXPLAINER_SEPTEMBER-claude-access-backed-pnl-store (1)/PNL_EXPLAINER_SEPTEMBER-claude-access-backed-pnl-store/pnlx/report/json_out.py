"""JSON metadata sidecar.

The CSV says what the numbers are.  This says what produced them: which dates,
which conventions, which inputs, what totalled, how much of the book was
covered, and what tied.  It is the file that makes a figure defensible six
months later, when the question is not "what was the residual" but "what was
the model doing when it produced that residual".

Everything is plain JSON - no numpy scalars, no NaN literals (which are not
valid JSON), no dataclasses.  A file that a strict parser rejects is not a
record of anything.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import math
from pathlib import Path
from typing import Any, Mapping

import numpy as np

from .. import __version__
from ..aggregate import PortfolioSummary
from ..config import AppConfig
from ..engine import AttributionResult
from ..snail import SnailTrail

__all__ = ["build_metadata", "write_metadata", "jsonable"]


def jsonable(value: Any) -> Any:
    """Coerce anything the model produces into valid JSON.

    NaN and infinity become null.  `json.dump` will happily emit the bare
    tokens `NaN` and `Infinity`, which are not JSON and which several strict
    parsers refuse - so a file that looked fine here would fail to load
    somewhere else, which is the worst place to discover it.
    """
    if value is None:
        return None
    if isinstance(value, (np.floating, float)):
        f = float(value)
        return None if (math.isnan(f) or math.isinf(f)) else f
    if isinstance(value, (np.integer, int)) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, (np.bool_, bool)):
        return bool(value)
    if isinstance(value, np.datetime64):
        return None if np.isnat(value) else str(value.astype("datetime64[D]"))
    if isinstance(value, (_dt.date, _dt.datetime)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, np.ndarray):
        return [jsonable(v) for v in value.tolist()]
    if isinstance(value, Mapping):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [jsonable(v) for v in value]
    return value


def _fingerprint(path: Path) -> dict[str, Any]:
    """Enough to tell whether an input file has changed since the last run."""
    if not path.is_file():
        return {"path": str(path), "present": False}
    data = path.read_bytes()
    return {
        "path": str(path),
        "present": True,
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "modified": _dt.datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
    }


def build_metadata(
    result: AttributionResult,
    summary: PortfolioSummary,
    trail: SnailTrail | None = None,
    outputs: Mapping[str, Path] | None = None,
) -> dict[str, Any]:
    """Assemble the metadata document for one run."""
    cfg: AppConfig = result.config
    paths = cfg.paths

    document: dict[str, Any] = {
        "schema": "pnlx.daily_pnl_metadata/1",
        "generator": {
            "package": "pnlx",
            "version": __version__,
            "run_at": _dt.datetime.now().isoformat(timespec="seconds"),
        },
        "run": {
            "as_of": cfg.run.as_of.isoformat(),
            "prior": cfg.run.prior.isoformat(),
            "days": cfg.run.days,
            "year_fraction": cfg.run.year_fraction,
            "base_ccy": cfg.run.base_ccy,
            "label": cfg.run.label,
            "book": cfg.run.book,
        },
        # The whole configuration, verbatim.  A figure cannot be reproduced from
        # the numbers alone: the conventions decide what they mean.
        "config": jsonable(cfg.to_dict()),
        "inputs": {
            "bonds": _fingerprint(paths.input(paths.bonds)),
            "futures": _fingerprint(paths.input(paths.futures)),
            "swaps": _fingerprint(paths.input(paths.swaps)),
            "curves": _fingerprint(paths.input(paths.curves)),
            "fx": _fingerprint(paths.input(paths.fx)),
            "spread_overrides": _fingerprint(paths.input(paths.spread_overrides)),
            "config_file": _fingerprint(paths.input(Path(str(cfg.source_path or "")).name))
            if cfg.source_path
            else {"present": False},
        },
        "counts": {
            "bonds": len(result.bonds),
            "futures": len(result.futures),
            "swaps": len(result.swaps),
        },
        "headline": jsonable(summary.headline),
        "bridge": jsonable(summary.bridge.to_records()),
        "asset_classes": jsonable(summary.asset_class_records()),
        "framework_mix": jsonable(summary.framework_mix),
        "by_portfolio": jsonable(summary.by_portfolio),
        "by_currency": jsonable(summary.by_currency),
        "quality": jsonable(summary.quality),
        "curves": jsonable(result.curves_described)
        if hasattr(result, "curves_described")
        else None,
        "tie_outs": {
            line.label: jsonable(line.amount)
            for line in summary.bridge.lines
            if line.kind == "check"
        },
        "conventions": _conventions(cfg),
    }

    if trail is not None:
        document["snail"] = {
            "intraday": jsonable(trail.step_records()),
            "history": jsonable(trail.point_records()),
            "reading": jsonable(trail.diagnostics),
        }

    if outputs:
        document["outputs"] = {k: str(v) for k, v in outputs.items()}

    return document


def _conventions(cfg: AppConfig) -> dict[str, str]:
    """The choices that change a number, stated in words rather than flags.

    A future reader should not have to know what `attribution_risk_date:
    prior` implies; the sentence is here so the file explains itself.
    """
    return {
        "risk_date": (
            "Attribution DV01 is struck at the OPENING of the period, so the risk "
            "used to explain a move is the risk before it. Hedging metrics use "
            "closing risk."
            if cfg.model.attribution_risk_date == "prior"
            else "Attribution DV01 is struck at the CLOSE, reproducing the legacy "
            "workbook. This leaves a standing residual proportional to the size "
            "of the move."
        ),
        "coupon_carry": (
            "Coupon carry is accrued(T0) - accrued(T-1) plus coupons actually paid, "
            "so the bridge closes across an ex-coupon date."
            if cfg.model.coupon_carry == "exact"
            else "Coupon carry is a smooth notional * coupon * year-fraction accrual. "
            "A coupon paid inside the period leaves a residual its size."
        ),
        "funding": (
            "Funding is inside the bridge; the practical PnL feed must be funded too."
            if cfg.model.include_funding_in_bridge
            else "Funding is an economic-carry memo OUTSIDE the bridge, because "
            "practical PnL is a mark-to-market figure with no financing leg."
        ),
        "fx": (
            "The second-order FX term (local move x FX move) is booked explicitly, "
            "so the FX decomposition is an exact identity."
            if cfg.model.fx_cross_term
            else "The second-order FX term is not booked and falls into the residual."
        ),
        "convexity": {
            "bump": f"Convexity by central-difference reprice at "
                    f"{cfg.model.convexity_bump_bp:g}bp.",
            "analytic": "Convexity at a vanishing bump - the analytic derivative.",
            "supplied": "Convexity taken from the position file as given.",
        }[cfg.model.convexity_source],
        "pull_to_par": (
            "Pull to par prices each remaining cash flow at the FORWARD rate the "
            "prior curve implies, so the roll-down of an upward-sloping curve is "
            "not booked as PnL the desk never earned."
            if cfg.model.pull_to_par.enabled
            else "Pull to par is disabled; carry is coupon only."
        ),
        "futures_dv01": (
            "FUT_PX_VAL_BP is treated as already being cash per contract per bp."
            if cfg.futures.dv01_includes_point_value
            else "FUT_PX_VAL_BP is treated as PRICE POINTS per bp and multiplied by "
            "the point value. Desk check: DV01/contracts should read 60-90 for a Bund."
        ),
        "swap_dv01": (
            "Swap risk is the figure the swap system publishes, falling back to the "
            "internal annuity model where it is absent."
            if cfg.swaps.dv01_source == "supplied"
            else "Swap risk is the internal annuity model throughout - an "
            "approximation that ignores the real payment schedule."
        ),
        "signs": (
            "Notional and contract counts are SIGNED and are the only sign driver. "
            "DV01 carries the position's sign, so every first-order leg is "
            "-DV01 * delta_bp and hedges net against bonds without special-casing."
        ),
        "units": (
            "Prices per 100 nominal; yields and curves in percent; spreads and all "
            "deltas in basis points; PnL, market value and DV01 in the base currency."
        ),
    }


def write_metadata(
    result: AttributionResult,
    summary: PortfolioSummary,
    trail: SnailTrail | None = None,
    outputs: Mapping[str, Path] | None = None,
) -> Path:
    """Write the metadata JSON and return its path."""
    cfg = result.config
    filename = cfg.output.json
    if cfg.output.date_stamp_filenames:
        stem, _, suffix = filename.rpartition(".")
        filename = f"{stem}_{cfg.run.as_of.isoformat()}.{suffix}"

    path = cfg.paths.output_dir / filename
    path.parent.mkdir(parents=True, exist_ok=True)

    document = build_metadata(result, summary, trail, outputs)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(document, fh, indent=2, allow_nan=False, sort_keys=False)
        fh.write("\n")
    return path
