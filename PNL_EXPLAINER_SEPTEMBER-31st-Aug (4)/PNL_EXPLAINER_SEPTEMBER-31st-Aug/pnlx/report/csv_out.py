"""CSV outputs: the daily PnL explained, and the history the snail trail reads.

Four files per run:

    daily_pnl_explained.csv   one row per bond, every attribution column
    daily_pnl_bonds.csv       the same, kept under its own name so a caller can
                              point at "the bonds file" without knowing which of
                              the two is the headline
    daily_pnl_futures.csv     one row per futures position
    daily_pnl_swaps.csv       one row per swap

and one file that grows:

    history/daily_pnl_history.csv   one row per RUN, appended

The history file is deliberately narrow - headline figures only.  It is read
back on every run to build the snail trail, so it has to stay cheap to parse
and stable in shape.  Re-running the same date REPLACES that date's row rather
than appending a second one, because a trail with two points for one day is
wrong in a way that is hard to see and easy to believe.
"""

from __future__ import annotations

import csv
import datetime as _dt
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from ..aggregate import PortfolioSummary
from ..columns import ColumnFrame
from ..config import AppConfig
from ..engine import AttributionResult

__all__ = ["write_csv_outputs", "append_history", "write_frame", "HISTORY_FIELDS"]


#: The narrow, stable schema of the history file.  Adding a column here is
#: backward compatible (older rows read as blank); removing one is not.
HISTORY_FIELDS: tuple[str, ...] = (
    "as_of",
    "prior",
    "label",
    "book",
    "base_ccy",
    "days",
    "bond_count",
    "practical_pnl",
    "theoretical_pnl",
    "total_explained",
    "residual_pnl",
    "not_attributed",
    "residual_pct",
    "duration_total",
    "convexity",
    "carry_total",
    "carry_coupon",
    "carry_roll_to_par",
    "pnl_fx",
    "model_hedge_pnl",
    "actual_hedge_pnl",
    "hedge_basis_pnl",
    "funding_carry_memo",
    "unlinked_hedge_pnl",
    "bond_dv01",
    "hedge_dv01",
    "residual_dv01",
    "mv_current",
    "coverage_explained",
    "rows_ok",
    "attribution_risk_date",
    "coupon_carry",
    "spread_framework_override",
    "run_at",
)


def _fmt(value: Any, float_format: str) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        return float_format % value
    if isinstance(value, (_dt.date, _dt.datetime)):
        return value.isoformat()[:10]
    return str(value)


def write_frame(frame: ColumnFrame, path: Path, float_format: str = "%.6f") -> Path:
    """Write one column frame, in schema order, with schema header names."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(frame.names)
        for row in frame.rows():
            writer.writerow([_fmt(v, float_format) for v in row])
    return path


def write_csv_outputs(
    result: AttributionResult, summary: PortfolioSummary
) -> dict[str, Path]:
    """Write every per-position CSV for a run."""
    cfg = result.config
    out_dir = cfg.paths.output_dir
    fmt = cfg.output.float_format
    stamp = cfg.run.as_of.isoformat() if cfg.output.date_stamp_filenames else ""

    def name(filename: str) -> Path:
        if not stamp:
            return out_dir / filename
        stem, _, suffix = filename.rpartition(".")
        return out_dir / f"{stem}_{stamp}.{suffix}"

    written: dict[str, Path] = {}
    written["explained"] = write_frame(result.bonds, name(cfg.output.csv), fmt)
    written["bonds"] = write_frame(result.bonds, name(cfg.output.bonds_csv), fmt)
    written["futures"] = write_frame(result.futures, name(cfg.output.futures_csv), fmt)
    written["swaps"] = write_frame(result.swaps, name(cfg.output.swaps_csv), fmt)
    return written


def history_row(result: AttributionResult, summary: PortfolioSummary) -> dict[str, Any]:
    """The one row this run contributes to the history file."""
    cfg = result.config
    h = summary.headline
    q = summary.quality
    return {
        "as_of": cfg.run.as_of.isoformat(),
        "prior": cfg.run.prior.isoformat(),
        "label": cfg.run.label,
        "book": cfg.run.book,
        "base_ccy": cfg.run.base_ccy,
        "days": cfg.run.days,
        "bond_count": q.get("bond_count", 0),
        "practical_pnl": h["practical_pnl"],
        "theoretical_pnl": h["theoretical_pnl"],
        "total_explained": h["total_explained"],
        "residual_pnl": h["residual_pnl"],
        "not_attributed": h["not_attributed"],
        "residual_pct": h["residual_pct"],
        "duration_total": h["duration_total"],
        "convexity": h["convexity"],
        "carry_total": h["carry_total"],
        "carry_coupon": h["carry_coupon"],
        "carry_roll_to_par": h["carry_roll_to_par"],
        "pnl_fx": h["pnl_fx"],
        "model_hedge_pnl": h["model_hedge_pnl"],
        "actual_hedge_pnl": h["actual_hedge_pnl"],
        "hedge_basis_pnl": h["hedge_basis_pnl"],
        "funding_carry_memo": h["funding_carry_memo"],
        "unlinked_hedge_pnl": h["unlinked_hedge_pnl"],
        "bond_dv01": h["bond_dv01"],
        "hedge_dv01": h["hedge_dv01"],
        "residual_dv01": h["residual_dv01"],
        "mv_current": h["mv_current"],
        "coverage_explained": q.get("coverage_explained", float("nan")),
        "rows_ok": q.get("rows_ok", 0),
        "attribution_risk_date": cfg.model.attribution_risk_date,
        "coupon_carry": cfg.model.coupon_carry,
        "spread_framework_override": cfg.model.spread_framework_override,
        "run_at": _dt.datetime.now().isoformat(timespec="seconds"),
    }


def append_history(
    result: AttributionResult, summary: PortfolioSummary
) -> tuple[Path, list[dict[str, str]]]:
    """Append this run to the history file and return the whole history.

    Rewrites the file rather than opening it in append mode, for two reasons:
    re-running a date must REPLACE its row (a trail with two points for one day
    is a silent lie), and a schema that has gained a column must not leave the
    older rows misaligned against the new header.
    """
    cfg = result.config
    path = cfg.paths.output_dir / cfg.paths.history
    path.parent.mkdir(parents=True, exist_ok=True)

    existing: list[dict[str, str]] = []
    if path.is_file():
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            existing = [dict(r) for r in csv.DictReader(fh)]

    row = history_row(result, summary)
    stamp = str(row["as_of"])
    existing = [r for r in existing if str(r.get("as_of", "")) != stamp]

    normalised = [{k: _fmt(r.get(k), "%.6f") for k in HISTORY_FIELDS} for r in existing]
    normalised.append({k: _fmt(row.get(k), "%.6f") for k in HISTORY_FIELDS})
    normalised.sort(key=lambda r: r["as_of"])

    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(HISTORY_FIELDS))
        writer.writeheader()
        writer.writerows(normalised)

    return path, normalised
