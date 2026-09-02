"""Command line entry point.

    python -m pnlx                              run with config/pnl_explainer.yaml
    python -m pnlx --config other.yaml          a different configuration
    python -m pnlx --as-of 2025-09-30 --prior 2025-09-29
    python -m pnlx --no-excel                   CSV and JSON only
    python -m pnlx --check                      run and exit non-zero if a tie-out breaks

`--check` is the one intended for a scheduler.  A daily job that produces a
report nobody reads is worth little; one that fails loudly when the bridge
stops closing is worth a great deal.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import sys
import time
from dataclasses import replace
from pathlib import Path
from typing import Sequence

from . import __version__
from .aggregate import build_portfolio
from .config import AppConfig, ConfigError, load_config
from .engine import AttributionEngine
from .loaders import LoadError, load_inputs
from .report import append_history, write_csv_outputs, write_metadata, write_workbook
from .report.csv_out import history_row
from .snail import SnailTrail, build_history_snail, build_intraday_snail, load_history

__all__ = ["main", "run"]

#: How far a tie-out may be from zero before `--check` calls it a break.
#: Floating-point summation over a few hundred positions lands around 1e-9;
#: 1e-4 of a currency unit is a hundredth of a cent, which is not a real break
#: and is well clear of the noise.
TIE_OUT_TOLERANCE = 1e-4


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pnlx",
        description="Fixed-income PnL attribution: bonds, bond futures and interest-rate swaps.",
    )
    parser.add_argument("--config", type=Path, default=None,
                        help="path to the YAML configuration (default: config/pnl_explainer.yaml)")
    parser.add_argument("--as-of", type=str, default=None,
                        help="override the closing date (T0), as YYYY-MM-DD")
    parser.add_argument("--prior", type=str, default=None,
                        help="override the opening date (T-1), as YYYY-MM-DD")
    parser.add_argument("--input-dir", type=Path, default=None,
                        help="override the input directory")
    parser.add_argument("--output-dir", type=Path, default=None,
                        help="override the output directory")
    parser.add_argument("--framework", type=str, default=None,
                        help="force one spread framework for the whole run "
                             "(G/I/ASW/Z/OAS/OIS/SOFR/MIXED/REVIEW)")
    parser.add_argument("--risk-date", choices=("prior", "current"), default=None,
                        help="where the attribution DV01 is struck "
                             "(prior = the attribution convention; current = the legacy one)")
    parser.add_argument("--no-excel", action="store_true", help="skip the workbook")
    parser.add_argument("--no-history", action="store_true",
                        help="do not append this run to the history file")
    parser.add_argument("--check", action="store_true",
                        help="exit non-zero if a bridge tie-out does not close")
    parser.add_argument("--quiet", action="store_true", help="print nothing but errors")
    parser.add_argument("--version", action="version", version=f"pnlx {__version__}")
    return parser


def _apply_overrides(config: AppConfig, args: argparse.Namespace) -> AppConfig:
    """Command-line arguments win over the file, and are re-validated.

    Rebuilt through the dataclasses rather than mutated, so an override that
    produces an impossible configuration - a closing date before the opening
    one, say - fails here rather than three modules later.
    """
    run = config.run
    if args.as_of or args.prior:
        run = replace(
            run,
            as_of=_dt.date.fromisoformat(args.as_of) if args.as_of else run.as_of,
            prior=_dt.date.fromisoformat(args.prior) if args.prior else run.prior,
        )

    paths = config.paths
    if args.input_dir:
        paths = replace(paths, input_dir=args.input_dir)
    if args.output_dir:
        paths = replace(paths, output_dir=args.output_dir)

    model = config.model
    if args.framework is not None:
        model = replace(model, spread_framework_override=args.framework)
    if args.risk_date is not None:
        model = replace(model, attribution_risk_date=args.risk_date)

    return replace(config, run=run, paths=paths, model=model)


def run(config: AppConfig, *, write_excel: bool = True, write_history: bool = True):
    """Load, attribute, aggregate, write.  Returns everything it built."""
    inputs = load_inputs(config)
    result = AttributionEngine(config, inputs).run()
    summary = build_portfolio(result)

    steps = build_intraday_snail(result.bonds, summary)

    history_path = config.paths.output_dir / config.paths.history
    if write_history:
        history_path, history = append_history(result, summary)
        points, reading = build_history_snail(history)
    else:
        history = load_history(history_path)
        points, reading = build_history_snail(history, today=history_row(result, summary))

    trail = SnailTrail(steps=steps, points=points, diagnostics=reading)

    outputs = write_csv_outputs(result, summary)
    if write_history:
        outputs["history"] = history_path
    if write_excel:
        outputs["excel"] = write_workbook(result, summary, trail)
    outputs["metadata"] = write_metadata(result, summary, trail, outputs)

    return result, summary, trail, outputs


def _report(result, summary, trail, outputs, elapsed: float) -> None:
    cfg = result.config
    h = summary.headline
    q = summary.quality
    bar = "-" * 74

    print(bar)
    print(f"PnL Explainer {__version__}   {cfg.run.prior} -> {cfg.run.as_of}   "
          f"({cfg.run.days}d, base {cfg.run.base_ccy})")
    print(bar)
    print(f"  bonds {len(result.bonds):>4}   futures {len(result.futures):>4}   "
          f"swaps {len(result.swaps):>4}   in {elapsed * 1000:.0f} ms")
    print()
    for label, key in (
        ("Practical PnL", "practical_pnl"),
        ("Theoretical PnL", "theoretical_pnl"),
        ("Total explained", "total_explained"),
        ("Residual", "residual_pnl"),
    ):
        print(f"  {label:<22}{h[key]:>18,.2f}")
    print(f"  {'Residual %':<22}{h['residual_pct']:>17.2%}")
    print()
    print("  legs")
    for label, key in (
        ("duration", "duration_total"),
        ("convexity", "convexity"),
        ("carry", "carry_total"),
        ("fx", "pnl_fx"),
        ("hedge (model)", "model_hedge_pnl"),
        ("hedge basis", "hedge_basis_pnl"),
    ):
        print(f"    {label:<20}{h[key]:>18,.2f}")
    print()
    print(f"  coverage {q.get('coverage_explained', 0):.1%} attributed, "
          f"{q.get('rows_ok', 0)}/{q.get('bond_count', 0)} rows clean")
    print(f"  snail: {trail.diagnostics.get('verdict', '-')}")
    print()
    print("  tie-outs")
    for line in summary.bridge.lines:
        if line.kind != "check":
            continue
        ok = abs(line.amount) < TIE_OUT_TOLERANCE
        print(f"    {'OK ' if ok else 'BREAK'}  {line.label:<58}{line.amount:>14.8f}")
    print()
    print("  written")
    for name, path in outputs.items():
        print(f"    {name:<12}{path}")
    print(bar)


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        config = _apply_overrides(load_config(args.config), args)
    except (ConfigError, ValueError) as exc:
        print(f"configuration error: {exc}", file=sys.stderr)
        return 2

    started = time.perf_counter()
    try:
        result, summary, trail, outputs = run(
            config, write_excel=not args.no_excel, write_history=not args.no_history
        )
    except LoadError as exc:
        print(f"input error: {exc}", file=sys.stderr)
        return 2
    elapsed = time.perf_counter() - started

    if not args.quiet:
        _report(result, summary, trail, outputs, elapsed)

    if args.check:
        breaks = [
            line for line in summary.bridge.lines
            if line.kind == "check" and abs(line.amount) >= TIE_OUT_TOLERANCE
        ]
        if breaks:
            for line in breaks:
                print(f"TIE-OUT BREAK: {line.label} = {line.amount:.8f}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
