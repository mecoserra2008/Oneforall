"""How fast is it, and does it stay linear?

The migration's performance claim is that the attribution is array arithmetic
over a whole book rather than a loop over positions, so cost should grow
roughly linearly with the book and stay far below anything a desk would notice.
This measures it rather than asserting it.

    python tools/benchmark.py
    python tools/benchmark.py --sizes 500 2000 10000 --repeat 5
"""

from __future__ import annotations

import argparse
import statistics
import sys
import time
from dataclasses import replace
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pnlx.aggregate import build_portfolio  # noqa: E402
from pnlx.config import load_config  # noqa: E402
from pnlx.engine import AttributionEngine  # noqa: E402
from pnlx.instruments import BondBook, BondPosition, FutureBook, SwapBook  # noqa: E402
from pnlx.loaders import InputBundle, load_inputs  # noqa: E402


def grow(bundle: InputBundle, target: int) -> InputBundle:
    """Repeat the sample book until it holds `target` bonds.

    Each copy gets a distinct ISIN, so hedge linkage stays one-to-one and the
    scatter-add is doing the same work per position it would on a real book of
    that size - a book of duplicates would collapse into one bucket and
    flatter the result.
    """
    base = list(bundle.bonds.positions)
    positions: list[BondPosition] = []
    copy_index = 0
    while len(positions) < target:
        suffix = f"{copy_index:04d}"
        for position in base:
            if len(positions) >= target:
                break
            positions.append(replace(position, isin=f"{position.isin[:8]}{suffix}"))
        copy_index += 1
    return InputBundle(
        bonds=BondBook.from_positions(positions),
        futures=bundle.futures,
        swaps=bundle.swaps,
        curves=bundle.curves,
        fx=bundle.fx,
        override_codes=bundle.override_codes,
        override_notes=bundle.override_notes,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sizes", type=int, nargs="+", default=[35, 250, 1000, 5000, 20000])
    parser.add_argument("--repeat", type=int, default=3)
    args = parser.parse_args()

    config = load_config()
    base = load_inputs(config)

    print(f"{'bonds':>8}  {'engine ms':>10}  {'aggregate ms':>13}  {'total ms':>10}  {'us/bond':>9}")
    print("-" * 60)

    for size in args.sizes:
        bundle = base if size == len(base.bonds) else grow(base, size)

        engine_times, aggregate_times = [], []
        for _ in range(args.repeat):
            start = time.perf_counter()
            result = AttributionEngine(config, bundle).run()
            middle = time.perf_counter()
            build_portfolio(result)
            end = time.perf_counter()
            engine_times.append((middle - start) * 1000)
            aggregate_times.append((end - middle) * 1000)

        engine = statistics.median(engine_times)
        aggregate = statistics.median(aggregate_times)
        total = engine + aggregate
        print(
            f"{len(bundle.bonds):>8}  {engine:>10.1f}  {aggregate:>13.1f}  "
            f"{total:>10.1f}  {total * 1000 / len(bundle.bonds):>9.1f}"
        )

    print()
    print("Per-bond cost should FALL as the book grows: the fixed cost of setting up")
    print("each array pass is amortised over more positions. A per-bond cost that")
    print("stays flat would mean something is looping per position.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
