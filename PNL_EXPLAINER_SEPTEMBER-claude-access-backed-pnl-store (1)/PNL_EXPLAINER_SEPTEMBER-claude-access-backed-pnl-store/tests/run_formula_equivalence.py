#!/usr/bin/env python3
"""Prove the modEconFormulas builders emit exactly the strings WritePNLRow
used to build inline."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import modules  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from vba_run import BasicRunner                          # noqa: E402


def main() -> int:
    r = BasicRunner()
    # The whole formula library: every function here is a pure string builder.
    r.add_module(modules.path_of("modEconFormulas"))
    # The real column constants, so the cases pin geometry as well as shape.
    r.add_constants(modules.path_of("modPNL"), ("PCOL_", "BCOL_", "CFG_"))
    r.add_source((ROOT / "tests" / "test_formula_equivalence.bas").read_text())

    out = r.run("RunTests")
    print(out)
    return 0 if "failed=0" in out else 1


if __name__ == "__main__":
    raise SystemExit(main())
