#!/usr/bin/env python3
"""Execute tests/test_bbg_fallback.bas against the real modPNL source.

The order of the security fallback chain decides which Bloomberg id a bond's
fields are actually asked of, and its LENGTH is capped by Excel's 64-level
nesting limit.  Neither fails loudly: a wrong order silently prices a bond off
the wrong quote source, and a formula over the limit is refused by Excel so the
write fails and the column stays empty.

The builders are pure string assembly, so they run for real here.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import modules                                            # noqa: E402
from vba_run import BasicRunner                           # noqa: E402

PURE_FROM_MODPNL = [
    "BondSecurityFallbackCols",
    "BondFallbackColCount",
    "BBGTryBDPFieldExprR1C1_Short",
    "BBGFirstBDPMultiFieldFormulaR1C1",
    "GuardedBDPExprR1C1",
    "WrapIfPresent",
    "RC",
    "ColIdx",
]


def main() -> int:
    r = BasicRunner()
    r.add_module(modules.path_of("modEconFormulas"))      # XlText, XlBlank
    r.add_procedures(modules.path_of("modPNL"), PURE_FROM_MODPNL)
    r.add_source((ROOT / "tests" / "test_bbg_fallback.bas").read_text())

    out = r.run("RunTests")
    print(out)
    return 0 if "failed=0" in out else 1


if __name__ == "__main__":
    raise SystemExit(main())
