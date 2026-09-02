#!/usr/bin/env python3
"""Execute tests/test_pull_to_par.bas against the real src/modPNL.bas source."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import modules  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from vba_run import BasicRunner                          # noqa: E402

PURE_FROM_MODPNL = [
    "BondPullToParCore",
    "BondPullPvOnCurve",
    "BondPullModelDirtyPrice",
    "BondPullSchedule",
    "BondPullYears",
    "BondPullInterp",
    "BondPullCurveType",
]


def main() -> int:
    r = BasicRunner()
    # modImpRepo is pure arithmetic end to end - take it whole.
    r.add_module(modules.path_of("modImpRepo"))
    r.add_procedures(modules.path_of("modPNL"), PURE_FROM_MODPNL)
    r.add_source((ROOT / "tests" / "test_pull_to_par.bas").read_text())

    out = r.run("RunTests")
    print(out)
    return 0 if "failed=0" in out else 1


if __name__ == "__main__":
    raise SystemExit(main())
