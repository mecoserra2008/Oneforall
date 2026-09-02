#!/usr/bin/env python3
"""Find the four VBA modules, whatever they are currently called.

The modules are exported from the VBE and land in the repository root under
whatever name the export dialog was given that day - modPNL_25th_Aug.txt became
modPNL_26th_Aug.txt became Dashboard.txt, and each rename silently took every
tool offline until somebody hand-copied the files into src/.  A check that only
runs when you remember to stage it is not a check.

So nothing here matches on a filename.  Each module is identified by a procedure
only it defines, which survives being renamed, re-exported, or moved between
src/ and the root.  If a module genuinely is missing, that is reported as a
missing MODULE rather than as a mysteriously empty result.
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# logical module -> a procedure that only that module defines
SIGNATURES = {
    "modPNL": "SetupWorkbookFinalLayout",
    "modDashboard": "BuildDashboard_Step6",
    "modEconFormulas": "HedgeEfficiencyFml",
    "modImpRepo": "ImpliedRepoBloomberg",
    "modStore": "Store_SaveCurrentRun",
}


def _candidates() -> list[Path]:
    """src/ first, so a staged working copy wins over the exported originals."""
    return sorted((ROOT / "src").glob("*.bas")) + sorted(ROOT.glob("*.txt"))


def discover() -> dict[str, Path]:
    """Map each logical module name to the file that defines it."""
    found: dict[str, Path] = {}
    for path in _candidates():
        try:
            head = path.read_text(encoding="latin-1")
        except OSError:
            continue
        for name, proc in SIGNATURES.items():
            if name in found:
                continue
            if re.search(r"^(?:Public |Private )?(?:Sub|Function) "
                         + re.escape(proc) + r"\b", head, re.M):
                found[name] = path
    return found


def path_of(name: str) -> Path:
    """The file defining one module, or a clear error naming what is missing."""
    found = discover()
    if name not in found:
        raise SystemExit(
            f"cannot find {name}: no file under {ROOT} defines "
            f"{SIGNATURES.get(name, '?')}. Looked in src/*.bas and *.txt.")
    return found[name]


def all_paths() -> list[Path]:
    """Every module file, deduplicated, in a stable order."""
    found = discover()
    seen: list[Path] = []
    for name in SIGNATURES:
        p = found.get(name)
        if p is not None and p not in seen:
            seen.append(p)
    return seen


if __name__ == "__main__":
    for name, path in sorted(discover().items()):
        print(f"{name:18} {path.relative_to(ROOT)}")
