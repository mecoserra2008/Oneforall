#!/usr/bin/env python3
"""Check config/pipeline.yaml against the code it describes.

The config carries the WIRING - paths, sheets, columns, which formula feeds which
field, and the wave order. The maths stays in Economic_formula_library.txt. That
split is only worth anything if it is enforced, because the failure it prevents is
silent: naming a builder whose economics differ from the live cell looks exactly
like rewiring and changes the model.

Six builders are that hazard. The library says so itself, in its own header:

    CouponCarryFml   re-applies PositionSide to an already-signed Notional,
                     which silently flips every short
    FxPnLFml         blanks on missing inputs, where the live cell returns 0 -
                     so one missing FX would blank the whole explained total

This refuses a config that names any of them.

    python3 tools/check_pipeline_config.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    raise SystemExit("PyYAML is required: pip install pyyaml")

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config" / "pipeline.yaml"
LIBRARY = ROOT / "Economic_formula_library.txt"

sys.path.insert(0, str(Path(__file__).resolve().parent))
import field_graph  # noqa: E402


def builders() -> set[str]:
    text = LIBRARY.read_text(encoding="latin-1")
    return set(re.findall(r"^Public Function (\w+)", text, re.M))


def main() -> int:
    cfg = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))
    have = builders()
    problems: list[str] = []

    cat = cfg.get("catalogue", {})
    aliases: dict[str, str] = cat.get("aliases", {}) or {}
    forbidden: dict[str, str] = cat.get("do_not_wire", {}) or {}
    live: list[str] = cat.get("live", []) or []

    # 1) every name the catalogue mentions must exist in the library
    for group, names in (("live", live), ("do_not_wire", list(forbidden)),
                         ("aliases", list(aliases.values()))):
        for name in names:
            if name not in have:
                problems.append(f"catalogue.{group}: {name} is not in "
                                f"Economic_formula_library.txt")

    # 2) the library's own header must still agree with the do_not_wire list -
    #    if somebody clears a warning there, this file must not keep enforcing it
    header = LIBRARY.read_text(encoding="latin-1")[:6000]
    for name in forbidden:
        if name not in header:
            problems.append(f"catalogue.do_not_wire lists {name}, but the "
                            f"library header no longer mentions it - re-read the "
                            f"header before trusting this entry")

    # 3) no field may name a forbidden builder, directly or through an alias
    fields = cfg.get("fields", []) or []
    for f in fields:
        fml = f.get("formula")
        if not fml:
            problems.append(f"field {f.get('name')}: no formula")
            continue
        if fml == "inline_live":
            if not f.get("expression"):
                problems.append(f"field {f['name']}: inline_live needs an expression")
            continue
        if fml in ("copy",):
            continue
        target = aliases.get(fml, fml)
        if target in forbidden:
            problems.append(
                f"field {f['name']}: formula {fml} -> {target}, which is on the "
                f"do-not-wire list ({forbidden[target]})")
        elif target not in have:
            problems.append(f"field {f['name']}: formula {fml} -> {target}, "
                            f"which is not a builder in the library")

    # 4) wave numbers must match the graph the formula record proves
    graph = field_graph.build()
    by_name = {row[1]: graph["wave"][row[0]] for row in graph["rows"]}
    for f in fields:
        name, wave = f.get("name"), f.get("wave")
        if name in by_name and wave is not None and by_name[name] != wave:
            problems.append(f"field {name}: config says wave {wave}, "
                            f"tools/field_graph.py proves wave {by_name[name]}")

    # 5) the two coverage books must not share a contract-count column - they
    #    genuinely differ, and a copy-paste between them is the likely edit
    rtj = cfg["sources"]["hedges_rtj"]["variants"]["future"]["columns"]["contracts"]
    rt = cfg["sources"]["hedges_rt"]["columns"]["contracts"]["column"]
    if rtj == rt:
        problems.append(f"hedges_rtj and hedges_rt both read contracts from "
                        f"column {rtj}; RTJ is O and RT is P")

    print(f"{len(fields)} fields, {len(have)} builders, "
          f"{len(forbidden)} on the do-not-wire list")
    if problems:
        print(f"\n{len(problems)} problem(s):")
        for p in problems:
            print(f"  {p}")
        return 1
    print("\nconfig agrees with the library and the dependency graph")
    return 0


if __name__ == "__main__":
    sys.exit(main())
