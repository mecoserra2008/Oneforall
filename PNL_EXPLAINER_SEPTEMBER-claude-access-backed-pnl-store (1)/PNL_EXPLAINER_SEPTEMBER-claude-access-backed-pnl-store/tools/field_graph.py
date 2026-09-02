#!/usr/bin/env python3
"""The PNL_Attribution dependency graph: which field feeds which, and in what order.

WHY THIS EXISTS

A field is not an "input" or an "output".  `Bond_DV01_Opening` is produced by one
step and consumed by twelve others; `Spread_Framework_Auto` is produced by one and
consumed by six.  Labelling either as an input or an output is wrong half the time,
and a schema built on that labelling encodes the mistake permanently.

So the contract carries the EDGES instead, and the wave number falls out of them:

    ProducedInWave   the topological rank - nothing in wave N reads wave N or later
    ConsumedBy       the fields that read this one

"Is this an input?" is then a question you ask per consumer, against the edge list,
rather than a property you guess from a name.

WHERE THE EDGES COME FROM

docs/FORMULAS.md, which tools/dump_formulas.py generates by evaluating the string
concatenation in WritePNLRow the way VBA would.  So the graph is derived from the
code that writes the formulas, not from a hand-transcribed list, and it cannot
drift from them.  Regenerate that file first if WritePNLRow changes.

    python3 tools/dump_formulas.py > docs/FORMULAS.md
    python3 tools/field_graph.py

Row numbers in that file are a substitution pattern, not data: PNL_Attribution is
written at row 5 and Bonds at row 6, so `L5` is this row's Bond_DV01_Current and
`'Bonds'!AF6` is this bond's DV01_EUR.  Self-references are therefore the `...5`
ones, and they are what the edges are built from.
"""
from __future__ import annotations

import collections
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FORMULAS = ROOT / "docs" / "FORMULAS.md"

#: PNL_Attribution's own row in the generated record.  Anything ending in this
#: digit and not qualified by a sheet name is a same-row reference.
SELF_ROW = 5

#: A reference into another sheet: 'Bonds'!AF6, Config!B4, Swaps!$L$5.  Stripped
#: out before self-references are hunted, so Bonds!AF6 is never read as AF6.
EXTERNAL = re.compile(r"'?([A-Za-z_][A-Za-z_0-9 ]*)'?!\$?([A-Z]{1,2})\$?\d+")

#: A same-row reference.  The negative lookbehind keeps it from firing inside a
#: range (`$L$5:$L$999`) or immediately after a sheet name.
SELF = re.compile(r"(?<![A-Za-z0-9_!$:])\$?([A-Z]{1,2})\$?%d(?![0-9])" % SELF_ROW)


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def read_formulas(path: Path = FORMULAS) -> list[tuple[str, str, str, str]]:
    """(letter, field, key, formula) per row of the generated formula record."""
    rows = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|"):
            continue
        cells = [c.strip().strip("`") for c in line.strip().strip("|").split("|")]
        if len(cells) < 4 or cells[0] == "Col" or set(cells[0]) <= set("-"):
            continue
        rows.append(tuple(cells[:4]))
    return rows


def build(rows=None) -> dict:
    """The graph: deps, consumers, waves, and the external reads per field."""
    rows = rows if rows is not None else read_formulas()
    by_letter = {r[0]: r for r in rows}

    deps: dict[str, set[str]] = {}
    external: dict[str, set[str]] = collections.defaultdict(set)

    for letter, field, key, formula in rows:
        stripped = EXTERNAL.sub(" ", formula)
        d = {m.group(1) for m in SELF.finditer(stripped)} & set(by_letter)
        d.discard(letter)
        deps[letter] = d
        for m in EXTERNAL.finditer(formula):
            external[letter].add(f"{m.group(1)}!{m.group(2)}")

    consumers: dict[str, set[str]] = collections.defaultdict(set)
    for letter, d in deps.items():
        for src in d:
            consumers[src].add(letter)

    # Topological rank.  A cycle would mean a formula chain Excel could not
    # evaluate either, so it is a hard error rather than something to work around.
    wave: dict[str, int] = {}
    pending = dict(deps)
    n = 0
    while pending:
        ready = [c for c, d in pending.items() if all(x in wave for x in d)]
        if not ready:
            raise SystemExit(
                "cycle in the PNL_Attribution formula graph among: "
                + ", ".join(sorted(pending)))
        for c in ready:
            wave[c] = n
            pending.pop(c)
        n += 1

    return {
        "rows": rows,
        "by_letter": by_letter,
        "deps": deps,
        "consumers": {k: sorted(v, key=col_num) for k, v in consumers.items()},
        "wave": wave,
        "external": {k: sorted(v) for k, v in external.items()},
        "wave_count": n,
    }


def main() -> int:
    g = build()
    by_wave = collections.defaultdict(list)
    for letter, w in g["wave"].items():
        by_wave[w].append(letter)

    print(f"PNL_Attribution: {len(g['rows'])} fields, {g['wave_count']} waves, no cycles\n")
    for w in sorted(by_wave):
        cols = sorted(by_wave[w], key=col_num)
        print(f"WAVE {w}  ({len(cols)} fields)")
        for c in cols:
            field = g["by_letter"][c][1]
            needs = sorted(g["deps"][c], key=col_num)
            ext = g["external"].get(c, [])
            tail = []
            if needs:
                tail.append("needs " + ",".join(needs))
            if ext:
                tail.append("reads " + ",".join(ext[:4]) + ("..." if len(ext) > 4 else ""))
            print(f"   {c:<3} {field:<32} {'  |  '.join(tail)}")
        print()

    print("Fields that are an output here and an input elsewhere, most first:")
    for c, cons in sorted(g["consumers"].items(), key=lambda t: -len(t[1]))[:12]:
        print(f"   {c:<3} {g['by_letter'][c][1]:<32} feeds {len(cons):>2}: "
              f"{','.join(cons)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
