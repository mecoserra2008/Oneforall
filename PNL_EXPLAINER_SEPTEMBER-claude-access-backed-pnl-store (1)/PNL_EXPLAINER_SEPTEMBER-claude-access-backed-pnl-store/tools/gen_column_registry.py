#!/usr/bin/env python3
"""Generate the Access column registry from the constants that exist today.

Meta_Column is the table that replaces the PCOL_/BCOL_/FCOL_/WCOL_/CVCOL_ letter
constants: it says which column sits where, and moving one becomes an UPDATE
instead of an edit to VBA plus three checkers plus a reconciliation.

Seeding it by hand would mean typing 330 rows and getting some of them wrong, so
it is generated from the workbook as it stands.  The registry therefore starts
out agreeing with the sheets exactly, which is the only starting point from
which the cutover can be verified.

    python3 tools/gen_column_registry.py            # report what it found
    python3 tools/gen_column_registry.py --write    # write the .sql files

Writes:
    access/seed_meta_column.sql   INSERT rows for Meta_Column, all five sheets
    access/alter_fact_pnl.sql     ALTER TABLE for Fact_PnlAttribution's measures

See docs/ACCESS_ARCHITECTURE.md.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402
import classify_columns  # noqa: E402
import field_graph  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent

# constant prefix -> the sheet it lays out
SHEETS = {
    "BCOL_": "Bonds",
    "PCOL_": "PNL_Attribution",
    "FCOL_": "Futures",
    "WCOL_": "Swaps",
    "CVCOL_": "OIS_Curves",
}

# Ordinals are spaced so a column can be inserted between two others without
# renumbering anything.  Ten is enough room and stays readable.
ORDINAL_STEP = 10

# Fields that are text or whole numbers rather than measures.  Everything not
# named here is a DOUBLE, which is right for every PnL and risk figure.
#
# CHECKED AGAINST THE FORMULAS, not guessed from the names, because two of them
# read like statuses and are not:
#
#   Duration_Identity_Check  is a RESIDUAL - PnL_Duration_Total - (-DV01 * dy) -
#                            and blank when its inputs are not numeric
#   SpreadPnL_Used           is a SWITCH over the framework that returns one of
#                            the spread PnL amounts
#
# Both are DOUBLE.  Getting either wrong would store a number as text and
# silently break every SUM over it.
TEXT_FIELDS = {
    "ISIN", "Name", "CCY", "Portfolio", "AcctgCat",
    "Attribution_Status", "Spread_Framework_Auto", "Spread_Framework_Reason",
    "Row_Exclusion_Reason",
}
LONG_FIELDS = {
    "Days", "Row_Valid",
    "Futures_Match_Count", "Swap_All_Match_Count", "PlainSwap_Match_Count",
    "SyntheticSwap_Match_Count",
}

# Constants whose "' col N = ..." comment is not a usable field name.
#
# OIS_Curves lays the three currencies out side by side, and all three tenor
# columns are commented "Years".  On a sheet that is unambiguous - they are in
# different places - but a registry addresses by NAME, so three columns cannot
# all be called Years.  This is exactly the kind of thing the move surfaces:
# position was carrying meaning that the name did not.
FIELD_NAME_OVERRIDE = {
    "CVCOL_YEARS":    "EUR_Years",
    "CVCOL_YEARS_S":  "USD_Years",
    "CVCOL_YEARS_AJ": "GBP_Years",
}


def sql_str(s: str) -> str:
    """A string literal for Access SQL: single quotes, doubled inside."""
    return "'" + s.replace("'", "''") + "'"


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def pnl_layout(text: str) -> list[tuple[str, str, str]]:
    """(constant, contract key, display label) from PnlLayout, in sheet order."""
    src = re.search(r"(?:Public|Private) Function PnlLayout.*?\nEnd Function",
                    text, re.S)
    if not src:
        return []
    return re.findall(
        r'AddPnlCol spec,\s*(\w+),\s*"([^"]+)",\s*"((?:[^"]|"")*)"', src.group(0))


def letter_constants(text: str) -> dict[str, tuple[str, int]]:
    """constant name -> (letter, column number)."""
    out = {}
    for m in re.finditer(
            r'^Private Const (\w+)\s+As String\s*=\s*"([A-Z]{1,3})"', text, re.M):
        out[m.group(1)] = (m.group(2), col_num(m.group(2)))
    return out


def field_name_of(const: str, prefix: str) -> str:
    """A constant's field name, when the layout table does not give one.

    The constants carry it in the trailing comment - `' col 12 = DaysLeft` -
    which is the desk's own name for the column and the right thing to put in
    the registry.  Falling back to the constant's own suffix keeps a column
    that has lost its comment from being dropped silently.
    """
    return const[len(prefix):].title().replace("__", "_")


def collect(pnl_text: str) -> dict[str, list[dict]]:
    """Every column of every sheet, keyed by sheet, in sheet order."""
    consts = letter_constants(pnl_text)
    named = {}          # constant -> (field, label) where the layout knows better
    for const, key, label in pnl_layout(pnl_text):
        named[const] = (key, label.replace('""', '"'))

    # the trailing "' col N = FieldName" comment on each constant
    commented = dict(re.findall(
        r"^Private Const (\w+)\s+As String\s*=\s*\"[A-Z]{1,3}\"\s*'\s*col \d+ = (\S+)",
        pnl_text, re.M))

    sheets: dict[str, list[dict]] = {s: [] for s in SHEETS.values()}
    for const, (letter, n) in consts.items():
        for prefix, sheet in SHEETS.items():
            if not const.startswith(prefix):
                continue
            if const in FIELD_NAME_OVERRIDE:
                field = label = FIELD_NAME_OVERRIDE[const]
            elif const in named:
                field, label = named[const]
            else:
                field = commented.get(const) or field_name_of(const, prefix)
                label = field
            sheets[sheet].append(
                {"const": const, "letter": letter, "col": n,
                 "field": field, "label": label})
            break

    for rows in sheets.values():
        rows.sort(key=lambda r: r["col"])
    return sheets


def access_type(field: str) -> str:
    if field in TEXT_FIELDS:
        return "TEXT(128)"
    if field in LONG_FIELDS:
        return "LONG"
    return "DOUBLE"



def writer_order(pnl_text: str) -> dict[str, int]:
    """Procedure name -> its position in the order its *FormulasSafe wrapper calls it.

    For every sheet except PNL_Attribution the wave IS the writer order: the
    wrappers call their writers in a sequence the comments call mandatory, because
    a later writer consumes what an earlier one produced (the ticker must exist
    before any BDP that keys off it).  Reading the order out of the wrapper keeps
    this from being a second, hand-maintained copy of that sequence.
    """
    order: dict[str, int] = {}
    for wrapper in re.finditer(
            r"^(?:Public |Private )?(?:Sub|Function) (Write\w*FormulasSafe)\b(.*?)^End (?:Sub|Function)",
            pnl_text, re.M | re.S):
        seen: list[str] = []
        for call in re.finditer(r"\b(Write\w+)\s*(?:ws|\()", wrapper.group(2)):
            name = call.group(1)
            if name.endswith("Safe") or name in seen:
                continue
            seen.append(name)
        for i, name in enumerate(seen, start=1):
            order.setdefault(name, i)
    return order



#: Procedures that put values on a sheet BEFORE any formula writer runs - the
#: OPICS query, the hedge load, and the one-off layout setup.  Everything they
#: write is wave 0 by definition: it is what the writers read.
INGEST_PROCS = {
    "OPICS query table",
    "LabelBondsCommentRows",
    "SetupCurvesLayout",
    "SetupSwapsFinalHeaders",
    "AppendSwapMapRowsToSwaps",
    "AppendCoverageFuturesRowsToFutures",
    "EnrichCoverageFuturesFromOPICS",
    "LinkSwapToBond",
    "LinkFutureToBond",
    "LoadHedges_Step2",
}

#: Written only by a diagnostic path.  The normal writer for the same column is
#: named here so the column still gets its real wave rather than a null.
DEBUG_WRITER_REAL = {
    "RestoreBondMarketFormulas_Debug": "WriteBondTickerFormulas",
}


def resolve_wave(row: dict, worder: dict[str, int], pnl_text: str):
    """The wave a non-PNL column is produced in, or None if nothing writes it.

    classify_columns names the procedure that physically writes the cell, which is
    not always the procedure the *FormulasSafe wrapper calls: the curve columns are
    written by shared helpers (WriteBQLDown_T0 and friends), and two bond columns
    are attributed to a debug routine that happens to write them too.  Both still
    have a real position in the run order; this finds it rather than leaving a hole
    that would look like "nothing produces this".
    """
    by = row.get("by") or ""
    if row.get("kind") == "UNWRITTEN" or not by:
        return None
    if by in INGEST_PROCS:
        return 0
    if by in worder:
        return worder[by]
    if by in DEBUG_WRITER_REAL:
        return worder.get(DEBUG_WRITER_REAL[by])

    # A shared helper: "WriteBQLDown_T0 (hard-coded letter)".  Take the bare name
    # and find which ordered writer's body calls it.
    helper = by.split(" (")[0]
    if helper in worder:
        return worder[helper]
    for writer, body in sorted(_writer_bodies(pnl_text, worder).items(),
                               key=lambda t: worder[t[0]]):
        if re.search(r"\b%s\b" % re.escape(helper), body):
            return worder[writer]
    return None


_BODY_CACHE: dict[int, dict[str, str]] = {}


def _writer_bodies(pnl_text: str, worder: dict[str, int]) -> dict[str, str]:
    """Each ordered writer's body, extracted once.

    Without the cache this re-scanned the whole 14k-line module per writer per
    unresolved column, which took minutes.
    """
    key = id(pnl_text)
    if key not in _BODY_CACHE:
        bodies = {}
        for m in re.finditer(
                r"^(?:Public |Private )?(?:Sub|Function) (\w+)\b(.*?)^End (?:Sub|Function)",
                pnl_text, re.M | re.S):
            if m.group(1) in worder:
                bodies[m.group(1)] = m.group(2)
        _BODY_CACHE[key] = bodies
    return _BODY_CACHE[key]


def add_provenance(sheets: dict[str, list[dict]], pnl_text: str,
                   others: list[str]) -> None:
    """Stamp SourceKind / ProducedBy / ProducedInWave / ConsumedBy onto every row."""
    kinds = classify_columns.classify(pnl_text, others)
    by_const = {r["const"]: r for r in kinds.values()}

    graph = field_graph.build()
    pnl_wave = graph["wave"]
    pnl_cons = graph["consumers"]
    pnl_field = {ltr: row[1] for ltr, row in graph["by_letter"].items()}

    worder = writer_order(pnl_text)
    # QUERY columns arrive before any writer runs, so they are wave 0 and the
    # writers start at 1.
    for sheet, rows in sheets.items():
        for r in rows:
            info = by_const.get(r["const"], {})
            r["kind"] = info.get("kind", "UNWRITTEN")
            r["by"] = info.get("by", "")
            if sheet == "PNL_Attribution":
                r["wave"] = pnl_wave.get(r["letter"])
                r["consumed"] = ",".join(
                    pnl_field.get(c, c) for c in pnl_cons.get(r["letter"], []))
            else:
                r["consumed"] = ""
                r["wave"] = resolve_wave(r, worder, pnl_text)


def main() -> int:
    pnl_text = modules.path_of("modPNL").read_text(encoding="latin-1")
    dash_text = modules.path_of("modDashboard").read_text(encoding="latin-1")

    # PnlLayout moved between modules across branches; take whichever has it.
    layout_source = pnl_text if pnl_layout(pnl_text) else dash_text
    if layout_source is dash_text:
        pnl_text = pnl_text + "\n" + dash_text

    sheets = collect(pnl_text)

    print(f"{'sheet':18} {'columns':>8}   {'first':<28} {'last'}")
    print("-" * 78)
    total = 0
    for sheet, rows in sorted(sheets.items()):
        if not rows:
            continue
        total += len(rows)
        print(f"{sheet:18} {len(rows):>8}   "
              f"{rows[0]['letter'] + ' ' + rows[0]['field']:<28} "
              f"{rows[-1]['letter']} {rows[-1]['field']}")
    print("-" * 78)
    print(f"{'':18} {total:>8}   registry rows")

    dupes = []
    for sheet, rows in sheets.items():
        seen: dict[str, str] = {}
        for r in rows:
            if r["field"] in seen:
                dupes.append(f"{sheet}: {r['field']} on both "
                             f"{seen[r['field']]} and {r['letter']}")
            seen[r["field"]] = r["letter"]
    if dupes:
        print("\nfield names that are not unique on their sheet:")
        for d in dupes:
            print("   ", d)
        print("   (Meta_Column indexes on (TargetSheet, FieldName), so these "
              "must be resolved before seeding)")

    if "--write" not in sys.argv:
        print("\n(--write to emit access/seed_meta_column.sql and "
              "access/alter_fact_pnl.sql)")
        return 1 if dupes else 0

    others = []
    for name in ("modDashboard", "modEconFormulas", "modImpRepo"):
        try:
            others.append(modules.path_of(name).read_text(encoding="latin-1"))
        except Exception:
            pass
    add_provenance(sheets, pnl_text, others)

    seed = ["-- GENERATED by tools/gen_column_registry.py - do not hand-edit.",
            "-- Regenerate after changing the workbook's column constants.",
            "--",
            "-- One INSERT per statement: Access executes one statement per",
            "-- Execute call and has no batch separator.",
            ""]
    for sheet, rows in sorted(sheets.items()):
        if not rows:
            continue
        seed.append(f"-- {sheet}: {len(rows)} columns")
        for i, r in enumerate(rows, start=1):
            wave = "Null" if r.get("wave") is None else str(r["wave"])
            seed.append(
                "INSERT INTO Meta_Column "
                "(TargetSheet, FieldName, DisplayLabel, Ordinal, IsVisible, "
                "SourceKind, ProducedBy, ProducedInWave, ConsumedBy) "
                f"VALUES ({sql_str(sheet)}, {sql_str(r['field'])}, "
                f"{sql_str(r['label'])}, {i * ORDINAL_STEP}, True, "
                f"{sql_str(r.get('kind', ''))}, {sql_str(r.get('by', ''))}, "
                f"{wave}, {sql_str(r.get('consumed', ''))});")
        seed.append("")
    (ROOT / "access" / "seed_meta_column.sql").write_text("\n".join(seed))

    pnl = sheets["PNL_Attribution"]
    keys = {"ISIN", "Portfolio"}          # already in the CREATE TABLE shell
    alter = ["-- GENERATED by tools/gen_column_registry.py - do not hand-edit.",
             "--",
             "-- The measure columns of Fact_PnlAttribution.  Generated from the",
             "-- same source as Meta_Column so the table and the registry cannot",
             "-- drift apart.  RunID, ISIN and Portfolio are in schema.sql.",
             ""]
    for r in pnl:
        if r["field"] in keys:
            continue
        alter.append(f"ALTER TABLE Fact_PnlAttribution "
                     f"ADD COLUMN [{r['field']}] {access_type(r['field'])};")
    (ROOT / "access" / "alter_fact_pnl.sql").write_text("\n".join(alter) + "\n")

    print(f"\nwrote access/seed_meta_column.sql   {total} rows")
    print(f"wrote access/alter_fact_pnl.sql     "
          f"{len(pnl) - len(keys & {r['field'] for r in pnl})} measure columns")
    return 1 if dupes else 0


if __name__ == "__main__":
    raise SystemExit(main())
