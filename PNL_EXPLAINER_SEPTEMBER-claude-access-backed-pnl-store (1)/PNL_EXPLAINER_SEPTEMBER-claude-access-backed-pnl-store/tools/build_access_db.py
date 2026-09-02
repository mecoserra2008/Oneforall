#!/usr/bin/env python3
"""Create a PnL run database and apply the full DDL.

access/schema.sql names this file as the way to apply itself.  It did not exist,
so the only executable path to a database was AccEnsureSchema in modAccess, which
creates SIX of the sixteen tables - the position store - and nothing else.  The
ten design-ahead tables (Curve_Point, Fact_BondRisk, Fact_HedgePosition,
Fact_PnlAttribution, Raw_CoverageRow, Raw_BloombergPoint, Instrument, Portfolio,
CoverageRelation, Meta_Column) could not be created by anything at all.

    python3 tools/build_access_db.py --check
        parse every .sql file, report the statements and the tables they make.
        Runs anywhere; this is what CI should call.

    python3 tools/build_access_db.py --create "M:\\path\\PNL_Run_20260901_173000.accdb"
        create the file and execute the DDL.  Windows only - it needs the Access
        Database Engine, matching the bitness of the Office that will read it.

    python3 tools/build_access_db.py --name-for 2026-09-01T17:30:00
        print the run-database filename for a retrieval timestamp.

ONE DATABASE PER RUN

Each run gets its own .accdb, named for the moment the data was retrieved.  The
Run table still exists inside it and still holds one row, so Run_Issue keeps its
parent and a history step can attach several files and UNION them.  The index
database (see --index) records where they all are.

Access executes ONE statement per Execute call and has no batch separator, which
is why the files are split into statements here rather than sent whole.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ACCESS = ROOT / "access"

#: Applied in this order.  The ALTERs must follow the CREATE they extend, and the
#: seed must follow the ALTER that adds the columns it fills.
DDL_ORDER = [
    "schema.sql",
    "alter_run_scope.sql",
    "alter_meta_column.sql",
    "alter_fact_pnl.sql",
    "seed_meta_column.sql",
]

#: Access rejects a statement it cannot parse, and the two most common reasons in
#: a hand-written .sql are a trailing semicolon inside a batch and a comment line
#: mid-statement.  Both are stripped here rather than in every caller.
COMMENT = re.compile(r"^\s*--.*$", re.M)


def statements(sql: str) -> list[str]:
    """Split a .sql file into executable statements, comments removed."""
    out = []
    for chunk in COMMENT.sub("", sql).split(";"):
        chunk = chunk.strip()
        if chunk:
            out.append(chunk)
    return out


def load_ddl() -> list[tuple[str, list[str]]]:
    files = []
    for name in DDL_ORDER:
        path = ACCESS / name
        if not path.exists():
            raise SystemExit(f"missing DDL file: access/{name}")
        files.append((name, statements(path.read_text(encoding="utf-8"))))
    return files


def run_db_name(stamp: str) -> str:
    """The run-database filename for an ISO-ish retrieval timestamp.

    The timestamp is when the data was RETRIEVED, not when the file was written,
    so a re-save of the same pull lands on the same name and a genuinely new pull
    never collides with an older one.
    """
    digits = re.sub(r"\D", "", stamp)
    if len(digits) < 14:
        raise SystemExit(
            f"need a full date and time, got {stamp!r} "
            "(e.g. 2026-09-01T17:30:00)")
    return f"PNL_Run_{digits[:8]}_{digits[8:14]}.accdb"


def check() -> int:
    files = load_ddl()
    total = 0
    tables, indexes, alters, inserts = [], 0, 0, 0
    for name, stmts in files:
        total += len(stmts)
        for s in stmts:
            head = s.split(None, 3)
            verb = " ".join(head[:2]).upper()
            if verb == "CREATE TABLE":
                tables.append(re.sub(r"[\[\]]", "", head[2]).strip("("))
            elif verb.startswith("CREATE") and "INDEX" in verb:
                indexes += 1
            elif verb == "ALTER TABLE":
                alters += 1
            elif verb == "INSERT INTO":
                inserts += 1
        print(f"  {name:<28} {len(stmts):>4} statements")
    print(f"\n  {'total':<28} {total:>4} statements")
    print(f"\n  {len(tables)} tables, {indexes} indexes, {alters} alters, "
          f"{inserts} seed rows")
    print("\n  tables: " + ", ".join(sorted(tables)))

    missing = [t for t in ("Run", "Pos_Bond", "Pos_Swap", "Pos_Future",
                           "Meta_Column", "Curve_Point", "Fact_BondRisk",
                           "Fact_HedgePosition", "Fact_PnlAttribution",
                           "Raw_CoverageRow", "Raw_BloombergPoint")
               if t not in tables]
    if missing:
        print("\n  MISSING from the DDL: " + ", ".join(missing))
        return 1
    return 0


def create(target: Path) -> int:
    if sys.platform != "win32":
        raise SystemExit(
            "--create needs Windows and the Access Database Engine.\n"
            "Use --check here; run --create on the desk machine.")
    import win32com.client  # noqa: PLC0415  (Windows-only dependency)

    if target.exists():
        raise SystemExit(f"refusing to overwrite an existing run database: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)

    conn_for = ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source={};",
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={};")
    made = False
    for tmpl in conn_for:
        try:
            cat = win32com.client.Dispatch("ADOX.Catalog")
            cat.Create(tmpl.format(target))
            made = True
            break
        except Exception:
            continue
    if not made:
        raise SystemExit(
            "could not create the .accdb - is the Access Database Engine "
            "installed, matching the bitness of your Office?")

    conn = win32com.client.Dispatch("ADODB.Connection")
    for tmpl in conn_for:
        try:
            conn.Open(tmpl.format(target))
            break
        except Exception:
            continue

    applied = failed = 0
    for name, stmts in load_ddl():
        for s in stmts:
            try:
                conn.Execute(s)
                applied += 1
            except Exception as exc:                      # noqa: BLE001
                failed += 1
                print(f"  FAILED in {name}: {exc}\n    {s[:120]}")
    conn.Close()
    print(f"\ncreated {target}\n  {applied} statements applied, {failed} failed")
    return 1 if failed else 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--check", action="store_true",
                   help="parse the DDL and report; runs anywhere")
    g.add_argument("--create", metavar="PATH",
                   help="create the .accdb and apply the DDL (Windows only)")
    g.add_argument("--name-for", metavar="TIMESTAMP",
                   help="print the run-database filename for a retrieval timestamp")
    a = ap.parse_args(argv)

    if a.name_for:
        print(run_db_name(a.name_for))
        return 0
    if a.check:
        return check()
    return create(Path(a.create))


if __name__ == "__main__":
    sys.exit(main())
