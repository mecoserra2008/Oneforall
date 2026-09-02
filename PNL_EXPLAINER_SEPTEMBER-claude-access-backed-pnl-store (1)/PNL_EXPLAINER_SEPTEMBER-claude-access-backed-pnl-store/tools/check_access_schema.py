#!/usr/bin/env python3
"""Keep the three descriptions of the Access store in agreement.

There are three, and they are written in three languages, so nothing but a
check like this can stop them drifting:

  access/schema.sql        the DDL a human reads, and the one that gets pasted
                           into Access when somebody builds the database by hand
  modAccess AccTableSpec   the DDL that actually runs, as arrays
  modPNL PnlPositionFieldMap
                           which sheet column each stored field comes from

Drift between them does not fail loudly.  A field in the sheet map with no
column in the table is refused by the provider at save time - one row at a
time, inside a transaction, on a desk machine.  A column in the table that no
sheet field feeds is simply always Null, which reads as "we never had that
position detail" rather than as a bug.

  A001  a table modAccess creates is not in schema.sql
  A002  a column differs between modAccess and schema.sql
  A003  a stored field has no column in its table
  A004  a Pos_ column nothing feeds (and is not bookkeeping)
  A005  a position table is missing the run/key primary key

Run:  python3 tools/check_access_schema.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SCHEMA = ROOT / "access" / "schema.sql"

# Columns every position table carries for its own sake rather than because a
# sheet supplies them.
BOOKKEEPING = {"RunID", "PositionKey"}

# Which stored kind feeds which table.
KIND_TABLE = {
    "STORE_KIND_BOND": "Pos_Bond",
    "STORE_KIND_SWAP": "Pos_Swap",
    "STORE_KIND_FUTURE": "Pos_Future",
}


def parse_schema_sql(text: str) -> dict[str, dict[str, str]]:
    """table -> {column: type}, from the CREATE TABLE statements."""
    out: dict[str, dict[str, str]] = {}
    for m in re.finditer(
            r"CREATE TABLE\s+\[?(\w+)\]?\s*\((.*?)\n\);", text, re.S):
        table, body = m.group(1), m.group(2)
        cols: dict[str, str] = {}
        for line in body.splitlines():
            line = line.strip().rstrip(",")
            if not line or line.startswith("--"):
                continue
            if line.upper().startswith("CONSTRAINT") or \
                    line.upper().startswith("PRIMARY KEY"):
                continue
            cm = re.match(r"\[?(\w+)\]?\s+(.+)$", line)
            if cm:
                cols[cm.group(1)] = normalise_type(cm.group(2))
        out[table] = cols
    return out


def normalise_type(t: str) -> str:
    t = t.strip().rstrip(",")
    t = re.sub(r"\s+NOT\s+NULL$", "", t, flags=re.I)
    return re.sub(r"\s+", "", t).upper()


def parse_vba_specs(text: str) -> dict[str, dict[str, str]]:
    """table -> {column: type}, from AccTableSpec's Case arms."""
    body = _function_body(text, "AccTableSpec")
    out: dict[str, dict[str, str]] = {}
    table = None
    for line in body.splitlines():
        cm = re.match(r'\s*Case\s+"(\w+)"', line)
        if cm:
            table = cm.group(1)
            out[table] = {}
            continue
        if table is None:
            continue
        for em in re.finditer(r'"(\w+)\|([^"]+)"', line):
            out[table][em.group(1)] = normalise_type(em.group(2))
    return out


def parse_vba_pk(text: str) -> dict[str, str]:
    body = _function_body(text, "AccTablePK")
    out: dict[str, str] = {}
    for m in re.finditer(r'Case\s+"(\w+)":\s*AccTablePK\s*=\s*"([^"]*)"', body):
        out[m.group(1)] = m.group(2)
    return out


def parse_field_map(text: str) -> dict[str, list[str]]:
    """STORE_KIND_x -> [field names], from PnlPositionFieldMap."""
    body = _function_body(text, "PnlPositionFieldMap")
    out: dict[str, list[str]] = {}
    kind = None
    for line in body.splitlines():
        cm = re.match(r"\s*Case\s+(STORE_KIND_\w+)", line)
        if cm:
            kind = cm.group(1)
            out[kind] = []
            continue
        if kind is None:
            continue
        for em in re.finditer(r'"(\w+)\|"', line):
            out[kind].append(em.group(1))
    return out


def _function_body(text: str, name: str) -> str:
    m = re.search(
        r"^(?:Public |Private )?Function " + re.escape(name) + r"\b.*?^End Function",
        text, re.M | re.S)
    if not m:
        raise SystemExit(f"cannot find Function {name}")
    return m.group(0)


def main() -> int:
    schema = parse_schema_sql(SCHEMA.read_text(encoding="utf-8"))
    access_text = modules.path_of("modAccess").read_text(encoding="latin-1")
    pnl_text = modules.path_of("modPNL").read_text(encoding="latin-1")

    specs = parse_vba_specs(access_text)
    pks = parse_vba_pk(access_text)
    fields = parse_field_map(pnl_text)

    problems: list[str] = []

    for table, cols in sorted(specs.items()):
        if table not in schema:
            problems.append(
                f"A001 modAccess creates [{table}] but access/schema.sql has "
                f"no CREATE TABLE for it")
            continue
        for col, vtype in sorted(cols.items()):
            stype = schema[table].get(col)
            if stype is None:
                problems.append(
                    f"A002 {table}.{col} is created by modAccess but is not "
                    f"in access/schema.sql")
            elif stype != vtype:
                problems.append(
                    f"A002 {table}.{col} is {vtype} in modAccess and {stype} "
                    f"in access/schema.sql")
        for col in sorted(schema[table]):
            if col not in cols:
                problems.append(
                    f"A002 {table}.{col} is in access/schema.sql but "
                    f"modAccess would never create it")

    for kind, table in sorted(KIND_TABLE.items()):
        stored = fields.get(kind)
        if stored is None:
            problems.append(
                f"A003 modPNL has no PnlPositionFieldMap arm for {kind}")
            continue
        cols = specs.get(table, {})
        for name in stored:
            if name not in cols:
                problems.append(
                    f"A003 {kind} stores '{name}' but [{table}] has no such "
                    f"column")
        for col in sorted(cols):
            if col in BOOKKEEPING:
                continue
            if col not in stored:
                problems.append(
                    f"A004 [{table}].{col} exists but no {kind} field feeds "
                    f"it - it would always be Null")

    for table in sorted(KIND_TABLE.values()):
        if pks.get(table) != "RunID, PositionKey":
            problems.append(
                f"A005 [{table}] must be keyed on (RunID, PositionKey) - the "
                f"run separation and the duplicate guard are both that key; "
                f"found {pks.get(table)!r}")

    for line in problems:
        print(line)

    print(f"\n{len(specs)} table(s), {sum(len(v) for v in fields.values())} "
          f"stored field(s), {len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
