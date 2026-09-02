#!/usr/bin/env python3
r"""Classify every workbook column by WHERE ITS VALUE COMES FROM.

The Access split turns on one question asked of each of the 330 columns: is this
value FETCHED from outside Excel, or COMPUTED from values already fetched?

  fetched   must happen in Excel - Bloomberg add-in functions evaluate nowhere
            else - or in a query (OPICS, the coverage books).  These are the
            INPUTS, and they are what Access stores raw.
  computed  a pure transformation of columns already on a sheet.  These can move
            to Access and be evaluated once per run instead of being held as a
            live formula in every row.

Answering that by reading 16,000 lines is how the answer goes stale.  This reads
the writers and derives it.

    python3 tools/classify_columns.py            # the summary table
    python3 tools/classify_columns.py --csv      # one row per column
    python3 tools/classify_columns.py --sheet Bonds
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402

SHEETS = {"BCOL_": "Bonds", "PCOL_": "PNL_Attribution", "FCOL_": "Futures",
          "WCOL_": "Swaps", "CVCOL_": "OIS_Curves"}

# What makes a formula a FETCH rather than a computation.
BBG_RE = re.compile(r"\bBD[PHS]\s*\(|\bBQL\b|BQL\.", re.I)
# Roll-ups: computed, but ACROSS rows rather than within one.
AGG_RE = re.compile(r"\b(SUMIFS|COUNTIFS|SUMIF|COUNTIF|SUMPRODUCT)\s*\(", re.I)
# Parameter names that mean "the column this helper WRITES".
OUT_PARAMS = ("outcol", "targetcol", "col")


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def flat(src: str) -> str:
    """Join `_` continuations so one logical statement is one string."""
    return re.sub(r"_\r?\n\s*", " ", src)


def procedures(text: str) -> dict[str, str]:
    return {m.group(1): m.group(0) for m in re.finditer(
        r"^(?:Public |Private )?(?:Sub|Function) (\w+)\b.*?"
        r"^End (?:Sub|Function)[ \t]*$", text, re.M | re.S)}


def bbg_helpers(procs: dict[str, str]) -> set[str]:
    """Helpers that emit a Bloomberg call, transitively.

    A writer says `.FormulaR1C1 = BBGFirstBDPMultiFieldFormulaR1C1(...)`, so the
    BDP is a level or two down.  Without following it, every Bloomberg column
    reads as an ordinary derived one - the opposite of the truth.
    """
    direct = {n for n, b in procs.items() if BBG_RE.search(b)}
    changed = True
    while changed:
        changed = False
        for name, body in procs.items():
            if name in direct:
                continue
            if any(re.search(r"\b" + h + r"\b", body) for h in direct):
                direct.add(name)
                changed = True
    return direct


def udf_names(texts: list[str]) -> set[str]:
    out: set[str] = set()
    for t in texts:
        out |= set(re.findall(r"^Public Function (\w+)", t, re.M))
    return out


def classify(pnl: str, others: list[str],
             report: list | None = None) -> dict[str, dict]:
    procs = procedures(pnl)
    bbg = bbg_helpers(procs)
    udfs = udf_names([pnl] + others)
    fpnl = flat(pnl)

    consts = {m.group(1): m.group(2) for m in re.finditer(
        r'^Private Const (\w+) As String = "([A-Z]{1,3})"', pnl, re.M)}
    labels = dict(re.findall(
        r'^Private Const (\w+) As String = "[A-Z]{1,3}"\s*\'\s*col \d+ = (\S+)',
        pnl, re.M))
    for c, k in re.findall(
            r'AddPnlCol spec,\s*(\w+),\s*"([^"]+)"', pnl):
        labels[c] = k

    found: dict[str, dict] = {}

    # --- 1) formula writes: .Range(XCOL_x & ...).Formula|.FormulaR1C1 = <rhs> --
    assign = re.compile(
        r"\.(?:Range\(\s*(?P<r>\w+COL_\w+)[^)]*"
        r"|Cells\([^,]+,\s*(?:colNum\(\s*)?(?P<c>\w+COL_\w+)\s*\)?\s*)"
        r"\)\s*\.\s*(?:FormulaR1C1|[Ff]ormula)\s*=\s*(?P<rhs>.+?)$")
    for pname, body in procs.items():
        local: dict[str, str] = {}
        for line in flat(body).split("\n"):
            m = re.match(r"^\s*(\w+)\s*=\s*(.+)$", line)
            if m and ".Range(" not in line:
                v, rhs = m.group(1), m.group(2)
                if re.match(r"^" + re.escape(v) + r"\s*&", rhs.strip()):
                    local[v] = local.get(v, "") + " " + rhs
                else:
                    local[v] = rhs
            a = assign.search(line)
            if not a:
                continue
            const = a.group("r") or a.group("c")
            rhs = a.group("rhs").strip()
            expr = local.get(rhs, rhs) + " " + rhs
            called = set(re.findall(r"\b(\w+)\s*\(", expr))
            if (called & bbg) or BBG_RE.search(expr):
                kind = "BLOOMBERG"
            elif AGG_RE.search(expr) or any(
                    AGG_RE.search(procs.get(c, "")) for c in called if c in procs):
                kind = "AGGREGATE"
            elif called & udfs:
                kind = "UDF"
            else:
                kind = "DERIVED"
            found[const] = {"kind": kind, "by": pname}

    # --- 2) the column passed to a Bloomberg writer as its OUTPUT --------------
    #
    # The curve and prior-snapshot pulls assign no formula to a range: they call
    # `WriteBDPDown_Efficient ws, tickerCol, outCol, ...`, so the column is an
    # ARGUMENT.  Missing this shape reported OIS_Curves as having no Bloomberg
    # columns at all.
    #
    # Which argument is the output is read from the helper's own parameter names
    # rather than assumed by position: tickerCol is what it READS, outCol what
    # it WRITES.
    #
    # Some call sites pass a bare LETTER rather than the constant - see
    # hardcoded[] below, which is a finding in its own right - so a literal is
    # resolved back to whichever constant owns that letter on the sheet the
    # CALLING procedure writes.
    caller_sheet = {"Curve": "OIS_Curves", "Bond": "Bonds", "Future": "Futures",
                    "Swap": "Swaps"}
    by_letter: dict[str, dict[str, str]] = defaultdict(dict)
    for const, letter in consts.items():
        for pre, sheet in SHEETS.items():
            if const.startswith(pre):
                by_letter[sheet][letter] = const
                break

    hardcoded: list[tuple[str, str, str]] = []
    for pname, body in procs.items():
        sheet = next((s for k, s in caller_sheet.items() if k in pname), None)
        fbody = flat(body)
        for hname in sorted(bbg):
            sig = re.search(r"^(?:Public |Private )?Sub " + hname + r"\((.*?)\)\s*$",
                            fpnl, re.M)
            if not sig:
                continue
            params = [re.sub(r"^(?:ByVal|ByRef)\s+", "", q.strip()).split()[0]
                      for q in sig.group(1).split(",") if q.strip()]
            outs = [i for i, q in enumerate(params) if q.lower() in OUT_PARAMS]
            if not outs:
                continue
            for m in re.finditer(r"\b" + hname + r"[ \t]+([^\r\n]+)", fbody):
                args = [a.strip() for a in m.group(1).split(",")]
                for i in outs:
                    if i >= len(args):
                        continue
                    arg = args[i]
                    if re.fullmatch(r"\w+COL_\w+", arg):
                        found.setdefault(arg, {"kind": "BLOOMBERG", "by": hname})
                    else:
                        lit = re.fullmatch(r'"([A-Z]{1,3})"', arg)
                        if lit and sheet:
                            const = by_letter[sheet].get(lit.group(1))
                            hardcoded.append((pname, hname, lit.group(1)))
                            if const:
                                found.setdefault(
                                    const, {"kind": "BLOOMBERG",
                                            "by": hname + " (hard-coded letter)"})

    # --- 3) value writes: a recordset field or a mapped value = a QUERY input --
    val = re.compile(
        r"\.(?:Range\(\s*(?P<r>\w+COL_\w+)[^)]*"
        r"|Cells\([^,]+,\s*(?:colNum\(\s*)?(?P<c>\w+COL_\w+)\s*\)?\s*)"
        r"\)\s*\.\s*value\s*=")
    for pname, body in procs.items():
        for m in val.finditer(flat(body)):
            found.setdefault(m.group("r") or m.group("c"),
                             {"kind": "QUERY", "by": pname})

    # --- 4) the block the OPICS query owns ------------------------------------
    #
    # Bonds!A:BONDS_QUERY_LAST_COL is written by the Excel query, not by VBA, so
    # no writer will ever be found for it - but it is unambiguously an INPUT and
    # the module declares where it stops.
    q = re.search(r'Private Const BONDS_QUERY_LAST_COL As String = "([A-Z]+)"', pnl)
    if q:
        limit = col_num(q.group(1))
        for const, letter in consts.items():
            if const.startswith("BCOL_") and col_num(letter) <= limit:
                found.setdefault(const,
                                 {"kind": "QUERY", "by": "OPICS query table"})

    if report is not None:
        report.extend(hardcoded)

    rows: dict[str, dict] = {}
    for const, letter in consts.items():
        for pre, sheet in SHEETS.items():
            if const.startswith(pre):
                info = found.get(const, {"kind": "UNWRITTEN", "by": ""})
                rows[const] = {"sheet": sheet, "letter": letter,
                               "col": col_num(letter),
                               "field": labels.get(const, const[len(pre):]),
                               "const": const, **info}
                break
    return rows


ORDER = ["BLOOMBERG", "QUERY", "UDF", "AGGREGATE", "DERIVED", "UNWRITTEN"]


def main() -> int:
    pnl = modules.path_of("modPNL").read_text(encoding="latin-1")
    others = [modules.path_of(m).read_text(encoding="latin-1")
              for m in ("modDashboard", "modEconFormulas", "modImpRepo")]
    hardcoded: list = []
    rows = classify(pnl, others, hardcoded)

    if "--sheet" in sys.argv:
        want = sys.argv[sys.argv.index("--sheet") + 1]
        for r in sorted(rows.values(), key=lambda r: (r["sheet"], r["col"])):
            if r["sheet"] == want:
                print(f'  {r["letter"]:>3}  {r["field"]:<30} {r["kind"]:<10} {r["by"]}')
        return 0

    if "--csv" in sys.argv:
        print("sheet,letter,field,kind,written_by")
        for r in sorted(rows.values(), key=lambda r: (r["sheet"], r["col"])):
            print(f'{r["sheet"]},{r["letter"]},{r["field"]},{r["kind"]},{r["by"]}')
        return 0

    by_sheet: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for r in rows.values():
        by_sheet[r["sheet"]][r["kind"]] += 1

    w = max(len(s) for s in by_sheet)
    print(f"{'sheet':{w}} " + "".join(f"{k:>11}" for k in ORDER) + f"{'total':>8}")
    print("-" * (w + 11 * len(ORDER) + 8))
    tot: dict[str, int] = defaultdict(int)
    for sheet in sorted(by_sheet):
        c = by_sheet[sheet]
        for k in ORDER:
            tot[k] += c[k]
        print(f"{sheet:{w}} " + "".join(f"{c[k] or '':>11}" for k in ORDER)
              + f"{sum(c.values()):>8}")
    print("-" * (w + 11 * len(ORDER) + 8))
    print(f"{'':{w}} " + "".join(f"{tot[k]:>11}" for k in ORDER)
          + f"{sum(tot.values()):>8}")

    fetched = tot["BLOOMBERG"] + tot["QUERY"]
    computed = tot["UDF"] + tot["AGGREGATE"] + tot["DERIVED"]
    print(f"\n  fetched  (Excel or a query must do it): {fetched:>4}")
    print(f"  computed (can move to Access):          {computed:>4}")
    print(f"  unwritten (declared, nothing fills it): {tot['UNWRITTEN']:>4}")

    if hardcoded:
        seen = sorted(set(hardcoded))
        print(f"\n  {len(seen)} Bloomberg write(s) target a HARD-CODED column "
              f"letter, not a constant:")
        for pname, hname, letter in seen[:12]:
            print(f'      {pname} -> {hname}(... , "{letter}")')
        if len(seen) > 12:
            print(f"      ... and {len(seen) - 12} more")
        print("      The XCOL_ constants beside them claim to be the contract;"
              "\n      moving one of those columns would not move the pull.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
