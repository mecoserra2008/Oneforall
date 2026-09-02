#!/usr/bin/env python3
"""Sheet-geometry contract checks for the PNL Explainer workbook.

The whole module addresses cells through letter constants (BCOL_*, PCOL_*, ...)
so that a column can be moved in one place.  That only holds while the
constants themselves stay consistent, and nothing in Excel will tell you when
they stop:

  L001  a column letter is used by two constants on the same sheet
  L002  the "' col N =" comment does not match the letter
  L003  a column is skipped without a declared spacer constant
  L004  a header array's element count does not match its target range width
  L005  Bonds row geometry is inconsistent with the reserved comment band
  L006  a Bloomberg writer touches a column its range builder does not cover
  L007  the PNL_Attribution / Dashboard header contract has drifted
  L008  a Bonds/PNL write targets a hard-coded column letter

Run:  python3 tools/check_layout.py
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent

CONST_RE = re.compile(
    r'^\s*Private Const (?P<name>[A-Z][A-Z0-9_]*)\s+As String\s*=\s*'
    r'"(?P<letter>[A-Z]{1,3})"\s*(?:\'\s*col\s+(?P<num>\d+))?',
    re.M,
)

SHEET_PREFIXES = {
    "BCOL_": "Bonds",
    "PCOL_": "PNL_Attribution",
    "FCOL_": "Futures",
    "WCOL_": "Swaps",
    "CVCOL_": "OIS_Curves",
    "FMCOL_": "FutMap",
}

# Columns a sheet deliberately leaves empty.  A gap that is NOT listed here and
# has no spacer constant is reported: it is usually a column that was moved and
# left a hole, or a query-owned column nobody named.
DECLARED_GAPS = {
    # OIS_Curves lays the three currencies out in blocks with a blank between.
    "OIS_Curves": {17, 18, 34, 35},
    # Bonds A:L is query-owned; H has no macro constant because nothing reads it.
    "Bonds": {8},
}


class Problem:
    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message

    def __str__(self) -> str:
        return f"{self.code} {self.message}"


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def load_constants(text: str) -> dict[str, list[tuple[str, str, str | None]]]:
    by_sheet: dict[str, list[tuple[str, str, str | None]]] = defaultdict(list)
    for m in CONST_RE.finditer(text):
        name = m.group("name")
        for prefix, sheet in SHEET_PREFIXES.items():
            if name.startswith(prefix):
                by_sheet[sheet].append((name, m.group("letter"), m.group("num")))
                break
    return by_sheet


# --------------------------------------------------------------------------

def check_columns(by_sheet) -> list[Problem]:
    out: list[Problem] = []
    for sheet, entries in sorted(by_sheet.items()):
        seen: dict[str, str] = {}
        prev = 0
        gaps_allowed = DECLARED_GAPS.get(sheet, set())

        for name, letter, num in entries:
            n = col_num(letter)

            if letter in seen:
                out.append(Problem(
                    "L001",
                    f"{sheet}: column {letter} claimed by both "
                    f"{seen[letter]} and {name}"))
            else:
                seen[letter] = name

            if num is not None and int(num) != n:
                out.append(Problem(
                    "L002",
                    f"{sheet}: {name} = \"{letter}\" is column {n}, but its "
                    f"comment says col {num}"))

            if n > prev + 1:
                missing = [c for c in range(prev + 1, n) if c not in gaps_allowed]
                if missing:
                    out.append(Problem(
                        "L003",
                        f"{sheet}: column(s) {missing} skipped before {name} "
                        f"(\"{letter}\") with no spacer constant and no entry "
                        f"in DECLARED_GAPS"))
            prev = max(prev, n)
    return out


HEADER_ARRAY_RE = re.compile(
    r"\.Range\((\w+)\s*&\s*\w+\s*&\s*\":\"\s*&\s*(\w+)\s*&\s*\w+\)"
    r"\.value\s*=\s*Array\((.*)\)$")


def check_header_arrays(text: str) -> list[Problem]:
    """A header block writes Array(...) across COL_A..COL_B in one shot.

    Excel silently pads with #N/A when the array is short and silently drops
    the tail when it is long, so a miscount shows up as one wrong header - or
    as a header landing one column left of where every formula expects it.
    Removing a column from the middle of a block is exactly when this happens.
    """
    out: list[Problem] = []
    letters = {m.group("name"): col_num(m.group("letter"))
               for m in CONST_RE.finditer(text)}

    for line in _logical_lines(text):
        m = HEADER_ARRAY_RE.search(line)
        if not m:
            continue
        first, last, body = m.groups()
        if first not in letters or last not in letters:
            continue
        width = letters[last] - letters[first] + 1
        count = len(re.findall(r'"[^"]*"', body))
        if count != width:
            out.append(Problem(
                "L004",
                f"header block {first}..{last} spans {width} columns but "
                f"writes {count} headers"))
    return out


# Procedures whose first argument after the worksheet IS a column, so a string
# literal there is a hard-coded column rather than a value.
LITERAL_TARGET_RE = re.compile(
    r"\b(PutBondF)\s+\w+\s*,\s*\"([A-Z]{1,3})\"")


def check_literal_columns(text: str) -> list[Problem]:
    """A write must name its column through a constant, never a letter.

    WriteBondT0BQLFormulas used to write to "S", "AB", "AS" ... while the
    AddToUnion beside each one named the BCOL_ constant.  The two agreed only
    by luck.  The moment a column moved, the T-1 formula was written to one
    column while the Bloomberg wait registered another - so the pull landed on
    top of a T0 column and the intended column stayed blank, with no error
    anywhere.
    """
    out: list[Problem] = []
    for line in _logical_lines(text):
        for m in LITERAL_TARGET_RE.finditer(line):
            out.append(Problem(
                "L008",
                f"{m.group(1)} writes to the hard-coded column "
                f"\"{m.group(2)}\" - use the BCOL_ constant so the target "
                f"moves with the column"))
    return out


def check_bond_rows(text: str) -> list[Problem]:
    out: list[Problem] = []

    m = re.search(r"Private Const BOND_COMMENT_ROWS\s+As Long\s*=\s*(\d+)", text)
    if not m:
        return [Problem("L005", "BOND_COMMENT_ROWS is not declared - the "
                                "reserved comment band on Bonds has no owner")]
    comment_rows = int(m.group(1))

    if not re.search(r"Private Const BOND_HEADER_ROW\s+As Long\s*=\s*"
                     r"BOND_COMMENT_ROWS\s*\+\s*1", text):
        out.append(Problem(
            "L005",
            "BOND_HEADER_ROW must be derived as BOND_COMMENT_ROWS + 1, so the "
            "comment band and the header cannot drift apart"))

    if not re.search(r"Private Const BOND_DATA_ROW\s+As Long\s*=\s*"
                     r"BOND_HEADER_ROW\s*\+\s*1", text):
        out.append(Problem(
            "L005",
            "BOND_DATA_ROW must be derived as BOND_HEADER_ROW + 1"))

    if comment_rows < 2:
        out.append(Problem(
            "L005",
            f"BOND_COMMENT_ROWS is {comment_rows}; the desk reserves the first "
            f"2 rows of Bonds for comments"))

    # The Bonds sheet is query-driven, so something must move the query table
    # when the band changes.  A constant on its own is not enough.
    if "EnsureBondsCommentRows" not in text:
        out.append(Problem(
            "L005",
            "no EnsureBondsCommentRows: changing BOND_COMMENT_ROWS would move "
            "the macro columns without moving the OPICS query table"))

    # The macro block must start in the column right after the query's last.
    qm = re.search(r'Private Const BONDS_QUERY_LAST_COL\s+As String\s*=\s*'
                   r'"([A-Z]{1,3})"', text)
    fm = re.search(r'Private Const BCOL_DAYSLEFT\s+As String\s*=\s*'
                   r'"([A-Z]{1,3})"', text)
    if not qm:
        out.append(Problem(
            "L005",
            "BONDS_QUERY_LAST_COL is not declared - nothing states where the "
            "OPICS query stops and the macro columns begin"))
    elif fm:
        want = col_num(qm.group(1)) + 1
        got = col_num(fm.group(1))
        if got != want:
            out.append(Problem(
                "L005",
                f"the first macro column BCOL_DAYSLEFT is column {got}, but "
                f"the query ends at {qm.group(1)} so it must be column {want}. "
                f"A one-column disagreement writes macro output over a query "
                f"column, or leaves a blank stripe, and every bond row's two "
                f"halves come apart"))

    # Nothing outside the Bonds writers may borrow the bond row constants.
    for m in re.finditer(r"^(?P<line>.*BOND_HEADER_ROW.*)$", text, re.M):
        line = m.group("line")
        if "FMCOL_" in line or "FCOL_" in line or "WCOL_" in line \
                or "CVCOL_" in line:
            out.append(Problem(
                "L005",
                f"a non-Bonds sheet uses BOND_HEADER_ROW as a stand-in for its "
                f"own header row: {line.strip()[:90]}"))
    return out


# --- L006: Bloomberg writers vs their range builders -----------------------

BBG_CONTRACT = {
    # range builder            -> writers whose output it must cover
    "BondsBbgRange": [
        "WriteBonds_T1_BDP_Efficient",
        "WriteBondT0BQLFormulas",
        "WriteBondTickerFormulas",
    ],
    "FuturesBbgRange": [
        "WriteFutures_T1_BDP_Efficient",
        "WriteFutureT0BQLFormulas",
    ],
    "SwapsBbgRange": [
        "WriteSwaps_T1_BQL_Formulas",
        "WriteSwapsT0BQLFormulas",
    ],
}

PREFIX_FOR = {
    "BondsBbgRange": "BCOL_",
    "FuturesBbgRange": "FCOL_",
    "SwapsBbgRange": "WCOL_",
}


def _procedure_source(text: str, name: str) -> str:
    m = re.search(r"^(?:Public |Private )?(?:Sub|Function) " +
                  re.escape(name) + r"\b", text, re.M)
    if not m:
        return ""
    end = re.search(r"^End (?:Sub|Function)\s*$", text[m.start():], re.M)
    return text[m.start():m.start() + end.end()] if end else text[m.start():]


BBG_CALL_RE = re.compile(r"\bBD[PH]\s*\(|\bBQL\b|BQL\.", re.I)


def _logical_lines(src: str) -> list[str]:
    """Join `_` continuations, so a formula built across ten lines is one
    string to search."""
    out: list[str] = []
    buf: list[str] = []
    for raw in src.splitlines():
        body = raw.split("'")[0].rstrip() if raw.strip().startswith("'") else raw.rstrip()
        cont = body.endswith("_") and (len(body) < 2 or body[-2] in " \t")
        buf.append(body[:-1] if cont else body)
        if not cont:
            out.append(" ".join(b.strip() for b in buf))
            buf = []
    if buf:
        out.append(" ".join(b.strip() for b in buf))
    return out


def _columns_written(src: str, prefix: str) -> dict[str, str]:
    """Columns a writer assigns a formula to, mapped to the formula TEXT.

    Formulas here are frequently accumulated into a local (`fml = ...` then
    `fml = fml & ...`) before being assigned, so a naive scan of the assignment
    line sees only the variable name.  Tracking those accumulations is what
    lets the check distinguish a real Bloomberg pull from an ordinary derived
    formula - without it the check reports every derived column as a missing
    Bloomberg wait, and a check that cries wolf gets switched off.
    """
    written: dict[str, str] = {}
    locals_: dict[str, str] = {}

    assign_var = re.compile(r"^\s*(\w+)\s*=\s*(.*)$")
    to_range = re.compile(
        r"\.Range\(\s*(" + prefix + r"\w+)\s*&(.*?)\)\s*\.\s*"
        r"(?:FormulaR1C1|[Ff]ormula)\s*=\s*(.*)$")

    for line in _logical_lines(src):
        m = to_range.search(line)
        if m:
            const, rhs = m.group(1), m.group(3).strip()
            text = locals_.get(rhs, rhs)
            written[const] = written.get(const, "") + " " + text
            continue

        m = assign_var.match(line)
        if m and ".Range(" not in line:
            name, rhs = m.group(1), m.group(2)
            if re.match(r"^" + re.escape(name) + r"\s*&", rhs.strip()):
                locals_[name] = locals_.get(name, "") + " " + rhs
            else:
                locals_[name] = rhs

    for m in re.finditer(
            r"AddToUnion\s+\w+\s*,\s*ws\.Range\(\s*(" + prefix + r"\w+)",
            src):
        written.setdefault(m.group(1), "BQL")

    return written


def _columns_covered(src: str, prefix: str, letters: dict[str, int]) -> set[int]:
    """Column numbers a range builder unions, expanding A:B style spans."""
    covered: set[int] = set()
    for m in re.finditer(
            r"AddToUnion\s+\w+\s*,\s*ws\.Range\(\s*(" + prefix + r"\w+)"
            r"(?:[^\n]*?\":\"\s*&\s*(" + prefix + r"\w+))?", src):
        a, b = m.group(1), m.group(2)
        if a not in letters:
            continue
        lo = letters[a]
        hi = letters.get(b, lo) if b else lo
        covered.update(range(min(lo, hi), max(lo, hi) + 1))
    return covered


def check_bbg_ranges(text: str, by_sheet) -> list[Problem]:
    out: list[Problem] = []
    letters_by_prefix: dict[str, dict[str, int]] = defaultdict(dict)
    for m in CONST_RE.finditer(text):
        name = m.group("name")
        for prefix in SHEET_PREFIXES:
            if name.startswith(prefix):
                letters_by_prefix[prefix][name] = col_num(m.group("letter"))
                break

    for builder, writers in BBG_CONTRACT.items():
        prefix = PREFIX_FOR[builder]
        letters = letters_by_prefix[prefix]

        builder_src = _procedure_source(text, builder)
        if not builder_src:
            out.append(Problem("L006", f"range builder {builder} not found"))
            continue

        covered = _columns_covered(builder_src, prefix, letters)

        for writer in writers:
            wsrc = _procedure_source(text, writer)
            if not wsrc:
                out.append(Problem(
                    "L006", f"{builder} names writer {writer}, which does not "
                            f"exist"))
                continue
            for const, formula in sorted(_columns_written(wsrc, prefix).items()):
                if not BBG_CALL_RE.search(formula):
                    continue          # ordinary derived formula, nothing to wait for
                n = letters.get(const)
                if n is None:
                    continue
                if n not in covered:
                    out.append(Problem(
                        "L006",
                        f"{writer} writes a Bloomberg formula into {const}, "
                        f"which {builder} does not cover - RefreshMarketData "
                        f"would never wait for it"))
    return out


# --- L004 / L007: the PNL column contract ---------------------------------


# Parameters that carry a PNL contract key into a Dashboard helper.  Driven off
# the parameter NAME rather than a hard-coded list of call shapes: almost every
# key reaches DashN through a wrapper (DashSumFml, DashIndexCell, ...), and a
# checker that only understood DashN("...") silently stopped catching anything
# the moment a new wrapper was added - which is how it came to pass a module
# that asked for a column nothing declared.
KEY_PARAM_NAMES = ("headerName", "contractKey", "metricHeader")


def _split_args(s: str) -> list[str]:
    """Top-level comma split of a VBA argument list, respecting () and quotes."""
    out, depth, cur, in_str = [], 0, [], False
    i = 0
    while i < len(s):
        ch = s[i]
        if in_str:
            cur.append(ch)
            if ch == '"':
                if i + 1 < len(s) and s[i + 1] == '"':
                    cur.append(s[i + 1]); i += 2; continue
                in_str = False
        elif ch == '"':
            in_str = True; cur.append(ch)
        elif ch == "(":
            depth += 1; cur.append(ch)
        elif ch == ")":
            if depth == 0:
                out.append("".join(cur)); return out
            depth -= 1; cur.append(ch)
        elif ch == "," and depth == 0:
            out.append("".join(cur)); cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


def _pnl_keys_used(dash_text: str) -> set[str]:
    """Every PNL contract key modDashboard asks for, at any call depth."""
    flat = re.sub(r"_\r?\n\s*", " ", dash_text)          # join continuations

    # helper name -> index of the argument that carries a key
    positions: dict[str, int] = {}
    for m in re.finditer(
            r"^\s*(?:Public |Private )?(?:Function|Sub) (Dash\w+)\(([^)]*)\)",
            flat, re.M):
        params = _split_args(m.group(2))
        for idx, prm in enumerate(params):
            pm = re.search(r"(?:ByVal |ByRef )?(\w+) As ", prm)
            if pm and pm.group(1) in KEY_PARAM_NAMES:
                positions[m.group(1)] = idx
                break

    keys: set[str] = set()
    for name, idx in positions.items():
        for m in re.finditer(re.escape(name) + r"\(", flat):
            args = _split_args(flat[m.end():])
            if idx < len(args):
                lit = re.fullmatch(r'\s*"([^"]*)"\s*', args[idx])
                if lit:
                    keys.add(lit.group(1))
    return keys


def check_pnl_headers(text: str, dash_text: str) -> list[Problem]:
    """PnlLayout is the single declaration of every PNL_Attribution column.

    It pairs a column letter with a CONTRACT KEY (what the Dashboard asks for,
    published as the workbook name Pnl_<key>) and a DISPLAY LABEL (what the desk
    reads in row 4).  The Dashboard no longer searches the header row, so a
    reworded label can no longer break it - but a key it asks for that the table
    does not declare still can, because the name is never published and every
    formula referring to it evaluates to #NAME?.

    That is the drift this check exists to catch.
    """
    out: list[Problem] = []

    layout_src = _procedure_source(text, "PnlLayout")
    if not layout_src:
        return [Problem("L007", "PnlLayout not found - nothing declares the "
                                "PNL_Attribution columns")]

    entries = re.findall(
        r'AddPnlCol spec,\s*(\w+),\s*"([^"]+)",\s*"((?:[^"]|"")*)"', layout_src)
    if not entries:
        return [Problem("L007", "PnlLayout declares no columns")]

    keys = [k for _, k, _ in entries]
    consts = [c for c, _, _ in entries]

    dupe_keys = sorted({k for k in keys if keys.count(k) > 1})
    if dupe_keys:
        out.append(Problem(
            "L007",
            f"PnlLayout declares contract key(s) {dupe_keys} twice; the second "
            f"Names.Add overwrites the first, so one column silently reads the "
            f"other"))

    dupe_consts = sorted({c for c in consts if consts.count(c) > 1})
    if dupe_consts:
        out.append(Problem(
            "L007",
            f"PnlLayout declares column constant(s) {dupe_consts} twice"))

    # A key that is not a legal Excel defined-name fragment cannot be published.
    illegal = sorted(k for k in keys
                     if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", k))
    if illegal:
        out.append(Problem(
            "L007",
            f"contract key(s) {illegal} cannot be part of an Excel defined "
            f"name; Names.Add fails and the column is never published"))

    if not dash_text:
        return out

    declared = set(keys)

    required_src = _procedure_source(dash_text, "DashRequiredPnlKeys")
    required = set(re.findall(r'"([^"]+)"', required_src))
    used = _pnl_keys_used(dash_text)

    asked = used | required

    absent = sorted(asked - declared)
    if absent:
        out.append(Problem(
            "L007",
            f"modDashboard asks for PNL column(s) {absent} that PnlLayout does "
            f"not declare, so PublishPnlColumnNames never publishes them and "
            f"the Dashboard build raises 9901 (or spills #NAME?)"))

    # Everything the Dashboard touches must also be in the validated list, so a
    # missing name is reported once at the top of the build instead of as a
    # sheet full of errors.
    unvalidated = sorted(used - required)
    if unvalidated:
        out.append(Problem(
            "L007",
            f"modDashboard uses PNL column(s) {unvalidated} that "
            f"DashRequiredPnlKeys does not list, so DashValidateSourceHeaders "
            f"will not catch them going missing"))

    return out


# --------------------------------------------------------------------------

def main() -> int:
    path = modules.path_of("modPNL")
    if not path.exists():
        print(f"missing {path}", file=sys.stderr)
        return 2
    text = path.read_text()
    by_sheet = load_constants(text)

    problems: list[Problem] = []
    problems += check_columns(by_sheet)
    problems += check_header_arrays(text)
    problems += check_literal_columns(text)
    problems += check_bond_rows(text)
    problems += check_bbg_ranges(text, by_sheet)
    dash_path = modules.path_of("modDashboard")
    dash_text = dash_path.read_text() if dash_path.exists() else ""
    problems += check_pnl_headers(text, dash_text)

    for p in problems:
        print(str(p))

    total = sum(len(v) for v in by_sheet.values())
    print(f"\n{total} column constants across {len(by_sheet)} sheets, "
          f"{len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
