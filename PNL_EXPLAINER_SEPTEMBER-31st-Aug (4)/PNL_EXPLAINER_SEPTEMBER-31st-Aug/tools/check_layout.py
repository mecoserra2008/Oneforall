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
  L010  a constant does not address the column the REAL sheet has there
  L011  PNL_LAST_COL is not the last column PnlLayout declares
  L012  a write inside WritePNLRow does not target a PCOL_ constant
  L013  a cross-sheet reference names its column by a literal, or by a
        constant from the wrong sheet's family

L010 is the one that compares the code to the workbook rather than to itself.
Everything above it can pass while an entire block sits one column off the sheet
it addresses, which is what had happened - see docs/analysis/real_headers.tsv.

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

        # Sorted by COLUMN, not by the order the constants happen to be
        # declared in.  The gap walk below compares each column to the previous
        # one, so declaration order was silently load-bearing: reordering two
        # declarations without touching a single letter made L003 report gaps
        # that are not there, and could equally hide one that is.
        for name, letter, num in sorted(entries, key=lambda e: col_num(e[1])):
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

    # The row-4 header is whatever the desk wants it to say, so it is matched as
    # "the rest of the line" rather than as a string literal.  It used to be a
    # literal in every case; twelve of them are now PtVariacao() & "..." because
    # accented text is built from code points, and a literal-only pattern
    # silently stopped seeing those twelve columns - reporting them as
    # undeclared while they were declared all along.
    #
    # Only the constant and the contract key are used below; the header text is
    # captured for readability and deliberately unchecked.
    entries = re.findall(
        r'AddPnlCol spec,\s*(\w+),\s*"([^"]+)",\s*(.+?)\s*$',
        layout_src, re.M)
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
# L010  a constant does not address the column the real sheet has there
# --------------------------------------------------------------------------
#
# Every other rule here checks the constants against EACH OTHER.  They can all
# pass while the whole block sits one column to the right of the sheet it
# addresses - which is exactly what had happened on three sheets at once, for
# an unknown length of time, because nothing compared the code to the workbook.
#
# docs/analysis/real_headers.tsv is that comparison: the live header row, in
# order.  A constant is correct when the header at its position is the field the
# constant names.
#
# The two vocabularies do not always agree, and that is legitimate - the desk
# calls a column "Variação I spread bp" where the code calls it Delta_i_bp, and
# the OPICS query owns the Bonds A:K labels outright.  ALIASES records every
# such pair explicitly, so a rename is a visible edit here rather than a silent
# hole in the check.

HEADERS_TSV = ROOT / "docs" / "analysis" / "real_headers.tsv"

# constant name -> the header text the live sheet actually carries
ALIASES = {
    # Bonds A:K belong to the OPICS query, which names them its own way.
    "BCOL_CCY": "Currency",
    "BCOL_COUPON": "Coupon_Dec",
    "BCOL_COUPON_FREQ": "CouponFreq",
    "BCOL_MATURITY": "MaturityDate",
    # Column 10 is HEADED "Product" and HOLDS the portfolio.  Confirmed with
    # the desk.  It is not a mistake in the constant and it must not be
    # "corrected" by pointing BCOL_PORTFOLIO somewhere else - there is no
    # column headed Portfolio anywhere on Bonds.  The header comes from the
    # OPICS refresh that owns Bonds A:L, which is outside this module, so
    # renaming it is a change to that query and not to any code here.
    "BCOL_PORTFOLIO": "Product",
    "BCOL_BOOKVAL": "BookVal_EUR",
    # PNL_Attribution is read by the desk in Portuguese.
    "PCOL_BOND_DV01_CURRENT": "Bond BPVs",
    "PCOL_BOND_DV01_CREDIT_SPREAD": "Bond DV01 Credit spread",
    "PCOL_HEDGE_DV01": "Hedge (BPVs)",
    "PCOL_PLAINSWAP_DV01": "PlainSwap (BPVs)",
    "PCOL_FUTURES_RTJ_DV01": "Futures RTJ (BPVs)",
    "PCOL_FUTURES_RT_DV01": "Futures RT (BPVs)",
    "PCOL_DELTA_DIRTY_MV_EUR": "Variação Bond (EUR)",
    "PCOL_DELTA_Y_BP": "Variação yield bond (bps)",
    "PCOL_DELTA_R_BP": "Variação yield rf (bps)",
    "PCOL_DELTA_GOV_BP": "Variação yield Gov (bps)",
    "PCOL_DELTA_G_BP": "Variação Govi basis bp",
    "PCOL_DELTA_SWAP_BP": "Variação Swap bp",
    "PCOL_DELTA_Q_BP": "Variação Gov/Swap basis bp",
    "PCOL_DELTA_I_BP": "Variação I spread bp",
    "PCOL_DELTA_Z_BP": "Variação Z spread bp",
    "PCOL_DELTA_G_BP_T": "Variação G spread bp",
    "PCOL_DELTA_ASW_BP": "Variação ASW bp",
    "PCOL_DELTA_OAS_BP": "Variação OAS bp",
    "PCOL_CARRY_ROLLTOPAR": "Carry_Pull_ToPar",
    "PCOL_CARRY_FUNDING": "Funding_Carry_Memo",
    "PCOL_PNL_FUTURES": "Futures_Gov_Model_PnL",
    "PCOL_PNL_SWAP": "Swap_Curve_Model_PnL",
    "PCOL_TOTAL_HEDGE": "Hedge_Curve_Model_PnL",
    "PCOL_BASIS_PNL": "Hedge_Model_Residual_PnL",
    "PCOL_OFFICIAL_PNL": "Official_Total_PnL",
    "PCOL_RESIDUAL": "Unexplained_Residual_PnL",
    "PCOL_RESIDUAL_AX": "Unexplained_Residual_Pct",
    "PCOL_FUT_PNL_RAW": "Actual_Futures_PnL",
    "PCOL_SWAP_PLAIN_PNL_RAW": "Actual_PlainSwap_PnL",
    "PCOL_SWAP_SYNTHETIC_PNL_RAW": "Actual_SyntheticSwap_PnL",
    "PCOL_SWAP_ALL_PNL_RAW": "AllSwapRows_PnL_Diagnostic",
    "PCOL_ACTUAL_FUTURES_PNL_RAW": "Actual_Futures_PnL_Check",
    "PCOL_FUTURES_GOV_MODEL_PNL": "Futures_Gov_Model_PnL_Check",
    "PCOL_FUTURES_BASIS_PNL": "Futures_Model_Residual_PnL",
    "PCOL_ACTUAL_SWAP_PNL_RAW": "Actual_PlainSwap_PnL_Check",
    "PCOL_SWAP_CURVE_MODEL_PNL": "Swap_Curve_Model_PnL_Check",
    "PCOL_SWAP_BASIS_PNL": "Swap_Model_Residual_PnL",
    "PCOL_ACTUAL_HEDGE_PNL_RAW": "Actual_Hedge_PnL",
    "PCOL_HEDGE_BASIS_PNL": "Hedge_Model_Residual_PnL_Check",
    "PCOL_THEORETICALHEDGE_DV01": "Synthetic_Alternative_DV01",
    "PCOL_TOTAL_EXPLAINED": "Total_Model_Explained",
    "PCOL_FUT_MATCH_COUNT": "Futures_Match_Count",
    "PCOL_SWAP_PLAIN_MATCH_COUNT": "PlainSwap_Match_Count",
    "PCOL_SWAP_SYNTHETIC_MATCH_COUNT": "SyntheticSwap_Match_Count",
    # Swaps
    "WCOL_DEALID": "DealID",
    "WCOL_DV01_BBG": "DV01 BBG",
    "WCOL_PV_T0_MODEL": "PV_T0 da anuidade",
    "WCOL_PV_TM1_MODEL": "PV_T-1 da anuidade",
    "WCOL_SWAP_DV01_EUR": "Swap_DV01 Calculated In house",
    "WCOL_CPTY": "Cpty",
    "WCOL_MAP_CCY": "Map_CCY",
    "WCOL_SWAP_DV01_EUR_THEO": "Swap_DV01_EUR_Theoretical",
    # The BA:BI block duplicates P:X header-for-header, so the constants carry a
    # DUP prefix the headers carry as a suffix.  Both say the same thing.
    "WCOL_DUP_BQL_NPV_DIRECT_T0": "BQL_NPV_Direct_T0_DUP",
    "WCOL_DUP_BQL_NPV_FIXED_T0": "BQL_NPV_Fixed_T0_DUP",
    "WCOL_DUP_BQL_NPV_FLOAT_T0": "BQL_NPV_Float_T0_DUP",
    "WCOL_DUP_BQL_NPV_TOTAL_T0": "BQL_NPV_Total_T0_DUP",
    "WCOL_DUP_BQL_NPV_DIRECT_TM1": "BQL_NPV_Direct_T-1_DUP",
    "WCOL_DUP_BQL_NPV_FIXED_TM1": "BQL_NPV_Fixed_T-1_DUP",
    "WCOL_DUP_BQL_NPV_FLOAT_TM1": "BQL_NPV_Float_T-1_DUP",
    "WCOL_DUP_BQL_NPV_TOTAL_TM1": "BQL_NPV_Total_T-1_DUP",
    "WCOL_DUP_BQL_SWAP_PNL": "BQL_Swap_PnL_DUP",
    # Futures
    "FCOL_CTD_CF": "CTD_CF",
    "FCOL_FUT_VAL_PT": "FUT_VAL_PT",
    "FCOL_COVERAGEINFO_D": "CoverageInfo_D",
}


def _norm(s: str) -> str:
    """Compare a constant's name to a header, ignoring the cosmetic differences.

    The code cannot put a hyphen in an identifier, so the prior snapshot is
    `_TM1` in a constant and `_T-1` in a header.  They are the same column and
    always have been; folding them here keeps that out of the ALIASES table,
    which is for genuine vocabulary differences rather than for a spelling VBA
    forces on us.
    """
    t = re.sub(r"[^a-z0-9]", "", s.lower())
    return re.sub(r"tm1", "t1", t)


def load_real_headers() -> dict[str, dict[int, str]]:
    """sheet -> {1-based position: header text}, from the fixture."""
    out: dict[str, dict[int, str]] = defaultdict(dict)
    if not HEADERS_TSV.exists():
        return out
    for line in HEADERS_TSV.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) != 3:
            continue
        sheet, pos, header = parts
        out[sheet.strip()][int(pos)] = header.strip()
    return out


def check_pnl_row_targets(text: str) -> list[Problem]:
    """L012  every write inside WritePNLRow must target a PCOL_ constant.

    WritePNLRow builds one row of PNL_Attribution.  Every cell it writes belongs
    to that sheet, so every target is a PCOL_ constant - and when it is not, the
    write lands somewhere else entirely.

    It did.  Removing the Bond_DV01_Opening column replaced every REFERENCE to
    it with `dvOpen`, the string "'Bonds'!CL<row>", and left behind the line
    that used to WRITE it:

        ws.Range(dvOpen).formula = "=" & B & BCOL_DV01_OPENING_EUR & b2

    which aims a write at the Bonds sheet through the PNL sheet's Range object.
    Bonds!CL is the multiplier on every attribution leg.

    vba_lint cannot see this - ws.Range(<variable>) is perfectly legal VBA and
    the variable is defined - so the check has to know what this procedure is
    FOR, which is what makes it a layout rule rather than a lint rule.
    """
    out: list[Problem] = []

    body = _procedure_source(text, "WritePNLRow")
    if not body:
        return out

    # Comment lines are dropped first.  Without that the rule reports the
    # comment left where the offending write used to be, which explains what
    # went wrong by quoting it - and quoting it was enough to trip the check.
    code = "\n".join(
        line for line in body.splitlines()
        if not line.lstrip().startswith("'"))

    for m in re.finditer(r"ws\.Range\(\s*([A-Za-z_]\w*)", code):
        target = m.group(1)
        if target.startswith("PCOL_"):
            continue
        line = code[:m.start()].count("\n") + 1
        out.append(Problem(
            "L012",
            f"WritePNLRow writes to ws.Range({target} ...) about {line} line(s) "
            f"in - every cell this procedure writes belongs to "
            f"PNL_Attribution, so the target must be a PCOL_ constant"))
    return out


def check_pnl_extent(text: str) -> list[Problem]:
    """L011  PNL_LAST_COL must be the last column PnlLayout declares.

    It bounds the range every recalculation calculates.  Too small and the tail
    of the sheet is never recalculated - the cells hold the previous run's
    numbers and look entirely plausible.  Too large and blank columns are.

    There were two Debug.Asserts on this, and they had been failing: the
    constant said CI while the layout ended at CH.  An assertion only fires for
    someone running the macro in the VBE with the Immediate window open.
    """
    out: list[Problem] = []

    m = re.search(r'Private Const PNL_LAST_COL As String\s*=\s*"([A-Z]{1,3})"', text)
    if not m:
        return out
    declared = m.group(1)

    cols = re.findall(r"^\s*AddPnlCol spec, (PCOL_[A-Z0-9_]+),", text, re.M)
    if not cols:
        return out

    letters = {}
    for cm in CONST_RE.finditer(text):
        letters[cm.group("name")] = cm.group("letter")

    known = [letters[c] for c in cols if c in letters]
    if not known:
        return out
    last = max(known, key=col_num)

    if col_num(declared) != col_num(last):
        out.append(Problem(
            "L011",
            f'PNL_LAST_COL is "{declared}" (column {col_num(declared)}) but '
            f'PnlLayout\'s last column is "{last}" (column {col_num(last)}); '
            f"the recalculation range and the layout disagree"))
    return out


def check_against_real_headers(by_sheet) -> list[Problem]:
    out: list[Problem] = []
    real = load_real_headers()
    if not real:
        out.append(Problem("L010", f"{HEADERS_TSV} is missing - the column "
                                   f"constants cannot be checked against the "
                                   f"workbook they address"))
        return out

    for sheet, entries in sorted(by_sheet.items()):
        headers = real.get(sheet)
        if not headers:
            continue                       # sheet not covered by the fixture

        for name, letter, _num in sorted(entries, key=lambda e: col_num(e[1])):
            pos = col_num(letter)
            want = ALIASES.get(name, name.split("_", 1)[1])
            got = headers.get(pos)

            if got is None:
                out.append(Problem(
                    "L010", f"{sheet}: {name} addresses {letter} but the sheet "
                            f"has only {len(headers)} columns"))
            elif _norm(got) != _norm(want):
                out.append(Problem(
                    "L010", f"{sheet}!{letter}: {name} expects '{want}' but the "
                            f"sheet has '{got}'"))

        # And the other direction.  Everything above walks the CONSTANTS and
        # asks what the sheet has there, so a real column that no constant
        # names is invisible - which is how Futures Link_Source sat unnamed
        # until something needed to address it.  A column nothing can name is
        # a column nothing can read, write, snapshot or restore.
        claimed = {col_num(letter) for _, letter, _ in entries}
        gaps = DECLARED_GAPS.get(sheet, set())
        for pos in sorted(headers):
            if pos in claimed or pos in gaps:
                continue
            out.append(Problem(
                "L010", f"{sheet}: column {pos} is '{headers[pos]}' on the "
                        f"sheet but no constant names it"))
    return out


# --------------------------------------------------------------------------

# WritePNLRow builds every cross-sheet reference by concatenating a sheet
# prefix onto a column: `B & BCOL_MATURITY & b2`.  Which prefix goes with which
# constant family is the whole contract.
FORMULA_PREFIXES = {
    "B": ("Bonds", "BCOL_"),
    "fu": ("Futures", "FCOL_"),
    "sw": ("Swaps", "WCOL_"),
    "c": ("Config", "CFG_"),
}

# `B & <something>` - the something is either a quoted letter or an identifier.
PREFIX_USE_RE = re.compile(
    r'\b(B|fu|sw|c)\s*&\s*(?:"\$"\s*&\s*)?(?:"(?P<lit>[A-Z]{1,3})"|(?P<name>[A-Za-z_][A-Za-z0-9_]*))'
)


def check_formula_sheet_refs(text: str, by_sheet) -> list[Problem]:
    """A cross-sheet reference must name its column through the RIGHT family's
    constant.

    This is the check that was missing when the live workbook ended up reading
    `Bonds!CM` - one column past the end of a 90-column sheet.  A reference
    built that way is not an error to Excel: CM is a real, empty cell, so it
    reads as 0.  Every attribution leg multiplies by the opening DV01, so the
    whole duration and spread chain reported exactly zero, every ISNUMBER guard
    passed because 0 is a number, and Duration_Identity_Check tied perfectly
    because both of its sides were zero.

    Two ways to get there, both caught here:

      a literal letter    `B & "CM" & b2` - moves the column and the reference
                          stops agreeing, exactly as L008 says for writes
      the wrong family    `B & PCOL_ISIN & b2` - a PNL letter aimed at Bonds.
                          Both constants exist, both spell a real column, and
                          the reference lands on whatever Bonds has there.

    A constant of the right family that is out of the sheet's range is L010's
    job, so it is not repeated here.
    """
    out: list[Problem] = []
    known = {
        name
        for entries in by_sheet.values()
        for name, _letter, _num in entries
    }

    for src_name in ("WritePNLRow", "WritePNLSection"):
        src = _procedure_source(text, src_name)
        if not src:
            continue
        for line in _logical_lines(src):
            if line.lstrip().startswith("'"):
                continue
            for m in PREFIX_USE_RE.finditer(line):
                prefix = m.group(1)
                sheet, family = FORMULA_PREFIXES[prefix]

                if m.group("lit"):
                    out.append(Problem(
                        "L013",
                        f"{src_name} builds a {sheet} reference from the "
                        f"hard-coded column \"{m.group('lit')}\" - use the "
                        f"{family} constant so the reference moves with the "
                        f"column"))
                    continue

                name = m.group("name")
                if name.startswith(family):
                    continue
                if name in known:
                    out.append(Problem(
                        "L013",
                        f"{src_name} aims {name} at {sheet} - {sheet} columns "
                        f"are named by {family} constants, so this lands on "
                        f"whatever {sheet} happens to have at that letter"))
    return out


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
    problems += check_pnl_extent(text)
    problems += check_pnl_row_targets(text)
    problems += check_against_real_headers(by_sheet)
    problems += check_formula_sheet_refs(text, by_sheet)

    for p in problems:
        print(str(p))

    total = sum(len(v) for v in by_sheet.values())
    print(f"\n{total} column constants across {len(by_sheet)} sheets, "
          f"{len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
