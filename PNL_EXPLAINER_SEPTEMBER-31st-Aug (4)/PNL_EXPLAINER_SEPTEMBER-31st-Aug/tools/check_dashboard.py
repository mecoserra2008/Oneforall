#!/usr/bin/env python3
"""Render every formula the Dashboard writes, and check it.

The Dashboard is ~900 emitted formulas assembled from string fragments.  None of
them can be read for correctness in the VBA, and there is no Excel here to open
the result in - so this evaluates the concatenation, then checks each finished
formula for the faults that actually occur:

  D001  unbalanced brackets
  D002  a name that is never published, so the cell is #NAME?
  D003  a hedge-sheet column that does not exist on that sheet
  D004  a SUMIFS whose criteria do not pair up, so Excel rejects the formula
  D005  a bare column letter with no row - reads as an undefined name
  D006  a formula Excel will not accept (length, empty, dangling operator)
  D007  a formula the renderer could not evaluate at all

    python3 tools/check_dashboard.py            # check
    python3 tools/check_dashboard.py --list     # print every rendered formula
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules            # noqa: E402
import vbaeval            # noqa: E402

# Excel's own limit on a formula is 8192 characters.
MAX_FORMULA = 8192

# Sample bindings for the loop and cursor variables the build procedures use.
# Only the SHAPE of the emitted formula matters here, not which row it lands on.
SAMPLE = {
    "n": 7, "i": 2, "r": 30, "rowNum": 12, "outRow": 12, "slot": 1,
    "startRow": 100, "startCol": 1, "lastRow": 212, "rowCount": 208,
    "topRow": 40, "outCol": 4, "firstDataRow": 5, "lastDataRow": 212,
    "durationRow": 27, "firstMemo": 28, "lastMemo": 32, "firstComponent": 27,
    "lastComponent": 40, "explainedRow": 41, "residualRow": 42,
    "notAttributedRow": 43, "actualRow": 44, "fxHedgeRow": 40,
    "exposureRow": 103, "coverRow": 104, "bondPnlRow": 107,
    "hedgePnlRow": 108, "totalRow": 250, "firstDetail": 200, "lastDetail": 240,
    "dashRow": 12, "swLast": 220, "rankKeyIndex": 1, "top": 60, "firstPart": 60,
    "bucketRow": 12, "labelRow": 12, "col": 60, "keyName": "DashRankKey1",
    "posExpr": "1", "sw": "'Swaps'!", "fu": "'Futures'!",
    "metricHeader": "Unexplained_Residual_PnL",
    "headerName": "Unexplained_Residual_PnL",
    "titleText": "T", "label": "L", "caption": "C", "note": "N",
    "columnTitle": "C", "shapeName": "s", "macroName": "m",
    "assetExpr": "0", "hedgeExpr": "0", "pnlExpr": "0", "expr": "0",
    "amountFormula": "0", "metricFormula": "=1", "numberFormat": "#,##0",
    "isBold": False, "condition": "1=1", "family": "EURIBOR",
    "contractKey": "ISIN", "valueCol": "U", "bucketCols": "E+F",
    "sheetName": "Futures", "keyCol": "A", "linkCol": "J", "pnlCol": "U",
    "frameworks": ["G", "I"], "cols": ["E", "F"],
}

# Excel functions and structural tokens that are not workbook names.
EXCEL = {
    "IF", "IFERROR", "IFNA", "AND", "OR", "NOT", "N", "SUM", "SUMIF", "SUMIFS",
    "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "SUMPRODUCT", "ABS", "MAX", "MIN",
    "AVERAGE", "ROUND", "INDEX", "MATCH", "LARGE", "SMALL", "ROW", "ROWS",
    "COLUMN", "COLUMNS", "TEXT", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER",
    "LOWER", "ISNUMBER", "ISERROR", "ISNA", "ISBLANK", "ISTEXT", "NA", "TODAY",
    "NOW", "CONCATENATE", "SUBSTITUTE", "SEARCH", "FIND", "CHOOSE", "SWITCH",
    "LET", "FILTER", "SORT", "UNIQUE", "TEXTSPLIT", "TEXTJOIN", "SEQUENCE",
    "OFFSET", "INDIRECT", "VLOOKUP", "XLOOKUP", "HLOOKUP", "EXP", "LN", "LOG",
    "SQRT", "POWER", "SIGN", "INT", "MOD", "YEARFRAC", "DATE", "EOMONTH",
    "TRUE", "FALSE", "IFS", "AGGREGATE",
}


def build(paths) -> vbaeval.Vba:
    return vbaeval.load(paths, stubs={
        # Reach the workbook; a plausible constant is all the shape needs.
        "DashHedgeLastRow": lambda *a: 220,
        "DashLastPnlRow": lambda *a: 212,
        "DashPnlCol": lambda *a: 1,
        "DashKeyCol": lambda *a: 1,
        "DashColLetterOf": lambda *a: "A",
    })


def emitted(text: str, vba: vbaeval.Vba) -> list[tuple[str, str]]:
    """Run every build procedure and collect the formulas it writes.

    Running the whole procedure, rather than evaluating each ".formula ="
    expression on its own, is what makes the locals resolve - most of these
    formulas are assembled into a variable a dozen lines earlier.
    """
    flat = vbaeval.join_continuations(text)
    # Only the block builders.  They call the row helpers themselves, with the
    # real arguments; calling a helper directly would render it against
    # placeholder values and report faults that cannot happen.
    builders = [m.group(1) for m in re.finditer(
        r"^\s*Private Sub (DashBuild\w+)", flat, re.M)]

    for name in builders:
        params, body, _ = vba.procs[name]
        env = {p: SAMPLE.get(p, 1) for p in params}
        env.update({k: v for k, v in SAMPLE.items() if k not in env})
        try:
            vba.run(name, body, env)
        except Exception as exc:                              # noqa: BLE001
            vba.capture_errors.append((name, f"{type(exc).__name__}: {exc}"))
    return vba.captured


def hedge_columns(pnl_text: str) -> dict[str, set[str]]:
    """Declared column letters per hedge sheet, from the letter constants."""
    def letters(prefix: str) -> set[str]:
        return {m.group(1) for m in re.finditer(
            r'Private Const ' + prefix + r'\w+ As String = "([A-Z]{1,3})"',
            pnl_text)}
    return {"Futures": letters("FCOL_"), "Swaps": letters("WCOL_"),
            "Bonds": letters("BCOL_"), "PNL_Attribution": letters("PCOL_")}


def published_names(pnl_text: str) -> set[str]:
    src = re.search(r"Private Function PnlLayout.*?\nEnd Function",
                    pnl_text, re.S)
    prefix = re.search(r'Private Const PNL_NAME_PREFIX As String = "(\w+)"',
                       pnl_text)
    pre = prefix.group(1) if prefix else "Pnl_"
    return {pre + m.group(1) for m in re.finditer(
        r'AddPnlCol spec,\s*\w+,\s*"([^"]+)"', src.group(0) if src else "")}


def strip_strings(f: str) -> str:
    """Blank out string literals and sheet-qualified prefixes.

    A sheet name inside 'Futures'!$J$5 is not an identifier the workbook has to
    resolve, so leaving it in makes every hedge reference look like a missing
    name.  Checked separately, against the real column constants.
    """
    f = re.sub(r'"(?:[^"]|"")*"', '""', f)
    return re.sub(r"'[^']+'!", "", f)


def let_bindings(f: str) -> set[str]:
    """Names bound by LET(...) in this formula.

    Read from the formula rather than trusted from a naming convention: a LET
    variable can be called anything, and a checker that recognises them by a
    leading underscore both misses the ones without it and waves through a
    genuine typo that happens to have one.
    """
    out: set[str] = set()
    for m in re.finditer(r"\bLET\s*\(", f):
        args = _args_after(f, m.end() - 1)
        if not args:
            continue
        for i in range(0, len(args) - 1, 2):
            name = args[i].strip()
            if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name):
                out.add(name)
    return out


def check_one(proc: str, f: str, names: set[str],
              cols: dict[str, set[str]], extra_names: set[str]) -> list[str]:
    bad = []
    bare = strip_strings(f)
    let_names = let_bindings(f)

    if bare.count("(") != bare.count(")"):
        bad.append(f"D001 unbalanced brackets "
                   f"({bare.count('(')} open, {bare.count(')')} close)")

    if len(f) > MAX_FORMULA:
        bad.append(f"D006 {len(f)} characters; Excel's limit is {MAX_FORMULA}")
    if f.strip() in ("=", ""):
        bad.append("D006 empty formula")
    if re.search(r"[-+*/,(]\s*[)]", bare):
        bad.append("D006 empty argument or dangling operator")
    if ",," in bare.replace(",,", ",,"):
        if re.search(r",\s*,", bare):
            bad.append("D006 empty argument between commas")

    for tok in re.findall(r"(?<![\w.!$:])([A-Za-z_][A-Za-z0-9_.]*)", bare):
        if tok.upper() in EXCEL or tok in extra_names:
            continue
        if tok == vbaeval.OBJ:
            continue
        if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", tok):
            continue
        if tok in let_names:
            continue
        if tok.startswith("Pnl_"):
            if tok not in names:
                bad.append(f"D002 {tok} is not published by PnlLayout")
            continue
        if re.fullmatch(r"[A-Z]{1,3}", tok):
            bad.append(f"D005 bare column letter {tok} with no row")
            continue
        bad.append(f"D002 {tok} is not a published name or an Excel function")

    for m in re.finditer(r"'([^']+)'!\$?([A-Z]{1,3})\$?\d+", f):
        sheet, col = m.group(1), m.group(2)
        if sheet in cols and col not in cols[sheet]:
            bad.append(f"D003 '{sheet}'!{col} is not a declared column "
                       f"on that sheet")

    for m in re.finditer(r"\b(SUMIFS|COUNTIFS|AVERAGEIFS)\s*\(", f):
        args = _args_after(f, m.end() - 1)
        if args is None:
            continue
        n = len(args) - (1 if m.group(1) != "COUNTIFS" else 0)
        if n % 2:
            bad.append(f"D004 {m.group(1)} has {len(args)} arguments; "
                       f"criteria must pair up")
            continue

        # Every range in one SUMIFS must cover the same rows.  Excel answers
        # a mismatch with #VALUE!, and the usual way to cause one is to filter
        # a PNL_Attribution column (one row per bond) by a criterion taken off
        # a hedge sheet (one row per hedge) - two lists that are never the same
        # length.
        first = 0 if m.group(1) != "COUNTIFS" else -1
        ranges = [a for i, a in enumerate(args)
                  if (i == 0 and first == 0) or (i - max(first, 0)) % 2 == 0]
        spans = {_span(a) for a in ranges}
        spans.discard(None)
        if len(spans) > 1:
            bad.append(f"D004 {m.group(1)} mixes ranges of different lengths "
                       f"({', '.join(sorted(str(s) for s in spans))}); "
                       f"Excel returns #VALUE!")
    return bad


def _span(arg: str):
    """A range argument's identity: which list of rows it covers."""
    arg = arg.strip()
    m = re.fullmatch(r"(Pnl_\w+)", arg)
    if m:
        return "PNL_Attribution rows"
    m = re.fullmatch(r"'([^']+)'!\$?[A-Z]{1,3}\$?(\d+):\$?[A-Z]{1,3}\$?(\d+)",
                     arg)
    if m:
        return f"{m.group(1)} rows {m.group(2)}-{m.group(3)}"
    return None


def _args_after(f: str, open_paren: int):
    depth, in_str, start, out = 0, False, open_paren + 1, []
    i = open_paren
    while i < len(f):
        ch = f[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(f) and f[i + 1] == '"':
                    i += 2; continue
                in_str = False
        elif ch == '"':
            in_str = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                out.append(f[start:i]); return out
        elif ch == "," and depth == 1:
            out.append(f[start:i]); start = i + 1
        i += 1
    return None


def main() -> int:
    dash_path = modules.path_of("modDashboard")
    pnl_path = modules.path_of("modPNL")
    dash_text = dash_path.read_text(encoding="latin-1")
    pnl_text = pnl_path.read_text(encoding="latin-1")

    vba = build({"dash": dash_path})
    names = published_names(pnl_text)
    cols = hedge_columns(pnl_text)

    # Names the Dashboard defines for itself (the hidden ranking stage, the
    # chart sources) are legitimate references even though PnlLayout has never
    # heard of them.
    extra = set(re.findall(r'DashAddOrReplaceName\s*_?\s*"([^"]+)"', dash_text))
    extra |= {"DashRankKey" + str(i) for i in range(1, 9)}
    extra |= set(re.findall(r'"(DashChart\w+)"', dash_text))

    rows = emitted(dash_text, vba)

    if "--list" in sys.argv:
        for proc, f in rows:
            print(f"{proc}\n  {f}\n")
        return 0

    problems = 0
    partial = 0
    seen: set[tuple[str, str]] = set()
    for proc, f in rows:
        if vbaeval.OBJ in f:
            partial += 1
            f = f.replace(vbaeval.OBJ, "1")
        for msg in check_one(proc, f, names, cols, extra):
            key = (proc, msg)
            if key in seen:
                continue
            seen.add(key)
            problems += 1
            print(f"{msg}\n       in {proc}\n       {f[:230]}")

    for proc, err in vba.capture_errors:
        problems += 1
        print(f"D007 {proc}: a formula could not be rendered\n       {err}")

    note = (f" ({partial} contain a value read from a cell, checked for "
            f"structure only)" if partial else "")
    print(f"\n{len(rows)} Dashboard formula(s) checked{note}, "
          f"{problems} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
