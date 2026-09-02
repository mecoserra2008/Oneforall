#!/usr/bin/env python3
"""Render every PNL_Attribution formula exactly as WritePNLRow builds it.

Hand-transcribing formulas into documentation produces a document that is right
on the day it is written and wrong from the next commit onwards, without anybody
noticing - which is worse than having none, because people trust it.  This reads
the VBA and evaluates the string concatenation the way VBA would, so the record
is generated from the same source the workbook is.

Row numbers are arbitrary (PNL row 5, Bonds row 6); only the substitution
pattern matters.

    python3 tools/dump_formulas.py            # markdown table
    python3 tools/dump_formulas.py --check    # non-zero if anything is unrenderable
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402

PNL_ROW = "5"
BOND_ROW = "6"


def vba_str(expr: str, env: dict[str, str]) -> str:
    """Evaluate a VBA expression of string literals, & and known identifiers."""
    out: list[str] = []
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            j, buf = i + 1, []
            while j < len(expr):
                if expr[j] == '"':
                    if j + 1 < len(expr) and expr[j + 1] == '"':
                        buf.append('"'); j += 2; continue
                    break
                buf.append(expr[j]); j += 1
            out.append("".join(buf)); i = j + 1
        elif ch in " &":
            i += 1
        else:
            m = re.match(r"[A-Za-z_][A-Za-z0-9_]*", expr[i:])
            if not m:
                raise ValueError(expr[i:i + 40])
            name = m.group(0)
            i += len(name)
            if i < len(expr) and expr[i] == "(":          # a helper call
                depth, k = 0, i
                while k < len(expr):
                    if expr[k] == "(":
                        depth += 1
                    elif expr[k] == ")":
                        depth -= 1
                        if depth == 0:
                            break
                    k += 1
                arg = expr[i + 1:k]
                i = k + 1
                out.append(call(name, arg, env))
            else:
                if name not in env:
                    raise KeyError(name)
                out.append(env[name])
    return "".join(out)


def call(name: str, arg: str, env: dict[str, str]) -> str:
    """The few helpers WritePNLRow inlines whose output is a fragment."""
    if name in ("Chr", "Chr$"):
        return chr(int(arg))
    if name in ("XlBlank", "XlBlank$"):
        return '""'
    if name in ("XlText", "XlText$"):
        inner = vba_str(arg, env)
        return '"' + inner.replace('"', '""') + '"'
    if name in ("Replace", "Replace$"):
        a = split_args(arg)
        return vba_str(a[0], env).replace(vba_str(a[1], env), vba_str(a[2], env))

    args = split_args(arg)

    if name == "FutDv01":
        return f"({env['PCOL_FUTURES_RTJ_DV01']}{PNL_ROW}"\
               f"+{env['PCOL_FUTURES_RT_DV01']}{PNL_ROW})"

    if name == "YieldOnlyPnLFml":
        d = f"{env['PCOL_BOND_DV01_OPENING']}{PNL_ROW}"
        y = f"{env['PCOL_DELTA_Y_BP']}{PNL_ROW}"
        return f'IF(AND(ISNUMBER({d}),ISNUMBER({y})),-{d}*{y},"")'

    if name == "FutClassCrit":
        cls = vba_str(args[2], env)
        col = env["FCOL_HEDGE_CLASS"]
        return (f",{env['Fu']}${col}${env['DATA_ROW']}"
                f":${col}${env['futLast']},\"{cls}\"")

    if name == "ChainSumFml":
        cols = [env[c.strip()] for c in
                split_args(re.match(r"\s*Array\((.*)\)\s*$", args[1], re.S).group(1))]
        guard = ",".join(f"ISNUMBER({c}{PNL_ROW})" for c in cols)
        body = "+".join(f"{c}{PNL_ROW}" for c in cols)
        return f'IF(AND({guard}),{body},"")'

    if name == "DV01BlendFml":
        fut = vba_str(args[1], env)
        swp = vba_str(args[2], env)
        f = call("FutDv01", "p", env)
        s = f"{env['PCOL_PLAINSWAP_DV01']}{PNL_ROW}"
        return (f'LET(_f,ABS{f},_s,ABS({s}),_t,_f+_s,'
                f'IF(_t=0,"",IFERROR(_f/_t*({fut})+_s/_t*({swp}),"")))')

    # Anything else: find the builder's own definition and evaluate its body
    # with the arguments bound.  The modEconFormulas builders are all one
    # assignment of concatenated literals and parameters, so this renders them
    # rather than leaving the reader with a function name to go and look up.
    found = builder_body(name)
    if found is not None:
        sig, body = found
        bound = dict(env)
        for prm in sig:
            bound.setdefault(prm, "False")            # optional flags default off
        for prm, val in zip(sig, args):
            try:
                bound[prm] = vba_str(val, env)
            except (KeyError, ValueError):
                bound[prm] = val.strip()
        return replay(name, body, bound)

    raise KeyError(f"{name}(")


_BUILDERS: dict[str, tuple[list[str], str] | None] = {}


def builder_body(name: str):
    """(parameter names, right-hand side) for a single-assignment builder."""
    if name in _BUILDERS:
        return _BUILDERS[name]

    _BUILDERS[name] = None
    for mod in ("modEconFormulas", "modPNL"):
        try:
            text = modules.path_of(mod).read_text(encoding="latin-1")
        except SystemExit:
            continue
        m = re.search(r"^(?:Public |Private )?Function " + re.escape(name)
                      + r"\((.*?)\) As String\r?\n(.*?)\nEnd Function",
                      re.sub(r"_\r?\n\s*", "", text), re.S | re.M)
        if not m:
            continue
        params = [re.sub(r"^(?:ByVal |ByRef )?(\w+).*$", r"\1", a.strip())
                  for a in split_args(m.group(1)) if a.strip()]
        _BUILDERS[name] = (params, m.group(2))
        break
    return _BUILDERS[name]


_IN_PROGRESS: set[str] = set()


def replay(name: str, body: str, bound: dict[str, str]) -> str:
    """Run a builder's body far enough to get its return value.

    Enough VBA for the modEconFormulas builders and no more: assignments, and
    If/Else on a parameter that is already bound.  DiffFormula needs the branch
    because its scaled-by-100 variant is a different formula, and every call
    site here takes the default.
    """
    if name in _IN_PROGRESS:
        raise ValueError(f"{name} is recursive")
    _IN_PROGRESS.add(name)
    try:
        return _replay(name, body, bound)
    finally:
        _IN_PROGRESS.discard(name)


def _replay(name: str, body: str, bound: dict[str, str]) -> str:
    skipping = False
    depth_taken: list[bool] = []

    for raw in body.split("\n"):
        line = raw.strip()
        if not line or line.startswith("'"):
            continue

        m = re.match(r"If (\w+) Then$", line)
        if m:
            taken = str(bound.get(m.group(1), "")).lower() in ("true", "-1")
            depth_taken.append(taken)
            skipping = not taken
            continue
        if line == "Else":
            if depth_taken:
                skipping = depth_taken[-1]
            continue
        if line == "End If":
            if depth_taken:
                depth_taken.pop()
            skipping = False
            continue
        if skipping or line.startswith(("Dim ", "Exit ")):
            continue

        m = re.match(r"(\w+)\s*=\s*(.+)$", line)
        if not m:
            continue
        try:
            bound[m.group(1)] = vba_str(m.group(2), bound)
        except (KeyError, ValueError):
            pass

    if name not in bound:
        raise KeyError(f"{name} never assigned")
    return bound[name]


def split_args(s: str) -> list[str]:
    """Top-level comma split, respecting nesting and string literals."""
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
            depth -= 1; cur.append(ch)
        elif ch == "," and depth == 0:
            out.append("".join(cur)); cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


def load_env(text: str) -> dict[str, str]:
    env: dict[str, str] = {}
    for m in re.finditer(
            r'Private Const (\w+)\s+As String = _?\s*\r?\n?\s*"((?:[^"]|"")*)"', text):
        env[m.group(1)] = m.group(2).replace('""', '"')
    for m in re.finditer(r'Private Const (\w+)\s+As String = "((?:[^"]|"")*)"', text):
        env[m.group(1)] = m.group(2).replace('""', '"')
    env.update({
        "p": PNL_ROW,
        "b2": BOND_ROW,
        "B": "'Bonds'!",
        "c": "'Config'!",
        "Fu": "'Futures'!",
        "sw": "'Swaps'!",
        "futLast": "999",
        "swLast": "999",
        "DATA_ROW": "5",
    })
    return env


def layout(text: str) -> list[tuple[str, str, str, str]]:
    src = re.search(r"Private Function PnlLayout.*?\nEnd Function", text, re.S).group(0)
    rows = []
    consts = dict(re.findall(
        r'Private Const (PCOL_\w+) As String = "([A-Z]{1,3})"', text))
    for m in re.finditer(
            r'AddPnlCol spec,\s*(\w+),\s*"([^"]+)",\s*"((?:[^"]|"")*)"', src):
        rows.append((consts.get(m.group(1), "?"), m.group(1),
                     m.group(2), m.group(3).replace('""', '"')))
    return rows


def locals_of(body: str, env: dict[str, str]) -> dict[str, str]:
    """Local String variables the procedure builds and then interpolates.

    WritePNLRow assembles a few shared fragments - the two framework chains, the
    spread-over-OIS expression, the sheet prefixes - before using them in
    several formulas.  Resolving them here is what makes those formulas render
    instead of showing up as a bare variable name.
    """
    resolved = dict(env)
    for _ in range(4):                                    # they reference each other
        for m in re.finditer(r"^\s*(?:Dim )?(\w+)(?: As String)?(?::| =) ?=? ?(.+)$",
                             body, re.M):
            name, rhs = m.group(1), m.group(2).strip()
            lead = re.match(re.escape(name) + r"\s*=\s*(.+)$", rhs)
            if lead:
                rhs = lead.group(1).strip()
            if name in resolved:
                continue
            if not (rhs.startswith(('"', "'")) or "&" in rhs
                    or re.match(r"\w+\(", rhs)):
                continue
            try:
                resolved[name] = vba_str(rhs, resolved)
            except (KeyError, ValueError):
                pass
    return resolved


def formulas(text: str, env: dict[str, str]) -> dict[str, str]:
    """PCOL constant -> rendered formula, for every write inside WritePNLRow."""
    body = re.search(r"Private Sub WritePNLRow\b.*?\nEnd Sub", text, re.S).group(0)
    body = re.sub(r"_\r?\n\s*", "", body)                 # join continuations
    env = locals_of(body, env)
    out: dict[str, str] = {}
    for m in re.finditer(
            r"ws\.Range\((PCOL_\w+) & p\)\.formula = (.+)", body):
        const, rhs = m.group(1), m.group(2).strip()
        try:
            out[const] = vba_str(rhs, env)
        except (KeyError, ValueError) as exc:
            fn = re.match(r"(\w+)\(", rhs)
            out[const] = (f"_built by {fn.group(1)}_" if fn
                          else f"_unrendered: {exc}_")
    return out


def main() -> int:
    text = modules.path_of("modPNL").read_text(encoding="latin-1")
    env = load_env(text)
    fml = formulas(text, env)
    rows = layout(text)

    unrendered = [c for c, f in fml.items() if f.startswith("_unrendered")]
    if "--check" in sys.argv:
        for c in unrendered:
            print(f"unrendered: {c} {fml[c]}", file=sys.stderr)
        return 1 if unrendered else 0

    def utf8(s: str) -> str:
        """The .bas files are UTF-8 bytes read as latin-1 to round-trip safely;
        the desk's Portuguese labels have to be decoded to be printed."""
        try:
            return s.encode("latin-1").decode("utf-8")
        except (UnicodeDecodeError, UnicodeEncodeError):
            return s

    print("| Col | Key | Sheet label | Formula |")
    print("|---|---|---|---|")
    for letter, const, key, label in rows:
        f = fml.get(const, "_(not written by WritePNLRow)_")
        f = f.replace("|", "\\|")
        if not f.startswith("_"):
            f = f"`{f}`"
        print(f"| {letter} | `{key}` | {utf8(label)} | {f} |")
    return 0


if __name__ == "__main__":
    sys.exit(main())
