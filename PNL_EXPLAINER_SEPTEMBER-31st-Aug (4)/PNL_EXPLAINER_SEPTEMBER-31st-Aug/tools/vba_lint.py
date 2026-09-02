#!/usr/bin/env python3
"""Static checker for the PNL Explainer VBA project.

Reproduces, outside Excel, the compile-time errors that actually stop this
workbook from running:

  E001  Unbalanced block          (If/End If, For/Next, With/End With, ...)
  E002  Variable not defined      (Option Explicit)
  E003  Sub or Function not defined
  E004  Ambiguous name detected   (same Public name in two modules)
  E005  Duplicate declaration in current scope
  E006  Too many line continuations (VBA allows 25 per logical line)
  E007  Logical line too long     (VBA hard limit 1023 chars)
  E008  Procedure never terminated / stray End
  E009  Assignment to a Function's own name missing (function has no return)
  E010  Identifier shadows a VBA keyword (Name, Base, Error, Line, ...)
  E011  Module-level declaration after the first procedure
  E012  Wrong number of arguments in a call
  W001  Private procedure is never called (dead code)

Usage:
    python3 tools/vba_lint.py [src/*.bas ...]
Exit code is non-zero when any error is reported.
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from vba_builtins import KEYWORDS, is_builtin           # noqa: E402
from vbaparse import (                                  # noqa: E402
    CONST_RE, Module, Procedure, code_of, parse_project, _names_from_decl,
)

MAX_CONTINUATIONS = 25
MAX_LINE_CHARS = 1023


class Finding:
    def __init__(self, code: str, module: str, line: int, message: str):
        self.code = code
        self.module = module
        self.line = line
        self.message = message

    def __str__(self) -> str:
        return f"{self.module}:{self.line}: {self.code} {self.message}"


# --------------------------------------------------------------------------
# E001 / E008  block structure
# --------------------------------------------------------------------------

_OPENERS = [
    (re.compile(r"^\s*(?:\w+:\s*)?If\b.*\bThen\s*$", re.I), "If"),
    (re.compile(r"^\s*(?:Public|Private|Global)?\s*(?:Select)\s+Case\b", re.I), "Select"),
    (re.compile(r"^\s*With\b", re.I), "With"),
    (re.compile(r"^\s*(?:Public|Private)?\s*(?:Type|Enum)\s+\w+", re.I), "TypeEnum"),
]

_RE_FOR = re.compile(r"^\s*For\b", re.I)
_RE_NEXT = re.compile(r"^\s*Next\b", re.I)
_RE_DO = re.compile(r"^\s*Do\b", re.I)
_RE_LOOP = re.compile(r"^\s*Loop\b", re.I)
_RE_WHILE = re.compile(r"^\s*While\b", re.I)
_RE_WEND = re.compile(r"^\s*Wend\b", re.I)
_RE_ENDIF = re.compile(r"^\s*End\s+If\b|^\s*EndIf\b", re.I)
_RE_ENDSELECT = re.compile(r"^\s*End\s+Select\b", re.I)
_RE_ENDWITH = re.compile(r"^\s*End\s+With\b", re.I)
_RE_ENDTYPE = re.compile(r"^\s*End\s+(?:Type|Enum)\b", re.I)
_RE_ELSE = re.compile(r"^\s*(?:Else|ElseIf)\b", re.I)


def check_blocks(mod: Module) -> list[Finding]:
    out: list[Finding] = []
    stack: list[tuple[str, int]] = []

    for ll in mod.logical:
        t = ll.code.strip()
        if not t:
            continue

        # a single-line If has code after Then
        if re.match(r"^\s*If\b", t, re.I) and re.search(r"\bThen\b", t, re.I):
            after = re.split(r"\bThen\b", t, maxsplit=1, flags=re.I)[1].strip()
            if after == "":
                stack.append(("If", ll.start))
            continue
        if _RE_ELSE.match(t):
            continue
        if _RE_ENDIF.match(t):
            if stack and stack[-1][0] == "If":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'End If' without a matching 'If'"))
            continue

        if re.match(r"^\s*Select\s+Case\b", t, re.I):
            stack.append(("Select", ll.start))
            continue
        if _RE_ENDSELECT.match(t):
            if stack and stack[-1][0] == "Select":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'End Select' without a matching 'Select Case'"))
            continue

        if re.match(r"^\s*With\b", t, re.I):
            stack.append(("With", ll.start))
            continue
        if _RE_ENDWITH.match(t):
            if stack and stack[-1][0] == "With":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'End With' without a matching 'With'"))
            continue

        if _RE_FOR.match(t):
            stack.append(("For", ll.start))
            continue
        if _RE_NEXT.match(t):
            if stack and stack[-1][0] == "For":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'Next' without a matching 'For'"))
            continue

        if _RE_DO.match(t):
            stack.append(("Do", ll.start))
            continue
        if _RE_LOOP.match(t):
            if stack and stack[-1][0] == "Do":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'Loop' without a matching 'Do'"))
            continue

        if _RE_WHILE.match(t):
            stack.append(("While", ll.start))
            continue
        if _RE_WEND.match(t):
            if stack and stack[-1][0] == "While":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'Wend' without a matching 'While'"))
            continue

        if re.match(r"^\s*(?:Public|Private)?\s*(?:Type|Enum)\s+\w+", t, re.I):
            stack.append(("TypeEnum", ll.start))
            continue
        if _RE_ENDTYPE.match(t):
            if stack and stack[-1][0] == "TypeEnum":
                stack.pop()
            else:
                out.append(Finding("E001", mod.name, ll.start,
                                   "'End Type/Enum' without a matching opener"))
            continue

        # procedure boundaries reset the block stack expectation
        if re.match(r"^\s*End\s+(Sub|Function|Property)\b", t, re.I):
            while stack:
                kind, where = stack.pop()
                out.append(Finding(
                    "E001", mod.name, where,
                    f"'{kind}' block never closed before End of procedure "
                    f"at line {ll.start}"))
            continue

    for kind, where in stack:
        out.append(Finding("E001", mod.name, where,
                           f"'{kind}' block never closed"))
    return out


def check_proc_termination(mod: Module) -> list[Finding]:
    out: list[Finding] = []
    for p in mod.procedures:
        last = p.lines[-1].text.strip() if p.lines else ""
        if not re.match(r"^\s*End\s+(Sub|Function|Property)\b", last, re.I):
            out.append(Finding("E008", mod.name, p.start,
                               f"{p.kind} {p.name} is never terminated "
                               f"with 'End {p.kind.split()[0]}'"))
    return out


# --------------------------------------------------------------------------
# identifier scanning
# --------------------------------------------------------------------------

_IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

_SKIP_AFTER = re.compile(r"\b(GoTo|GoSub|Resume|Lib|Alias|Declare)\s+$", re.I)


def _local_declarations(p: Procedure) -> dict[str, int]:
    names: dict[str, int] = {}
    for name in p.params:
        names[name.lower()] = p.start
    for ll in p.lines:
        t = ll.code.strip()
        m = CONST_RE.match(t)
        if m:
            for n in _names_from_decl(m.group("rest")):
                names.setdefault(n.lower(), ll.start)
            continue
        m = re.match(r"^(?:Dim|Static|ReDim)\s+(?:Preserve\s+)?(.+)$", t, re.I)
        if m:
            for n in _names_from_decl(m.group(1)):
                names.setdefault(n.lower(), ll.start)
            continue
        # For Each x In ... / For i = ...  (declared elsewhere normally, but
        # tolerate the pattern so a missing Dim is reported once, not twice)
        m = re.match(r"^For\s+Each\s+([A-Za-z_]\w*)", t, re.I)
        if m:
            names.setdefault(m.group(1).lower(), ll.start)
    return names


def _bare_identifiers(code: str):
    """Yield (name, position, followed_by_paren, preceded_by_dot)."""
    i = 0
    n = len(code)
    while i < n:
        m = _IDENT.search(code, i)
        if not m:
            return
        name = m.group(0)
        start, end = m.span()

        prev = code[:start].rstrip()
        preceded_by_dot = prev.endswith(".")
        # named argument   foo bar:=1
        after = code[end:].lstrip()
        is_named_arg = after.startswith(":=")
        follows_as = re.search(r"\bAs\s+$", code[:start], re.I) is not None
        follows_new = re.search(r"\bNew\s+$", code[:start], re.I) is not None
        follows_jump = _SKIP_AFTER.search(code[:start]) is not None
        follows_hash = prev.endswith("#") or prev.endswith("$") or prev.endswith("!")
        # &HFFFF / &O777 are numeric LITERALS.  Without this the scanner reads
        # the "HFFFF" as a bare identifier and reports E002 on a constant.
        follows_amp_radix = bool(re.search(r"&\s*$", code[:start])) and \
            name[:1] in "HhOo" and len(name) > 1 and \
            all(c in "0123456789abcdefABCDEF" for c in name[1:])
        if follows_amp_radix:
            i = end
            continue

        yield {
            "name": name,
            "start": start,
            "paren": after.startswith("("),
            "dot": preceded_by_dot,
            "named_arg": is_named_arg,
            "as_type": follows_as,
            "new_type": follows_new,
            "jump": follows_jump,
            "typed": follows_hash,
        }
        i = end


_LABEL_RE = re.compile(r"^\s*([A-Za-z_]\w*)\s*:\s*$")


def _labels_of(p: Procedure) -> set[str]:
    out = set()
    for ll in p.lines:
        m = _LABEL_RE.match(ll.code)
        if m:
            out.add(m.group(1).lower())
    return out


def _declaration_line(t: str) -> bool:
    return bool(re.match(
        r"^\s*(Dim|Const|Static|Private|Public|Global|Declare|Option|Type|Enum|"
        r"Sub|Function|Property|End|Attribute)\b", t, re.I))


def check_identifiers(mods: list[Module]) -> list[Finding]:
    out: list[Finding] = []

    proc_names: dict[str, list[Procedure]] = defaultdict(list)
    for mod in mods:
        for p in mod.procedures:
            proc_names[p.name.lower()].append(p)

    global_public: set[str] = set()
    for mod in mods:
        global_public |= mod.public_names

    for mod in mods:
        if not mod.option_explicit:
            out.append(Finding("E002", mod.name, 1,
                               "module has no 'Option Explicit'"))

        module_scope = set(mod.module_vars) | set(mod.module_consts) \
            | set(mod.enum_members) | set(mod.types) | set(mod.declares)

        visible_procs = {n for n in proc_names}          # any module's proc
        # a Private proc in another module is NOT visible; refine:
        visible_procs = set()
        for name, plist in proc_names.items():
            for p in plist:
                if p.module == mod.name or p.is_public:
                    visible_procs.add(name)
                    break

        for p in mod.procedures:
            locals_ = _local_declarations(p)
            labels = _labels_of(p)
            known = (set(locals_) | module_scope | labels
                     | {p.name.lower()} | visible_procs | global_public)

            for ll in p.lines:
                t = ll.code
                stripped = t.strip()
                if not stripped:
                    continue
                if _declaration_line(stripped):
                    continue
                if _LABEL_RE.match(stripped):
                    continue

                for tok in _bare_identifiers(t):
                    name = tok["name"]
                    low = name.lower()
                    if tok["dot"] or tok["named_arg"] or tok["as_type"] \
                            or tok["new_type"] or tok["jump"] or tok["typed"]:
                        continue
                    if is_builtin(name):
                        continue
                    if low in known:
                        continue
                    if tok["paren"]:
                        out.append(Finding(
                            "E003", mod.name, ll.start,
                            f"Sub or Function not defined: '{name}' "
                            f"(in {p.kind} {p.name})"))
                    else:
                        out.append(Finding(
                            "E002", mod.name, ll.start,
                            f"Variable not defined: '{name}' "
                            f"(in {p.kind} {p.name})"))
    return out


# --------------------------------------------------------------------------
# W001  unreferenced private procedures
# --------------------------------------------------------------------------

def check_dead_code(mods: list[Module]) -> list[Finding]:
    """Private procedures nothing calls.

    Reported as a warning, not an error: it is the signal that a refactor left
    a helper behind, which is how a module accumulates 1,000 lines nobody has
    read in two years.  Public procedures are exempt - they are button targets
    and worksheet UDFs, called from outside the source.
    """
    out: list[Finding] = []
    for mod in mods:
        text = "\n".join(ll.code for ll in mod.logical)
        for p in mod.procedures:
            if p.is_public:
                continue
            body_start = p.start
            pat = re.compile(r"(^|[^.\w])" + re.escape(p.name) + r"($|[^\w])",
                             re.I)
            hits = 0
            for other in mod.logical:
                if body_start <= other.start <= p.end:
                    continue          # its own definition and body
                if pat.search(other.code):
                    hits += 1
                    break
            if hits == 0:
                out.append(Finding("W001", mod.name, p.start,
                                   f"{p.kind} {p.name} is never called"))
    return out


# --------------------------------------------------------------------------
# E010  keyword shadowing
# --------------------------------------------------------------------------

# Words VBA/Basic will reject (or silently mis-parse) as a parameter or variable
# name.  These are the ones that read like perfectly ordinary finance
# identifiers, which is exactly why they get used: `base` for a curve's short
# end, `name` for a label, `error` for a residual, `date`/`time`/`line`/`step`.
# The failure is at COMPILE time, so nothing runs at all - and in a headless
# host it surfaces as a hang rather than a message.
SHADOW_TRAPS = {
    "base", "name", "error", "line", "date", "time", "step", "input",
    "output", "print", "write", "close", "open", "get", "put", "seek",
    "len", "left", "right", "mid", "string", "single", "currency", "type",
    "stop", "resume", "return", "select", "set", "read", "width", "spc",
    "tab", "kill", "reset", "lock", "load", "text", "binary", "random",
    "append", "shared", "module", "empty", "null",
}


def check_keyword_shadowing(mod: Module) -> list[Finding]:
    out: list[Finding] = []

    def flag(where: int, kind: str, name: str, ctx: str) -> None:
        out.append(Finding(
            "E010", mod.name, where,
            f"{kind} '{name}' shadows the VBA keyword '{name}'{ctx}"))

    for name, line in mod.module_vars.items():
        if name in SHADOW_TRAPS:
            flag(line, "module variable", name, "")
    for name, line in mod.module_consts.items():
        if name in SHADOW_TRAPS:
            flag(line, "module constant", name, "")

    for p in mod.procedures:
        ctx = f" (in {p.kind} {p.name})"
        for param in p.params:
            if param.lower() in SHADOW_TRAPS:
                flag(p.start, "parameter", param, ctx)
        for name, line in _local_declarations(p).items():
            if name in SHADOW_TRAPS and name not in {
                    x.lower() for x in p.params}:
                flag(line, "local variable", name, ctx)
    return out


# --------------------------------------------------------------------------
# E011  module-level declaration after the first procedure
# --------------------------------------------------------------------------

# VBA has a declarations SECTION, not merely declarations: every module-level
# Const, Dim, Type, Enum and Declare must appear before the first Sub or
# Function.  One placed lower down is a COMPILE error, so nothing in the module
# runs at all - and the VBE reports it at the offending line rather than at the
# code that uses the constant, which is the opposite of where you look.
#
# It is easy to do and invisible on review: appending a new block of code to the
# end of a module is the natural way to add a feature, and a constant that block
# needs travels down with it.  Everything still reads correctly.


def check_declaration_section(mod: Module) -> list[Finding]:
    out: list[Finding] = []
    if not mod.procedures:
        return out

    first = min(p.start for p in mod.procedures)
    first_name = next(p.name for p in mod.procedures if p.start == first)

    inside = set()
    for p in mod.procedures:
        inside.update(range(p.start, p.end + 1))

    for kind, table in (("Const", mod.module_consts),
                        ("variable", mod.module_vars),
                        ("Type", mod.types),
                        ("Declare", mod.declares)):
        for name, line in table.items():
            if line > first and line not in inside:
                out.append(Finding(
                    "E011", mod.name, line,
                    f"module-level {kind} '{name}' is declared after "
                    f"{first_name} (line {first}); VBA requires every "
                    f"module-level declaration to precede the first procedure"))
    return out


# --------------------------------------------------------------------------
# E012  wrong number of arguments
# --------------------------------------------------------------------------
#
# E003 asks whether a called name EXISTS.  It never asked how many arguments
# the call passes, so changing a procedure's signature and missing a call site
# left the project uncompilable with the suite fully green - which is exactly
# how two "Argument not optional" errors reached the VBE.
#
# Both call FORMS are checked, because Subs are normally called without
# parentheses and a `name(` pattern alone would have missed every one of them:
#
#     x = Foo(a, b)        expression form
#     Foo a, b             statement form
#     Call Foo(a, b)       statement form, parenthesised
#
# String literals are already blanked in LogicalLine.code, which matters more
# than it sounds: without it the checker reports ImpliedRepoBloomberg( inside a
# formula-builder string and PnlLayout ( inside a MsgBox caption, neither of
# which is a call at all.

_SIG_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\s*(?:\((?P<params>.*)\))?\s*(?:As\s+\w+)?\s*$",
    re.IGNORECASE)

# Statement-form lines that are not calls.  `Set`, `Let` and `Call` are handled
# separately; the rest simply never start a call.
_NOT_A_CALL = {
    "if", "elseif", "else", "end", "exit", "for", "next", "do", "loop",
    "while", "wend", "select", "case", "with", "on", "resume", "goto",
    "gosub", "return", "dim", "redim", "const", "static", "public", "private",
    "global", "friend", "sub", "function", "property", "type", "enum",
    "declare", "option", "attribute", "erase", "open", "close", "print",
    "write", "input", "line", "put", "get", "seek", "lock", "unlock", "name",
    "kill", "mkdir", "rmdir", "chdir", "chdrive", "randomize", "load",
    "unload", "stop", "error", "raiseevent", "implements", "set", "let",
    "then", "wend", "each", "in", "to", "step", "is", "new",
}


def _split_top(s: str) -> list[str]:
    """Split on commas at bracket depth zero."""
    out, cur, depth = [], "", 0
    for ch in s:
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _arity(params: str | None) -> tuple[int, int]:
    """(minimum, maximum) arguments a signature accepts."""
    if not params or not params.strip():
        return 0, 0
    parts = [p.strip() for p in _split_top(params) if p.strip()]
    required = sum(
        1 for p in parts
        if not re.match(r"^(Optional|ParamArray)\b", p, re.IGNORECASE))
    unbounded = any(re.match(r"^ParamArray\b", p, re.I) for p in parts)
    return required, (10 ** 6 if unbounded else len(parts))


def _count_args(text: str) -> int:
    return len([p for p in _split_top(text)]) if text.strip() else 0


def _statements(code: str) -> list[str]:
    """One logical line split into the statements VBA reads it as.

    `SetIfBlank ws, "A4", "label":  SetIfBlank ws, CFG_T0_DATE, ""` is two
    calls of three arguments, not one call of five - which is what a checker
    that ignores the `:` separator concludes, and it concluded it twenty-one
    times before this existed.

    `:=` is a named argument and never a separator; a trailing `:` on its own
    is a line label.
    """
    out, cur, depth = [], "", 0
    i = 0
    while i < len(code):
        ch = code[i]
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        if ch == ":" and depth == 0 and not code[i + 1:i + 2] == "=":
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    out.append(cur)
    return [s for s in out if s.strip()]


def check_arg_counts(mods: list[Module]) -> list[Finding]:
    out: list[Finding] = []

    arity: dict[str, tuple[int, int, str]] = {}
    for mod in mods:
        for ll in mod.logical:
            m = _SIG_RE.match(ll.text)
            if m:
                lo, hi = _arity(m.group("params"))
                arity[m.group("name").lower()] = (lo, hi, mod.name)

    def report(mod_name: str, line: int, name: str, got: int) -> None:
        lo, hi, where = arity[name.lower()]
        if lo <= got <= hi:
            return
        want = str(lo) if lo == hi else (
            f"{lo} or more" if hi > 10 ** 5 else f"{lo} to {hi}")
        out.append(Finding(
            "E012", mod_name, line,
            f"'{name}' called with {got} argument(s) but takes {want} "
            f"(defined in {where})"))

    def scan(mod_name: str, line: int, code: str) -> None:
        # expression form:  x = Foo(a, b)
        for m in re.finditer(r"(?<![\w.])([A-Za-z_]\w*)\s*\(", code):
            name = m.group(1)
            if name.lower() not in arity:
                continue
            i, depth, j = m.end(), 1, m.end()
            while j < len(code) and depth:
                if code[j] in "([":
                    depth += 1
                elif code[j] in ")]":
                    depth -= 1
                j += 1
            if depth:
                continue                    # unbalanced; E001 owns that
            inner = code[i:j - 1]
            if ":=" in inner:
                continue                    # named args, not positional
            report(mod_name, line, name, _count_args(inner))

        # statement form:  Foo a, b   /   Call Foo(a, b)
        s = code.strip()
        sm = re.match(r"^(?:Call\s+)?([A-Za-z_]\w*)\s*(.*)$", s)
        if not sm:
            return
        name, rest = sm.group(1), sm.group(2).strip()
        if name.lower() in _NOT_A_CALL or name.lower() not in arity:
            return
        if rest.startswith("(") or rest.startswith("="):
            return                          # handled above, or an assignment
        if ":=" in rest or "=" in rest.split("(")[0]:
            return
        report(mod_name, line, name, _count_args(rest))

    for mod in mods:
        for ll in mod.logical:
            if _SIG_RE.match(ll.text) or _declaration_line(ll.code.strip()):
                continue
            for stmt in _statements(ll.code):
                scan(mod.name, ll.start, stmt)

    return out


# --------------------------------------------------------------------------
# E004 / E005  name collisions
# --------------------------------------------------------------------------

def check_names(mods: list[Module]) -> list[Finding]:
    out: list[Finding] = []

    # duplicates inside one module
    for mod in mods:
        seen: dict[str, int] = {}
        for p in mod.procedures:
            low = p.name.lower()
            if low in seen:
                out.append(Finding("E005", mod.name, p.start,
                                   f"duplicate procedure name '{p.name}' "
                                   f"(first at line {seen[low]})"))
            else:
                seen[low] = p.start
        for name, line in list(mod.module_consts.items()):
            if name in seen:
                out.append(Finding("E005", mod.name, line,
                                   f"'{name}' declared both as a constant and "
                                   f"a procedure (line {seen[name]})"))
        for name, line in list(mod.module_vars.items()):
            if name in mod.module_consts:
                out.append(Finding("E005", mod.name, line,
                                   f"'{name}' declared twice at module level"))

    # public names shared across modules -> "Ambiguous name detected"
    owners: dict[str, list[tuple[str, int]]] = defaultdict(list)
    for mod in mods:
        for p in mod.procedures:
            if p.is_public:
                owners[p.name.lower()].append((mod.name, p.start))
        for name, line in mod.module_consts.items():
            if name in mod.public_names:
                owners[name].append((mod.name, line))
    for name, places in owners.items():
        if len({m for m, _ in places}) > 1:
            where = ", ".join(f"{m}:{ln}" for m, ln in places)
            out.append(Finding("E004", places[0][0], places[0][1],
                               f"Ambiguous name detected: Public '{name}' "
                               f"defined in more than one module ({where})"))
    return out


# --------------------------------------------------------------------------
# E006 / E007  line limits
# --------------------------------------------------------------------------

def check_line_limits(mod: Module) -> list[Finding]:
    out: list[Finding] = []
    for ll in mod.logical:
        if ll.segments - 1 > MAX_CONTINUATIONS:
            out.append(Finding("E006", mod.name, ll.start,
                               f"{ll.segments - 1} line continuations "
                               f"(VBA allows {MAX_CONTINUATIONS})"))
    for n, raw in enumerate(mod.raw, start=1):
        if len(raw) > MAX_LINE_CHARS:
            out.append(Finding("E007", mod.name, n,
                               f"physical line is {len(raw)} characters "
                               f"(VBA limit {MAX_LINE_CHARS})"))
    return out


# --------------------------------------------------------------------------
# E009  function with no return assignment
# --------------------------------------------------------------------------

def check_function_returns(mod: Module) -> list[Finding]:
    out: list[Finding] = []
    for p in mod.procedures:
        if p.kind.lower() not in ("function", "property get"):
            continue
        pat = re.compile(r"(^|[^.\w])" + re.escape(p.name) + r"\s*(\(.*\))?\s*=",
                         re.I)
        assigns = any(pat.search(ll.code) for ll in p.lines[1:])
        if not assigns:
            out.append(Finding("E009", mod.name, p.start,
                               f"{p.kind} {p.name} never assigns its return "
                               f"value"))
    return out


# --------------------------------------------------------------------------

def run(paths: list[Path]) -> list[Finding]:
    mods = parse_project(paths)
    findings: list[Finding] = []
    for mod in mods:
        findings += check_blocks(mod)
        findings += check_proc_termination(mod)
        findings += check_line_limits(mod)
        findings += check_function_returns(mod)
        findings += check_keyword_shadowing(mod)
        findings += check_declaration_section(mod)
    findings += check_names(mods)
    findings += check_arg_counts(mods)
    findings += check_identifiers(mods)
    findings += check_dead_code(mods)
    findings.sort(key=lambda f: (f.module, f.line, f.code))
    return findings


def main(argv: list[str]) -> int:
    root = Path(__file__).resolve().parent.parent
    if len(argv) > 1:
        paths = [Path(a) for a in argv[1:]]
    else:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import modules
        paths = modules.all_paths()
    if not paths:
        print("no VBA sources found", file=sys.stderr)
        return 2

    findings = run(paths)

    errors = [f for f in findings if not f.code.startswith("W")
              and f.code != "E009"]
    warnings = [f for f in findings if f.code.startswith("W")
                or f.code == "E009"]

    for f in errors:
        print(str(f))
    if warnings:
        print()
        for f in warnings:
            print(f"warning: {f}")

    print()
    print(f"{len(paths)} module(s) checked, "
          f"{len(errors)} error(s), {len(warnings)} warning(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
