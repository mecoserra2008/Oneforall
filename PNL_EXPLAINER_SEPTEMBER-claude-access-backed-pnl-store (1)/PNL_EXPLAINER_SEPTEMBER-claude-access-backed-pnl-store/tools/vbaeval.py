#!/usr/bin/env python3
"""Evaluate the VBA that builds Excel formula strings, without Excel.

Both this workbook's real problems and its documentation live in strings that
VBA concatenates at run time.  You cannot see a "#NAME?" or an unbalanced
bracket by reading the concatenation, and there is no Excel here to ask - so
this evaluates the concatenation the way VBA would, and hands back the finished
formula for something else to check or print.

It is a deliberately small interpreter: string expressions, the builtins these
modules actually use, calls into user procedures, and the control flow those
procedures actually contain (assignment, If/ElseIf/Else, For, For Each, Do
While, Exit).  Anything outside that raises, so an unrenderable formula is
reported rather than quietly rendered wrong.
"""
from __future__ import annotations

import re
from pathlib import Path

MAX_STEPS = 200000

# Excel enumeration members the modules refer to.  Their numeric values are
# irrelevant here - what matters is that they resolve rather than looking like
# an undeclared identifier.
EXCEL_ENUMS = {
    "xlContinuous": 1, "xlCenter": -4108, "xlLeft": -4131, "xlRight": -4152,
    "xlColumnClustered": 51, "xlBarClustered": 57, "xlLine": 4,
    "xlButtonControl": 0, "xlCalculationManual": -4135,
    "xlCalculationAutomatic": -4105, "xlUp": -4162, "xlDown": -4121,
    "xlByColumns": 2, "xlByRows": 1, "xlNext": 1, "xlPrevious": 2,
    "xlValues": -4163, "xlWhole": 1, "xlPart": 2, "xlNone": -4142,
    "vbInformation": 64, "vbExclamation": 48, "vbCritical": 16,
    "vbTextCompare": 1, "vbBinaryCompare": 0, "vbObjectError": -2147221504,
}


class VbaEvalError(Exception):
    pass


def split_args(s: str) -> list[str]:
    """Top-level comma split, respecting nesting and doubled quotes."""
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
        elif ch in "(":
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


OBJ = "\u00abobj\u00bb"


def _split_stmts(line: str) -> list[str]:
    """Split "a = 1: b = 2" at top level, leaving "Label:" alone."""
    out, depth, cur, in_str = [], 0, [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if in_str:
            cur.append(ch)
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    cur.append(line[i + 1]); i += 2; continue
                in_str = False
        elif ch == '"':
            in_str = True; cur.append(ch)
        elif ch == "(":
            depth += 1; cur.append(ch)
        elif ch == ")":
            depth -= 1; cur.append(ch)
        elif ch == ":" and depth == 0 and not (
                i + 1 < len(line) and line[i + 1] == "="):
            out.append("".join(cur)); cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return [s.strip() for s in out if s.strip()]


def join_continuations(text: str) -> str:
    return re.sub(r"[ \t]_\r?\n[ \t]*", " ", text)


class Vba:
    """One or more .bas modules, flattened, with their procedures callable."""

    def __init__(self, texts: dict[str, str],
                 stubs: dict[str, object] | None = None) -> None:
        self.stubs = dict(stubs or {})
        # Every "<target>.formula = <expr>" the interpreter walks past, with
        # the expression already evaluated.  Populated only while a caller has
        # asked for it; see check_dashboard.py.
        self.captured: list[tuple[str, str]] = []
        self.capture_errors: list[tuple[str, str]] = []
        self._current = "?"
        # Module-level state (Private dashRowCursor As Long) and the handful of
        # Excel enumeration constants these modules name.  Neither is declared
        # inside the procedures that read them.
        self.module_env: dict[str, object] = dict(EXCEL_ENUMS)
        self.consts: dict[str, object] = {}
        self.procs: dict[str, tuple[list[str], list[str], str]] = {}
        self._depth = 0

        for text in texts.values():
            flat = join_continuations(text)

            for m in re.finditer(
                    r"^\s*(?:Private|Public|Dim) (\w+) As (\w+)\s*$", flat, re.M):
                self.module_env.setdefault(
                    m.group(1),
                    0 if m.group(2) in ("Long", "Integer", "Double", "Single",
                                        "Currency", "Byte")
                    else (False if m.group(2) == "Boolean" else ""))

            for m in re.finditer(
                    r'^\s*(?:Public |Private |Global )?Const\s+(\w+)'
                    r'(?:\s+As\s+\w+)?\s*=\s*(.+?)\s*$', flat, re.M):
                # The value, ignoring any trailing ' comment - which most of
                # these lines carry, and which otherwise defeats the match.
                raw = m.group(2)
                lit = re.match(r'\s*"((?:[^"]|"")*)"', raw)
                if lit:
                    self.consts[m.group(1)] = lit.group(1).replace('""', '"')
                    continue
                num = re.match(r"\s*(-?\d+(?:\.\d+)?)(?:[eE][-+]?\d+)?", raw)
                if num and not re.match(r"\s*\w+\s*[&(]", raw):
                    v = num.group(0).strip()
                    self.consts[m.group(1)] = float(v) if "." in v or "e" in v.lower() else int(v)

            for m in re.finditer(
                    r'^\s*(?:Public |Private )?(Function|Sub) (\w+)\s*\(([^)]*)\)'
                    r'[^\r\n]*\r?\n(.*?)\r?\n\s*End \1', flat, re.S | re.M):
                params = [
                    re.sub(r"^(?:Optional )?(?:ByVal |ByRef )?(\w+).*$", r"\1",
                           a.strip())
                    for a in split_args(m.group(3)) if a.strip()
                ]
                defaults = {}
                for a in split_args(m.group(3)):
                    d = re.search(r"(\w+)[^=]*=\s*(.+)$", a.strip())
                    if d and "Optional" in a:
                        defaults[d.group(1)] = d.group(2).strip()
                self.procs[m.group(2)] = (params, list(m.group(4).split("\n")),
                                          m.group(2))
                self.procs[m.group(2)] = (params, m.group(4).split("\n"), m.group(2))
                self._defaults = getattr(self, "_defaults", {})
                self._defaults[m.group(2)] = defaults

    # -- expressions -------------------------------------------------------

    def expr(self, e: str, env: dict) -> object:
        """Evaluate one VBA expression.  Strings concatenate; numbers add."""
        e = e.strip()
        parts = self._split_concat(e)
        if len(parts) > 1:
            return "".join(self._as_str(self.expr(p, env)) for p in parts)
        return self._atom(e, env)

    @staticmethod
    def _as_str(v: object) -> str:
        if isinstance(v, bool):
            return "True" if v else "False"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    @staticmethod
    def _split_concat(e: str) -> list[str]:
        out, depth, cur, in_str = [], 0, [], False
        i = 0
        while i < len(e):
            ch = e[i]
            if in_str:
                cur.append(ch)
                if ch == '"':
                    if i + 1 < len(e) and e[i + 1] == '"':
                        cur.append(e[i + 1]); i += 2; continue
                    in_str = False
            elif ch == '"':
                in_str = True; cur.append(ch)
            elif ch == "(":
                depth += 1; cur.append(ch)
            elif ch == ")":
                depth -= 1; cur.append(ch)
            elif ch == "&" and depth == 0:
                out.append("".join(cur)); cur = []
            else:
                cur.append(ch)
            i += 1
        out.append("".join(cur))
        return [p for p in out if p.strip()]

    def _atom(self, e: str, env: dict) -> object:
        e = e.strip()
        if not e:
            return ""

        lit = re.fullmatch(r'"((?:[^"]|"")*)"', e)
        if lit:
            return lit.group(1).replace('""', '"')
        if re.fullmatch(r"-?\d+", e):
            return int(e)
        if re.fullmatch(r"-?\d*\.\d+", e):
            return float(e)
        if e in ("True", "False"):
            return e == "True"
        if e in ("vbCrLf", "vbNewLine"):
            return "\n"
        if e == "vbNullString":
            return ""

        if e.startswith("(") and e.endswith(")"):
            inner = e[1:-1]
            if split_args(inner) == [inner] and self._balanced(inner):
                return self.expr(inner, env)

        # Operators, lowest precedence first, split at the LAST top-level
        # occurrence so evaluation is left-associative.  Splitting by regex
        # instead read "((n - 1) Mod 26)" as "(n" minus "1) Mod 26".
        for ops in (("Or", "And"), ("<>", "<=", ">=", "=", "<", ">"),
                    ("+", "-"), ("*", "/", "\\", "Mod")):
            found = self._split_op(e, ops)
            if found:
                lhs, op, rhs = found
                a = self.expr(lhs, env)
                b = self.expr(rhs, env)
                try:
                    return self._apply(op, a, b)
                except TypeError:
                    raise VbaEvalError(f"cannot apply {op} to {a!r} and {b!r}")

        m = re.fullmatch(r"(\w+\$?)\s*\((.*)\)", e, re.S)
        if m:
            return self.call(m.group(1), split_args(m.group(2)), env)

        # A member reference on something this interpreter has no model of -
        # ws.Cells(...).value, a Range address, a cell's contents.  Rendered as
        # a sentinel so the surrounding formula still assembles and the caller
        # can see it was not fully resolved, rather than the whole formula
        # being lost to one unknowable term.
        if re.search(r"[A-Za-z_)\]]\.[A-Za-z_]", e):
            return OBJ

        name = e.rstrip("$")
        if name in env:
            return env[name]
        if name in self.consts:
            return self.consts[name]
        if name in self.module_env:
            return self.module_env[name]
        if name in self.procs:
            return self.call(name, [], env)
        if name in self.stubs:
            return self.call(name, [], env)
        raise VbaEvalError(f"unknown identifier {e!r}")

    @staticmethod
    def _apply(op: str, a, b):
        if op == "+":
            return str(a) + str(b) if isinstance(a, str) or isinstance(b, str) else a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if op == "/":
            return a / b
        if op == "\\":
            return int(a) // int(b)
        if op == "Mod":
            return int(a) % int(b)
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == ">":
            return a > b
        if op == "<":
            return a < b
        if op == ">=":
            return a >= b
        if op == "<=":
            return a <= b
        if op == "And":
            return Vba._truth(a) and Vba._truth(b)
        if op == "Or":
            return Vba._truth(a) or Vba._truth(b)
        raise VbaEvalError(f"unknown operator {op}")

    @staticmethod
    def _split_op(e: str, ops: tuple[str, ...]):
        """Last top-level occurrence of any of `ops`, or None."""
        depth, in_str, best = 0, False, None
        i = 0
        while i < len(e):
            ch = e[i]
            if in_str:
                if ch == '"':
                    if i + 1 < len(e) and e[i + 1] == '"':
                        i += 2; continue
                    in_str = False
                i += 1
                continue
            if ch == '"':
                in_str = True; i += 1; continue
            if ch == "(":
                depth += 1; i += 1; continue
            if ch == ")":
                depth -= 1; i += 1; continue
            if depth == 0:
                for op in ops:
                    if not e.startswith(op, i):
                        continue
                    if op[0].isalpha():
                        before = e[i - 1] if i else " "
                        after = e[i + len(op)] if i + len(op) < len(e) else " "
                        if before.isalnum() or after.isalnum() or after == "_":
                            continue
                    # a leading sign, not a binary operator
                    if op in "+-" and not e[:i].strip():
                        continue
                    if op in "+-" and e[:i].rstrip()[-1:] in "(,&*/+-=<>":
                        continue
                    best = (e[:i], op, e[i + len(op):])
                    i += len(op)
                    break
                else:
                    i += 1
                continue
            i += 1
        return best

    @staticmethod
    def _balanced(s: str) -> bool:
        depth, in_str = 0, False
        i = 0
        while i < len(s):
            ch = s[i]
            if in_str:
                if ch == '"':
                    if i + 1 < len(s) and s[i + 1] == '"':
                        i += 2; continue
                    in_str = False
            elif ch == '"':
                in_str = True
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    return False
            i += 1
        return depth == 0

    # -- calls -------------------------------------------------------------

    BUILTIN = {
        "CStr": lambda v: Vba._as_str(v),
        "CLng": lambda v: int(float(v)),
        "CDbl": lambda v: float(v),
        "Trim": lambda v: str(v).strip(),
        "LTrim": lambda v: str(v).lstrip(),
        "RTrim": lambda v: str(v).rstrip(),
        "UCase": lambda v: str(v).upper(),
        "LCase": lambda v: str(v).lower(),
        "Len": lambda v: len(str(v)),
        "Chr": lambda v: chr(int(v)),
        "Asc": lambda v: ord(str(v)[0]),
        "Abs": lambda v: abs(v),
        "Int": lambda v: int(v),
    }

    def call(self, name: str, args: list[str], env: dict) -> object:
        base = name.rstrip("$")

        if base in self.stubs:
            stub = self.stubs[base]
            vals = [self.expr(a, env) for a in args if a.strip()]
            return stub(*vals) if callable(stub) else stub

        if base in self.BUILTIN and len(args) == 1:
            return self.BUILTIN[base](self.expr(args[0], env))

        if base in ("IsError", "IsNull", "IsEmpty", "IsMissing"):
            return False
        if base in ("IsNumeric", "IsDate", "IsArray", "IsObject"):
            v = self.expr(args[0], env)
            if base == "IsNumeric":
                return isinstance(v, (int, float)) and not isinstance(v, bool)
            if base == "IsArray":
                return isinstance(v, list)
            return False
        if base == "Nz":
            return self.expr(args[0], env)
        if base == "IIf":
            return self.expr(args[1] if self.expr(args[0], env) else args[2], env)
        if base == "Left":
            return str(self.expr(args[0], env))[:int(self.expr(args[1], env))]
        if base == "Right":
            n = int(self.expr(args[1], env))
            s = str(self.expr(args[0], env))
            return s[-n:] if n else ""
        if base == "Mid":
            s = str(self.expr(args[0], env))
            st = int(self.expr(args[1], env)) - 1
            if len(args) > 2:
                return s[st:st + int(self.expr(args[2], env))]
            return s[st:]
        if base == "Replace":
            return str(self.expr(args[0], env)).replace(
                str(self.expr(args[1], env)), str(self.expr(args[2], env)))
        if base == "String":
            return str(self.expr(args[1], env)) * int(self.expr(args[0], env))
        if base == "Split":
            s = str(self.expr(args[0], env))
            sep = str(self.expr(args[1], env)) if len(args) > 1 else " "
            return s.split(sep)
        if base == "Array":
            return [self.expr(a, env) for a in args if a.strip()]
        if base == "LBound":
            return 0
        if base == "UBound":
            v = self.expr(args[0], env)
            return len(v) - 1
        if base == "InStr":
            vals = [self.expr(a, env) for a in args]
            if len(vals) >= 3 and isinstance(vals[0], int):
                return str(vals[1]).find(str(vals[2]), vals[0] - 1) + 1
            return str(vals[0]).find(str(vals[1])) + 1

        if base in self.procs:
            params, body, _ = self.procs[base]
            defaults = getattr(self, "_defaults", {}).get(base, {})
            local = {}
            for i, prm in enumerate(params):
                if i < len(args) and args[i].strip():
                    local[prm] = self.expr(args[i], env)
                elif prm in defaults:
                    local[prm] = self.expr(defaults[prm], env)
            for prm in params:
                local.setdefault(prm, "")
            return self.run(base, body, local)

        # a subscript into a bound array: parts(i)
        if base in env and isinstance(env[base], list):
            return env[base][int(self.expr(args[0], env))]

        raise VbaEvalError(f"unknown call {name}(")

    # -- statements --------------------------------------------------------

    def run(self, name: str, body: list[str], env: dict) -> object:
        self._depth += 1
        if self._depth > 40:
            self._depth -= 1
            raise VbaEvalError(f"{name}: recursion too deep")
        prev, self._current = self._current, name
        try:
            env.setdefault(name, "")
            # VBA initialises locals: a String is "", a numeric type is 0.
            # Seeding them from the Dim lines is what lets an accumulator like
            # "s = Chr$(65 + r) & s" work on its first pass, while an identifier
            # that was never declared still raises instead of silently reading
            # as empty.
            for ln in body:
                for m in re.finditer(r"\b(\w+)(?:\(\))?\s+As\s+(\w+)",
                                     re.sub(r"^\s*Dim\s+", "", ln.strip())
                                     if ln.strip().startswith("Dim ") else ""):
                    env.setdefault(
                        m.group(1),
                        0 if m.group(2) in ("Long", "Integer", "Double",
                                            "Single", "Currency", "Byte")
                        else (False if m.group(2) == "Boolean" else ""))
            self._block([ln.strip() for ln in body], env, name)
            return env.get(name, "")
        finally:
            self._depth -= 1
            self._current = prev

    def _block(self, lines: list[str], env: dict, name: str,
               steps: list[int] | None = None) -> None:
        steps = steps if steps is not None else [0]
        i = 0
        while i < len(lines):
            steps[0] += 1
            if steps[0] > MAX_STEPS:
                raise VbaEvalError(f"{name}: step limit")
            line = lines[i]
            i += 1
            if not line or line.startswith("'"):
                continue
            line = re.sub(r"\s+'(?:[^\"]|\"[^\"]*\")*$", "", line).strip()
            if not line or line.startswith(("Dim ", "ReDim ", "On Error",
                                            "Set ", "Const ", "Static ")):
                continue

            # "a = 1: b = 2" is two statements.  Splitting them is what makes
            # a trailing ": r = r + 1" - which this module uses a lot - run
            # instead of being read as part of the expression before it.
            parts = _split_stmts(line)
            if len(parts) > 1:
                self._block(parts, env, name, steps)
                continue
            if line in ("Exit Function", "Exit Sub"):
                return

            m = re.match(r"(.+?)\.(?:formula|Formula|FormulaR1C1)\s*=\s*(.+)$",
                         line, re.I)
            if m:
                try:
                    self.captured.append(
                        (self._current, self._as_str(self.expr(m.group(2), env))))
                except Exception as exc:                       # noqa: BLE001
                    self.capture_errors.append(
                        (self._current, f"{type(exc).__name__}: {exc}"
                                        f"  |  {m.group(2)[:120]}"))
                continue

            m = re.match(r"If (.+) Then$", line)
            if m:
                end = self._match_end(lines, i - 1, "If", "End If")
                self._run_if(lines, i - 1, end, env, name, steps)
                i = end + 1
                continue

            m = re.match(r"If (.+) Then (.+)$", line)
            if m and not m.group(2).startswith("'"):
                if self._truth(self.expr(m.group(1), env)):
                    self._block([m.group(2)], env, name, steps)
                continue

            m = re.match(r"Select Case (.+)$", line)
            if m:
                end = self._match_end(lines, i - 1, "Select", "End Select")
                self._run_select(m.group(1), lines, i - 1, end, env, name, steps)
                i = end + 1
                continue

            m = re.match(r"For Each (\w+) In (.+)$", line)
            if m:
                end = self._match_end(lines, i - 1, "For", "Next")
                for v in self.expr(m.group(2), env):
                    env[m.group(1)] = v
                    self._block(lines[i:end], env, name, steps)
                i = end + 1
                continue

            m = re.match(r"For (\w+) = (.+?) To (.+?)(?: Step (.+))?$", line)
            if m:
                end = self._match_end(lines, i - 1, "For", "Next")
                lo = int(self.expr(m.group(2), env))
                hi = int(self.expr(m.group(3), env))
                step = int(self.expr(m.group(4), env)) if m.group(4) else 1
                for v in range(lo, hi + 1, step) if step > 0 else range(lo, hi - 1, step):
                    env[m.group(1)] = v
                    self._block(lines[i:end], env, name, steps)
                i = end + 1
                continue

            m = re.match(r"Do While (.+)$", line)
            if m:
                end = self._match_end(lines, i - 1, "Do", "Loop")
                guard = 0
                while self._truth(self.expr(m.group(1), env)):
                    guard += 1
                    if guard > 5000:
                        raise VbaEvalError(f"{name}: Do While did not terminate")
                    self._block(lines[i:end], env, name, steps)
                i = end + 1
                continue

            m = re.match(r"(\w+)\s*=\s*(.+)$", line)
            if m:
                env[m.group(1)] = self.expr(m.group(2), env)
                continue

            m = re.match(r"(\w+)\((.+?)\)\s*=\s*(.+)$", line)
            if m:
                continue                                   # array element write

            m = re.match(r"(\w+)\s+(.+)$", line)
            if m and m.group(1) in self.procs:
                self.call(m.group(1), split_args(m.group(2)), env)
                continue
            if line in self.procs:
                self.call(line, [], env)
                continue
            # anything else (MsgBox, ws.Range..., Next, Loop, End If) is ignored

    def _run_select(self, sel_expr: str, lines: list[str], start: int, end: int,
                    env: dict, name: str, steps: list[int]) -> None:
        selector = self.expr(sel_expr, env)

        arms: list[tuple[list[str] | None, int]] = []
        depth = 0
        for j in range(start, end):
            ln = lines[j]
            if re.match(r"Select Case\b", ln):
                depth += 1
                continue
            if ln == "End Select":
                depth -= 1
                continue
            if depth != 1:
                continue
            if ln == "Case Else":
                arms.append((None, j + 1))
            else:
                m = re.match(r"Case (.+)$", ln)
                if m:
                    arms.append((split_args(m.group(1)), j + 1))

        bounds = [a[1] for a in arms] + [end]
        for k, (tests, first) in enumerate(arms):
            hit = tests is None
            for test in tests or []:
                test = test.strip()
                m = re.match(r"Is\s*(<>|>=|<=|=|<|>)\s*(.+)$", test)
                if m:
                    if self._truth(self._apply(m.group(1), selector,
                                               self.expr(m.group(2), env))):
                        hit = True
                elif self.expr(test, env) == selector:
                    hit = True
            if hit:
                last = bounds[k + 1] - 1 if k + 1 < len(arms) else end
                self._block(lines[first:last], env, name, steps)
                return

    def _run_if(self, lines: list[str], start: int, end: int, env: dict,
                name: str, steps: list[int]) -> None:
        """Run the first arm of an If block whose condition holds."""
        arms: list[tuple[str | None, int]] = []
        depth = 0
        for j in range(start, end):
            ln = lines[j]
            if re.match(r"If .+ Then$", ln):
                depth += 1
                if depth == 1:
                    arms.append((re.match(r"If (.+) Then$", ln).group(1), j + 1))
                continue
            if ln == "End If":
                depth -= 1
                continue
            if depth == 1 and re.match(r"ElseIf (.+) Then$", ln):
                arms.append((re.match(r"ElseIf (.+) Then$", ln).group(1), j + 1))
            elif depth == 1 and ln == "Else":
                arms.append((None, j + 1))

        bounds = [a[1] for a in arms] + [end]
        for k, (cond, first) in enumerate(arms):
            if cond is None or self._truth(self.expr(cond, env)):
                self._block(lines[first:bounds[k + 1] - 1 if k + 1 < len(arms)
                                  else end], env, name, steps)
                return

    @staticmethod
    def _truth(v: object) -> bool:
        if isinstance(v, str):
            return v not in ("", "False")
        return bool(v)

    @staticmethod
    def _match_end(lines: list[str], start: int, open_kw: str, close_kw: str) -> int:
        depth = 0
        for j in range(start, len(lines)):
            ln = lines[j]
            if open_kw == "If":
                if re.match(r"If .+ Then$", ln):
                    depth += 1
                elif ln == "End If":
                    depth -= 1
            elif open_kw == "Select":
                if re.match(r"Select Case\b", ln):
                    depth += 1
                elif ln == "End Select":
                    depth -= 1
            elif open_kw == "For":
                if re.match(r"For\b", ln):
                    depth += 1
                elif re.match(r"Next\b", ln):
                    depth -= 1
            else:
                if re.match(r"Do\b", ln):
                    depth += 1
                elif re.match(r"Loop\b", ln):
                    depth -= 1
            if depth == 0:
                return j
        raise VbaEvalError(f"unterminated {open_kw} at line {start}")


def load(paths: dict[str, Path], stubs: dict | None = None) -> Vba:
    return Vba({k: p.read_text(encoding="latin-1") for k, p in paths.items()},
               stubs)
