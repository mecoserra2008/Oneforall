"""Minimal VBA source model: logical lines, procedures, declarations.

Shared by the linters in this folder.  It is deliberately small and
string-based: the goal is not a full VBA front end, only enough structure to
reproduce the compile-time errors that actually bite this project
(Option Explicit violations, missing procedures, ambiguous names, unbalanced
blocks).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

# --------------------------------------------------------------------------
# Lexical helpers
# --------------------------------------------------------------------------

_STRING_RE = re.compile(r'"(?:[^"]|"")*"')


def strip_strings(text: str) -> str:
    """Replace string literals with equal-length blanks so quoted text never
    looks like code to the scanners below."""
    return _STRING_RE.sub(lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', text)


def strip_comment(line: str) -> str:
    """Remove a trailing VBA comment, respecting string literals."""
    out = []
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif not in_str and ch == "'":
            break
        elif not in_str and line[i:i + 4].lower() == "rem " and (i == 0 or line[i - 1] in " \t:"):
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def code_of(line: str) -> str:
    """Comment-free, string-blanked code for a physical line."""
    return strip_strings(strip_comment(line))


# --------------------------------------------------------------------------
# Logical lines (continuations joined)
# --------------------------------------------------------------------------

@dataclass
class LogicalLine:
    text: str          # joined raw text (comments removed, strings intact)
    code: str          # same, with string literals blanked
    start: int         # 1-based physical line number of the first segment
    end: int           # 1-based physical line number of the last segment
    segments: int      # number of physical lines joined


def logical_lines(lines: list[str]) -> list[LogicalLine]:
    """Join `_` continuations into logical lines."""
    result: list[LogicalLine] = []
    buf: list[str] = []
    start = 0
    for idx, raw in enumerate(lines, start=1):
        body = strip_comment(raw).rstrip()
        if not buf:
            start = idx
        cont = body.endswith("_") and (len(body) == 1 or body[-2] in " \t")
        buf.append(body[:-1] if cont else body)
        if cont:
            continue
        text = " ".join(part.strip() for part in buf if part.strip())
        result.append(
            LogicalLine(
                text=text,
                code=strip_strings(text),
                start=start,
                end=idx,
                segments=len(buf),
            )
        )
        buf = []
    if buf:
        text = " ".join(part.strip() for part in buf if part.strip())
        result.append(
            LogicalLine(text=text, code=strip_strings(text),
                        start=start, end=len(lines), segments=len(buf))
        )
    return result


# --------------------------------------------------------------------------
# Procedures
# --------------------------------------------------------------------------

PROC_RE = re.compile(
    r"^\s*(?:(?P<scope>Public|Private|Friend|Global)\s+)?"
    r"(?:(?P<static>Static)\s+)?"
    r"(?P<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

END_PROC_RE = re.compile(
    r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)

DECLARE_RE = re.compile(
    r"^\s*(?:Public|Private)?\s*Declare\s+(?:PtrSafe\s+)?"
    r"(?:Sub|Function)\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

CONST_RE = re.compile(
    r"^\s*(?:(?P<scope>Public|Private|Global)\s+)?Const\s+(?P<rest>.+)$",
    re.IGNORECASE,
)

DIM_RE = re.compile(
    r"^\s*(?:(?P<scope>Dim|Private|Public|Global|Static|ReDim)\s+)"
    r"(?:Preserve\s+)?(?P<rest>.+)$",
    re.IGNORECASE,
)

ENUM_RE = re.compile(
    r"^\s*(?:(?:Public|Private)\s+)?Enum\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

TYPE_RE = re.compile(
    r"^\s*(?:(?:Public|Private)\s+)?Type\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

_NAME_LIST_SPLIT = re.compile(r",(?![^()]*\))")


def _names_from_decl(rest: str) -> list[str]:
    """Pull declared identifiers out of the RHS of Dim/Const/Private/Public."""
    names: list[str] = []
    depth = 0
    current = []
    for ch in rest:
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            names.append("".join(current))
            current = []
        else:
            current.append(ch)
    names.append("".join(current))

    out = []
    for chunk in names:
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)", chunk)
        if m:
            out.append(m.group(1))
    return out


def _params_of(signature: str) -> list[str]:
    """Parameter names from a Sub/Function/Property signature."""
    m = re.search(r"\((.*)\)", signature, re.DOTALL)
    if not m:
        return []
    inner = m.group(1)
    names = []
    depth = 0
    current = []
    for ch in inner:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            names.append("".join(current))
            current = []
        else:
            current.append(ch)
    names.append("".join(current))

    out = []
    for chunk in names:
        chunk = chunk.strip()
        if not chunk:
            continue
        chunk = re.sub(
            r"^(?:Optional\s+|ByVal\s+|ByRef\s+|ParamArray\s+)+", "",
            chunk, flags=re.IGNORECASE,
        )
        m2 = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)", chunk)
        if m2:
            out.append(m2.group(1))
    return out


@dataclass
class Procedure:
    name: str
    kind: str
    scope: str
    module: str
    start: int
    end: int
    lines: list[LogicalLine] = field(default_factory=list)
    params: list[str] = field(default_factory=list)

    @property
    def is_public(self) -> bool:
        return self.scope.lower() != "private"


@dataclass
class Module:
    name: str
    path: Path
    raw: list[str]
    logical: list[LogicalLine]
    option_explicit: bool
    procedures: list[Procedure]
    module_vars: dict[str, int]     # name -> declaring physical line
    module_consts: dict[str, int]
    public_names: set[str]
    enum_members: dict[str, int]
    types: dict[str, int]
    declares: dict[str, int]


def parse_module(path: Path) -> Module:
    raw = path.read_text(encoding="utf-8", errors="replace").splitlines()
    logical = logical_lines(raw)

    option_explicit = any(
        re.match(r"^\s*Option\s+Explicit\b", ll.text, re.IGNORECASE)
        for ll in logical
    )

    procedures: list[Procedure] = []
    module_vars: dict[str, int] = {}
    module_consts: dict[str, int] = {}
    enum_members: dict[str, int] = {}
    types: dict[str, int] = {}
    declares: dict[str, int] = {}
    public_names: set[str] = set()

    current: Procedure | None = None
    in_enum = False
    in_type = False

    for ll in logical:
        text = ll.text.strip()
        if not text:
            if current:
                current.lines.append(ll)
            continue

        if current is None:
            # ---- module level ------------------------------------------
            if in_enum:
                if re.match(r"^\s*End\s+Enum\b", text, re.IGNORECASE):
                    in_enum = False
                else:
                    m = re.match(r"^\s*([A-Za-z_][A-Za-z0-9_]*)", text)
                    if m:
                        enum_members[m.group(1).lower()] = ll.start
                continue
            if in_type:
                if re.match(r"^\s*End\s+Type\b", text, re.IGNORECASE):
                    in_type = False
                continue

            m = ENUM_RE.match(text)
            if m:
                in_enum = True
                types[m.group("name").lower()] = ll.start
                continue
            m = TYPE_RE.match(text)
            if m:
                in_type = True
                types[m.group("name").lower()] = ll.start
                continue

            m = DECLARE_RE.match(text)
            if m:
                declares[m.group("name").lower()] = ll.start
                public_names.add(m.group("name").lower())
                continue

            m = PROC_RE.match(text)
            if m:
                current = Procedure(
                    name=m.group("name"),
                    kind=re.sub(r"\s+", " ", m.group("kind")).title(),
                    scope=(m.group("scope") or "Public").title(),
                    module=path.stem,
                    start=ll.start,
                    end=ll.end,
                    params=_params_of(text),
                )
                current.lines.append(ll)
                continue

            m = CONST_RE.match(text)
            if m:
                for n in _names_from_decl(m.group("rest")):
                    module_consts[n.lower()] = ll.start
                    if (m.group("scope") or "").lower() in ("public", "global"):
                        public_names.add(n.lower())
                continue

            m = re.match(r"^\s*(Dim|Private|Public|Global)\s+(?!Const\b|Declare\b|Type\b|Enum\b|Sub\b|Function\b|Property\b|Const\b)(.+)$",
                         text, re.IGNORECASE)
            if m:
                for n in _names_from_decl(m.group(2)):
                    module_vars[n.lower()] = ll.start
                    if m.group(1).lower() in ("public", "global"):
                        public_names.add(n.lower())
                continue
            continue

        # ---- inside a procedure ----------------------------------------
        current.lines.append(ll)
        if END_PROC_RE.match(text):
            current.end = ll.end
            procedures.append(current)
            current = None

    if current is not None:              # unterminated procedure
        current.end = len(raw)
        procedures.append(current)

    for p in procedures:
        if p.is_public:
            public_names.add(p.name.lower())

    return Module(
        name=path.stem,
        path=path,
        raw=raw,
        logical=logical,
        option_explicit=option_explicit,
        procedures=procedures,
        module_vars=module_vars,
        module_consts=module_consts,
        public_names=public_names,
        enum_members=enum_members,
        types=types,
        declares=declares,
    )


def parse_project(paths: Iterable[Path]) -> list[Module]:
    return [parse_module(Path(p)) for p in paths]
