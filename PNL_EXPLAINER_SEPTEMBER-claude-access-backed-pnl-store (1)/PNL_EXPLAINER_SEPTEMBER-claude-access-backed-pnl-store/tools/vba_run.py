#!/usr/bin/env python3
"""Execute VBA procedures from src/*.bas for real, headlessly.

There is no Excel here, so the numeric procedures are run through LibreOffice
Basic (`soffice --headless`), which is close enough to VBA for pure arithmetic,
dates and arrays.  Only *pure* procedures can be exercised this way - anything
that touches Range/Worksheet is out of scope, which is precisely why the model
code is split into a sheet-facing wrapper and a numeric core.

Usage (from a test file):
    from vba_run import BasicRunner
    r = BasicRunner()
    r.add_module("src/modImpRepo.bas")
    r.add_procedures("src/modPNL.bas", ["BondPullToParCore", ...])
    r.add_source(open("tests/test_pull_to_par.bas").read())
    print(r.run("RunTests"))
"""

from __future__ import annotations

import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vbaparse import parse_module                       # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

_MODULE_XBA = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" '
    '"module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="Module1" script:language="StarBasic">{body}</script:module>'
)

_SCRIPT_XLB = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" '
    '"library.dtd">\n'
    '<library:library xmlns:library="http://openoffice.org/2000/library" '
    'library:name="Standard" library:readonly="false" '
    'library:passwordprotected="false">\n'
    ' <library:element library:name="Module1"/>\n'
    '</library:library>\n'
)

_SCRIPT_XLC = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" '
    '"libraries.dtd">\n'
    '<library:libraries xmlns:library="http://openoffice.org/2000/library" '
    'xmlns:xlink="http://www.w3.org/1999/xlink">\n'
    ' <library:library library:name="Standard" xlink:href="$(USER)/basic/Standard/script.xlb/" '
    'xlink:type="simple" library:link="false"/>\n'
    '</library:libraries>\n'
)


class VbaRunError(RuntimeError):
    pass


# LibreOffice Basic has no Excel type library, so the Excel constants the model
# code relies on are supplied here with their real values.  CVErr(2042) is #N/A
# in both hosts, so IsError()/error-code assertions in the tests mean exactly
# what they mean in Excel.
PRELUDE = """Option VBASupport 1

Const xlErrNull   As Integer = 2000
Const xlErrDiv0   As Integer = 2007
Const xlErrValue  As Integer = 2015
Const xlErrRef    As Integer = 2023
Const xlErrName   As Integer = 2029
Const xlErrNum    As Integer = 2036
Const xlErrNA     As Integer = 2042
"""


def _strip_attributes(text: str) -> str:
    out = []
    for line in text.splitlines():
        if re.match(r"^\s*Attribute\s+VB_", line, re.I):
            continue
        out.append(line)
    return "\n".join(out)


def _parse_text(text: str):
    """Parse an in-memory .bas fragment (vbaparse works on files)."""
    tmp = Path(tempfile.mkstemp(suffix=".bas")[1])
    try:
        tmp.write_text(text)
        return parse_module(tmp)
    finally:
        tmp.unlink(missing_ok=True)


def _split_declarations(text: str) -> tuple[str, str]:
    """Separate a module's declaration section from its procedures.

    Basic, like VBA, requires every module-level Const/Dim/Type to appear
    BEFORE the first procedure.  Concatenating whole modules therefore has to
    hoist their declarations, or the second module's constants land after the
    first module's functions and the whole thing fails to compile - headless,
    that shows up as a modal dialog and a hang rather than an error message.
    """
    mod = _parse_text(text)
    lines = text.splitlines()
    in_proc = set()
    for proc in mod.procedures:
        in_proc.update(range(proc.start - 1, proc.end))

    decls = [l for i, l in enumerate(lines) if i not in in_proc]
    bodies = [l for i, l in enumerate(lines) if i in in_proc]
    return "\n".join(decls), "\n".join(bodies)


class BasicRunner:
    """Assemble Basic source from .bas fragments and execute one entry point."""

    def __init__(self) -> None:
        if SOFFICE is None:
            raise VbaRunError("soffice/libreoffice not found on PATH")
        self.decls: list[str] = []
        self.bodies: list[str] = []
        self.last_source_path: Path | None = None

    # -- source assembly ---------------------------------------------------

    def _add(self, text: str) -> None:
        decls, bodies = _split_declarations(_strip_attributes(text))
        if decls.strip():
            self.decls.append(decls)
        if bodies.strip():
            self.bodies.append(bodies)

    def add_module(self, path: str | Path, skip: set[str] | None = None) -> None:
        """Add a whole module, optionally dropping named procedures."""
        text = _strip_attributes(Path(path).read_text())
        if skip:
            mod = _parse_text(text)
            lines = text.splitlines()
            drop: set[int] = set()
            lowered = {s.lower() for s in skip}
            for proc in mod.procedures:
                if proc.name.lower() in lowered:
                    drop.update(range(proc.start - 1, proc.end))
            text = "\n".join(l for i, l in enumerate(lines) if i not in drop)
        self._add(text)

    def add_procedures(self, path: str | Path, names: list[str]) -> None:
        """Add only the named procedures, plus the constants they reference."""
        src = Path(path)
        text = _strip_attributes(src.read_text())
        mod = _parse_text(text)
        lines = text.splitlines()
        wanted = {n.lower() for n in names}

        bodies: list[str] = []
        found = set()
        for proc in mod.procedures:
            if proc.name.lower() in wanted:
                found.add(proc.name.lower())
                bodies.append("\n".join(lines[proc.start - 1:proc.end]))
        missing = wanted - found
        if missing:
            raise VbaRunError(f"procedures not found in {src}: {sorted(missing)}")

        # Only the module-level constants the extracted bodies actually
        # reference, taken as LOGICAL lines: several constants in this project
        # are `_`-continued, and slicing them by physical line silently yields a
        # truncated declaration that will not compile.
        joined = "\n".join(bodies)
        referenced = {
            name for name in mod.module_consts
            if re.search(r"(^|[^.\w])" + re.escape(name) + r"($|[^\w])",
                         joined, re.I)
        }

        consts: list[str] = []
        for ll in mod.logical:
            m = re.match(r"^\s*(?:Public|Private|Global)?\s*Const\s+"
                         r"([A-Za-z_]\w*)", ll.text, re.I)
            if m and m.group(1).lower() in referenced:
                consts.append(re.sub(
                    r"^\s*(?:Public|Private|Global)\s+Const\b", "Const",
                    ll.text, flags=re.I))

        if consts:
            self.decls.append("\n".join(consts))
        self.bodies.append("\n\n".join(bodies))

    def add_constants(self, path: str | Path, prefixes: tuple[str, ...]) -> None:
        """Add the module-level Const declarations matching any of `prefixes`.

        Lets a test reference the REAL column constants, so a test that pins a
        generated formula against a hand-written literal fails when a column
        moves - instead of quietly passing while the literal goes stale.
        """
        text = _strip_attributes(Path(path).read_text())
        mod = _parse_text(text)
        out: list[str] = []
        for ll in mod.logical:
            m = re.match(r"^\s*(?:Public|Private|Global)?\s*Const\s+"
                         r"([A-Za-z_]\w*)", ll.text, re.I)
            if m and m.group(1).startswith(prefixes):
                out.append(re.sub(r"^\s*(?:Public|Private|Global)\s+Const\b",
                                  "Const", ll.text, flags=re.I))
        if not out:
            raise VbaRunError(f"no constants matching {prefixes} in {path}")
        self.decls.append("\n".join(out))

    def add_source(self, text: str) -> None:
        self._add(text)

    # -- source rendering --------------------------------------------------

    def source(self) -> str:
        body = "\n\n".join(self.decls + self.bodies)
        # a single flat module: scope keywords carry no meaning and Option
        # Explicit is re-issued by the prelude
        body = re.sub(r"^\s*Option\s+Explicit\s*$", "", body, flags=re.I | re.M)
        body = re.sub(r"^(\s*)(?:Private|Public|Global)\s+(Function|Sub|Const)\b",
                      r"\1\2", body, flags=re.I | re.M)
        return PRELUDE + "\n\n" + body

    def _dump(self, profile: Path) -> None:
        dest = Path(tempfile.gettempdir()) / "vba_run_last.bas"
        dest.write_text(self.source())
        self.last_source_path = dest

    # -- execution ---------------------------------------------------------

    def run(self, entry: str = "RunTests", timeout: int = 240) -> str:
        home = Path(tempfile.mkdtemp(prefix="lo-home-"))
        profile = home / "profile"
        env = dict(os.environ)
        env["HOME"] = str(home)

        # A private UserInstallation per run, and no stale process left over.
        # LibreOffice is single-instance: an soffice.bin still up from an
        # earlier run - which is exactly what a Basic COMPILE error leaves
        # behind, sitting on a modal dialog - silently ADOPTS the next
        # invocation and runs ITS old module.  The symptom is a passing-looking
        # run against stale code, or "no output" with no error anywhere.
        subprocess.run(["pkill", "-x", "soffice.bin"], check=False)

        bootstrap = [
            SOFFICE, "--headless", "--norestore", "--terminate_after_init",
            f"-env:UserInstallation=file://{profile}",
        ]
        # Let LibreOffice build the profile itself; hand-rolling the directory
        # layout works on some builds and silently does nothing on others.
        subprocess.run(bootstrap, capture_output=True, text=True,
                       timeout=timeout, env=env)

        basic = profile / "user" / "basic" / "Standard"
        if not basic.is_dir():
            shutil.rmtree(home, ignore_errors=True)
            raise VbaRunError(
                f"LibreOffice did not create a Basic profile under {basic}")

        out_path = home / "vba-out.txt"
        source = self.source().replace("@@OUTPUT@@", str(out_path))
        self.last_source_path = Path(tempfile.gettempdir()) / "vba_run_last.bas"
        self.last_source_path.write_text(source)

        (basic / "Module1.xba").write_text(
            _MODULE_XBA.format(body=html.escape(source)))
        (basic / "script.xlb").write_text(_SCRIPT_XLB)

        cmd = [
            SOFFICE, "--headless", "--norestore", "--invisible",
            f"-env:UserInstallation=file://{profile}",
            f"vnd.sun.star.script:Standard.Module1.{entry}"
            f"?language=Basic&location=application",
        ]

        stdout = stderr = ""
        timed_out = False
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  timeout=timeout, env=env)
            stdout, stderr = proc.stdout, proc.stderr
        except subprocess.TimeoutExpired:
            timed_out = True
            # A Basic COMPILE error opens a modal dialog even headless, so the
            # process just sits there.  Kill it rather than hanging the suite.
            subprocess.run(["pkill", "-x", "soffice.bin"], check=False)

        if not out_path.exists():
            shutil.rmtree(home, ignore_errors=True)
            hint = ("timed out with no output - almost always a Basic COMPILE "
                    "error (modal dialog)" if timed_out
                    else "produced no output")
            raise VbaRunError(
                f"macro {hint}.\n"
                f"generated source: {self.last_source_path}\n"
                f"stdout: {stdout}\nstderr: {stderr}")

        result = out_path.read_text()
        shutil.rmtree(home, ignore_errors=True)
        return result
