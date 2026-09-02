#!/usr/bin/env python3
"""Self-test for tools/vba_lint.py.

A linter that quietly stops detecting things is worse than no linter: the
build stays green and the confidence is unearned.  Each case below is a
deliberately broken module that MUST produce a specific finding, plus a clean
module that must produce none.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import vba_lint                                          # noqa: E402

CLEAN = '''Option Explicit

Private Const GREETING As String = "hello"

Public Function Add(ByVal a As Long, ByVal b As Long) As Long
    Add = a + b
End Function

Public Sub Show()
    Dim total As Long
    Dim i As Long
    For i = 1 To 3
        If i > 1 Then
            total = total + Add(i, 1)
        End If
    Next i
    Debug.Print GREETING & CStr(total)
End Sub
'''

CASES: list[tuple[str, str, str]] = [
    (
        "E002 undeclared variable",
        "E002",
        '''Option Explicit

Public Sub Broken()
    Dim known As Long
    known = 1
    total = known + 1
End Sub
''',
    ),
    (
        "E002 missing Option Explicit",
        "E002",
        '''Public Sub Fine()
    Debug.Print 1
End Sub
''',
    ),
    (
        "E003 call to a procedure that does not exist",
        "E003",
        '''Option Explicit

Public Sub Broken()
    Dim x As Long
    x = NoSuchHelper(1, 2)
End Sub
''',
    ),
    (
        "E001 unbalanced If",
        "E001",
        '''Option Explicit

Public Sub Broken()
    Dim i As Long
    If i > 0 Then
        i = 1
End Sub
''',
    ),
    (
        "E001 Next without For",
        "E001",
        '''Option Explicit

Public Sub Broken()
    Dim i As Long
    i = 0
    Next i
End Sub
''',
    ),
    (
        "E005 duplicate procedure name",
        "E005",
        '''Option Explicit

Private Sub Twice()
    Debug.Print 1
End Sub

Private Sub Twice()
    Debug.Print 2
End Sub
''',
    ),
    (
        "E006 too many line continuations",
        "E006",
        "Option Explicit\n\nPublic Sub Broken()\n    Dim s As String\n    s = \"a\" _\n"
        + "        & \"b\" _\n" * 30
        + "        & \"c\"\nEnd Sub\n",
    ),
    (
        "E009 function never assigns its result",
        "E009",
        '''Option Explicit

Private Function Forgot(ByVal a As Long) As Long
    Dim b As Long
    b = a + 1
End Function

Public Sub Use()
    Debug.Print Forgot(1)
End Sub
''',
    ),
    (
        "E010 parameter shadows the Base keyword",
        "E010",
        '''Option Explicit

Public Function Slope(ByVal base As Double, ByVal step2 As Double) As Double
    Slope = base + step2
End Function
''',
    ),
    (
        "E010 local shadows the Name keyword",
        "E010",
        '''Option Explicit

Public Sub Broken()
    Dim name As String
    name = "x"
    Debug.Print name
End Sub
''',
    ),
    (
        # The trap: appending a feature to the END of a module is the natural
        # way to add one, and the constant it needs travels down with it.  VBA
        # rejects the whole module, at a line nobody was editing.
        "E011 module Const declared after the first procedure",
        "E011",
        '''Option Explicit

Public Sub Entry()
    Debug.Print LATE_ONE
End Sub

Private Const LATE_ONE As String = "x"
''',
    ),
    (
        "E011 module variable declared after the first procedure",
        "E011",
        '''Option Explicit

Public Sub Entry()
    gCounter = gCounter + 1
End Sub

Private gCounter As Long
''',
    ),
    (
        "W001 unreferenced private procedure",
        "W001",
        '''Option Explicit

Private Sub NobodyCallsThis()
    Debug.Print 1
End Sub

Public Sub Entry()
    Debug.Print 2
End Sub
''',
    ),
]


def run_on(source: str) -> list[str]:
    with tempfile.TemporaryDirectory() as d:
        path = Path(d) / "modProbe.bas"
        path.write_text(source)
        return [f.code for f in vba_lint.run([path])]


def main() -> int:
    failures = 0
    passes = 0

    clean = run_on(CLEAN)
    if clean:
        print(f"FAIL  clean module reports {clean}")
        failures += 1
    else:
        print("PASS  clean module reports nothing")
        passes += 1

    for label, expected, source in CASES:
        codes = run_on(source)
        if expected in codes:
            print(f"PASS  {label}")
            passes += 1
        else:
            print(f"FAIL  {label} -> expected {expected}, got {codes or 'nothing'}")
            failures += 1

    # E004 needs two modules
    with tempfile.TemporaryDirectory() as d:
        a = Path(d) / "modA.bas"
        b = Path(d) / "modB.bas"
        a.write_text("Option Explicit\n\nPublic Sub Shared_()\n    Debug.Print 1\nEnd Sub\n")
        b.write_text("Option Explicit\n\nPublic Sub Shared_()\n    Debug.Print 2\nEnd Sub\n")
        codes = [f.code for f in vba_lint.run([a, b])]
        if "E004" in codes:
            print("PASS  E004 same Public name in two modules")
            passes += 1
        else:
            print(f"FAIL  E004 -> got {codes or 'nothing'}")
            failures += 1

    print(f"\npassed={passes} failed={failures}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
