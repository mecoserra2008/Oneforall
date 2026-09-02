#!/usr/bin/env python3
"""Every .bas must be pure CRLF, and pure ASCII.

The VBE's File -> Import reads these as Windows text.  A file that has picked up
bare LFs - which any editor, or any script that round-trips it through Python's
universal newlines, will do silently - imports with joined or truncated lines
and fails to compile somewhere unrelated to the edit that caused it.  Cheap to
check, miserable to diagnose.

The ASCII rule is the same failure one layer down.  The VBE exports ANSI; an
editor opens the file as UTF-8; a script round-trips it; none of them announces
the change, and each pass mangles any accented character again.  Twelve
PNL_Attribution headers reached the sheet as "VariaAAo" that way, and the euro
sign in SwapFloatFamily had been through it FOUR times.  A file that is pure
ASCII cannot be corrupted by an encoding guess, so accented text is built from
code points instead - see PtVariacao in modPNL.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402


def main() -> int:
    bad = 0
    paths = modules.all_paths()
    for path in paths:
        data = path.read_bytes()
        lone = data.count(b"\n") - data.count(b"\r\n")
        stray = data.count(b"\r") - data.count(b"\r\n")
        if lone or stray:
            bad += 1
            print(f"{path.name}: {lone} bare LF, {stray} bare CR - must be CRLF")
        if data[:3] == b"\xef\xbb\xbf":
            bad += 1
            print(f"{path.name}: UTF-8 BOM; the VBE reads it as part of the "
                  f"first statement")

        high = [i for i, byte in enumerate(data) if byte > 127]
        if high:
            first = data[:high[0]].count(b"\n") + 1
            bad += 1
            print(f"{path.name}: {len(high)} non-ASCII byte(s), first on line "
                  f"{first} - build accented text from code points "
                  f"(ChrW$) so an encoding round trip cannot mangle it")
    print(f"\n{len(paths)} module(s) checked, {bad} problem(s)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
