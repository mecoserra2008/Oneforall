#!/usr/bin/env python3
"""Every ```mermaid block in docs/ must actually parse.

A broken diagram does not fail loudly - GitHub renders an error box in place of
the picture, and only somebody reading that page finds out.  The diagrams carry
a real part of the explanation here, so they are checked like anything else.

Needs mermaid-cli and a browser; skips with a clear message when either is
absent, rather than failing a suite that has nothing to do with it.
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"

BROWSERS = (
    "/opt/pw-browsers/chromium",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome",
)


def blocks() -> list[tuple[Path, int, str]]:
    out = []
    for md in sorted(DOCS.glob("*.md")):
        for i, body in enumerate(
                re.findall(r"```mermaid\n(.*?)\n```", md.read_text(), re.S)):
            out.append((md, i, body))
    return out


def main() -> int:
    found = blocks()
    if not found:
        print("no mermaid diagrams found")
        return 0

    mmdc = shutil.which("mmdc")
    if not mmdc:
        print(f"SKIP: {len(found)} diagram(s) not checked (mermaid-cli not "
              f"installed: npm i -g @mermaid-js/mermaid-cli)")
        return 0

    browser = next((b for b in BROWSERS if Path(b).exists()), None)

    with tempfile.TemporaryDirectory() as tmp:
        tmpd = Path(tmp)
        cfg = tmpd / "pptr.json"
        cfg.write_text(json.dumps({
            **({"executablePath": browser} if browser else {}),
            "args": ["--no-sandbox", "--disable-gpu"],
        }))

        bad = 0
        for md, i, body in found:
            src = tmpd / f"{md.stem}_{i}.mmd"
            src.write_text(body)
            r = subprocess.run(
                [mmdc, "-p", str(cfg), "-i", str(src),
                 "-o", str(src.with_suffix(".svg")), "-q"],
                capture_output=True, text=True, timeout=180)
            if r.returncode != 0:
                err = (r.stderr or r.stdout).strip().splitlines()
                if any("Could not find" in ln or "Failed to launch" in ln
                       for ln in err):
                    print(f"SKIP: no usable browser for mermaid-cli "
                          f"({len(found)} diagram(s) not checked)")
                    return 0
                bad += 1
                print(f"{md.name} diagram {i + 1}: "
                      + " / ".join(ln for ln in err
                                   if "rror" in ln or "xpect" in ln)[:300])

    print(f"\n{len(found)} diagram(s) checked, {bad} problem(s)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
