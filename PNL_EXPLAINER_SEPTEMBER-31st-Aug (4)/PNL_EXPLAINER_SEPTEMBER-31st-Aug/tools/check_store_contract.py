#!/usr/bin/env python3
"""Keep the run store and the sheets it snapshots in agreement.

The store is now three files' worth of agreement with no database to enforce
it, and every way it can drift is silent:

  modStore StoreSnapshotSheets   which sheets a run folder holds, and under
                                 which file name
  modPNL   PnlSheetExtent        how wide and how tall each of those sheets is
  docs/analysis/real_headers.tsv what the desk's workbook actually has

Nothing raises when these disagree.  A sheet listed for snapshot that
PnlSheetExtent does not know returns Empty, PnlRestoreSheet exits with 0, and
the run "loads" with that sheet still holding whatever was there before - live
formulas among restored values, reported nowhere.  An extent one column short
writes a file one column short, and the missing column restores blank.

  S001  a snapshotted sheet PnlSheetExtent cannot describe
  S002  a sheet PnlSheetExtent describes but no run folder holds
  S003  two sheets writing to one file name
  S004  a snapshot sheet name that no SH_ constant spells
  S005  an extent narrower than the sheet's real header row
  S006  a stored position kind with no field map, or a key field missing

Run:  python3 tools/check_store_contract.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import modules  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
HEADERS_TSV = ROOT / "docs" / "analysis" / "real_headers.tsv"

# Sheets PnlSheetExtent describes that a run folder deliberately does not hold.
# Config is snapshotted through StoreConfigSnapshot as named cells rather than
# as a rectangle, so it has no extent and needs no exemption here.
NOT_SNAPSHOTTED: dict[str, str] = {}

# What each position kind is identified by is READ OUT OF StoreKeyForRow rather
# than listed here.  A list here would be a fourth description to keep in step,
# and the checker would eventually be the thing that was wrong.


def function_body(text: str, name: str) -> str:
    m = re.search(
        r"^(?:Public |Private )?Function " + re.escape(name)
        + r"\b.*?^End Function",
        text, re.M | re.S)
    if not m:
        raise SystemExit(f"cannot find Function {name}")
    return m.group(0)


def parse_snapshot_sheets(text: str) -> list[tuple[str, str]]:
    """[(sheet, file), ...] from StoreSnapshotSheets."""
    body = function_body(text, "StoreSnapshotSheets")
    return re.findall(r'"([^"|]+)\|([^"]+)"', body)


def parse_extent_sheets(text: str) -> dict[str, str]:
    """SH_ constant -> the last-column expression, from PnlSheetExtent."""
    body = function_body(text, "PnlSheetExtent")
    out: dict[str, str] = {}
    sheet = None
    for line in body.splitlines():
        cm = re.match(r"\s*Case\s+(SH_\w+)\s*$", line)
        if cm:
            sheet = cm.group(1)
            continue
        if sheet is None:
            continue
        am = re.search(r"PnlSheetExtent\s*=\s*Array\(", line)
        if am:
            # The arms span two or three continued lines; the third element is
            # the last column and it is the one that decides the width.
            rest = body[body.index(line):]
            joined = " ".join(
                ln.strip().rstrip("_").strip()
                for ln in rest.splitlines()[:4])
            args = re.search(r"Array\((.*?)\)", joined)
            if args:
                parts = [p.strip() for p in args.group(1).split(",")]
                if len(parts) >= 3:
                    out[sheet] = parts[2]
            sheet = None
    return out


def parse_sheet_constants(text: str) -> dict[str, str]:
    """SH_ constant -> its literal sheet name."""
    return {
        m.group(1): m.group(2)
        for m in re.finditer(
            r"^(?:Public |Private )?Const (SH_\w+)\s+As String\s*=\s*\"([^\"]*)\"",
            text, re.M)
    }


def parse_col_constants(text: str) -> dict[str, str]:
    """Column constant -> its letter."""
    return {
        m.group(1): m.group(2)
        for m in re.finditer(
            r"^(?:Public |Private )?Const "
            r"((?:BCOL|PCOL|FCOL|WCOL|CVCOL|FMCOL)_\w+)"
            r"\s+As String\s*=\s*\"([A-Z]{1,3})\"",
            text, re.M)
    }


def parse_field_map(text: str) -> dict[str, list[str]]:
    """STORE_KIND_x -> [field names], from PnlPositionFieldMap."""
    body = function_body(text, "PnlPositionFieldMap")
    out: dict[str, list[str]] = {}
    kind = None
    for line in body.splitlines():
        cm = re.match(r"\s*Case\s+(STORE_KIND_\w+)", line)
        if cm:
            kind = cm.group(1)
            out[kind] = []
            continue
        if kind is None:
            continue
        for em in re.finditer(r'"(\w+)\|"', line):
            out[kind].append(em.group(1))
    return out


def parse_key_fields(store_text: str, kinds: dict[str, str]) -> dict[str, list[str]]:
    """STORE_KIND_x -> the field names StoreKeyForRow reads for that kind."""
    body = function_body(store_text, "StoreKeyForRow")
    literal_to_kind = {v: k for k, v in kinds.items()}
    out: dict[str, list[str]] = {}
    kind = None
    for line in body.splitlines():
        cm = re.match(r'\s*Case\s+"(\w+)"', line)
        if cm:
            kind = literal_to_kind.get(cm.group(1))
            if kind:
                out[kind] = []
            continue
        if kind is None:
            continue
        for em in re.finditer(r'StoreCell\([^)]*?"(\w+)"\)', line):
            out[kind].append(em.group(1))
    return out


def parse_kind_constants(text: str) -> dict[str, str]:
    """STORE_KIND_x -> its literal value."""
    return {
        m.group(1): m.group(2)
        for m in re.finditer(
            r"^(?:Public |Private )?Const (STORE_KIND_\w+)"
            r"\s+As String\s*=\s*\"([^\"]*)\"",
            text, re.M)
    }


def real_header_widths() -> dict[str, int]:
    """sheet -> how many columns the desk's workbook actually has."""
    out: dict[str, int] = {}
    for line in HEADERS_TSV.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        sheet, pos = parts[0], parts[1]
        try:
            out[sheet] = max(out.get(sheet, 0), int(pos))
        except ValueError:
            continue
    return out


def col_num(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def main() -> int:
    store_text = modules.path_of("modStore").read_text(encoding="latin-1")
    pnl_text = modules.path_of("modPNL").read_text(encoding="latin-1")

    snapshot = parse_snapshot_sheets(store_text)
    extents = parse_extent_sheets(pnl_text)
    sh_consts = parse_sheet_constants(pnl_text)
    cols = parse_col_constants(pnl_text)
    fields = parse_field_map(pnl_text)
    kinds = parse_kind_constants(pnl_text)
    key_fields = parse_key_fields(store_text, kinds)
    widths = real_header_widths()

    # SH_ constant -> literal, restricted to the ones PnlSheetExtent handles.
    extent_names = {
        sh_consts.get(k, k): v for k, v in extents.items()
    }

    problems: list[str] = []

    seen_files: dict[str, str] = {}
    for sheet, fname in snapshot:
        if sheet not in extent_names:
            problems.append(
                f"S001 a run folder holds {fname} for sheet '{sheet}', but "
                f"PnlSheetExtent has no arm for it - the snapshot would be "
                f"empty and the restore would silently leave the sheet alone")
        if sheet not in sh_consts.values():
            problems.append(
                f"S004 StoreSnapshotSheets names '{sheet}' but no SH_ constant "
                f"spells it - one of the two is a typo, and a typo here is a "
                f"sheet that never gets stored")
        if fname in seen_files:
            problems.append(
                f"S003 '{sheet}' and '{seen_files[fname]}' both write {fname} "
                f"- the second overwrites the first")
        seen_files[fname] = sheet

    snapshotted = {s for s, _ in snapshot}
    for const, letter_expr in sorted(extents.items()):
        name = sh_consts.get(const, const)
        if name in snapshotted or name in NOT_SNAPSHOTTED:
            continue
        problems.append(
            f"S002 PnlSheetExtent can snapshot '{name}' but no run folder "
            f"holds it - a restore would leave it as it was")

    for sheet, fname in snapshot:
        want = widths.get(sheet)
        if want is None:
            continue
        expr = extent_names.get(sheet)
        if expr is None:
            continue
        letter = cols.get(expr, expr.strip('"'))
        if not re.fullmatch(r"[A-Z]{1,3}", letter):
            continue
        have = col_num(letter)
        if have < want:
            problems.append(
                f"S005 '{sheet}' really has {want} columns but its extent "
                f"stops at {letter} ({have}) - {want - have} column(s) would "
                f"be missing from {fname} and blank after a restore")

    if not key_fields:
        problems.append(
            "S006 StoreKeyForRow has no arm this checker can read - the "
            "position keys are unchecked")

    for kind, keys in sorted(key_fields.items()):
        stored = fields.get(kind)
        if not stored:
            problems.append(
                f"S006 modPNL has no PnlPositionFieldMap arm for {kind}")
            continue
        for k in keys:
            if k not in stored:
                problems.append(
                    f"S006 {kind} builds its position key from '{k}' but no "
                    f"longer stores it - saved runs would stop matching and "
                    f"the duplicate guard would pass everything")

    for line in problems:
        print(line)

    print(f"\n{len(snapshot)} snapshotted sheet(s), "
          f"{sum(len(v) for v in fields.values())} stored position field(s), "
          f"{len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
