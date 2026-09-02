# essentials

Everything needed to build and run the PnL Explainer. Nothing else.

```
vba/      the five modules to import into the VBE
access/   the DDL — 454 statements, 16 tables
tools/    build_access_db.py, which applies that DDL
config/   pipeline.yaml — the wiring
docs/     RUNBOOK.md (do this) and PIPELINE_SPEC.md (what it all means)
```

**Start with [docs/RUNBOOK.md](docs/RUNBOOK.md).** Ten steps, in order, each
ending with something you can check.

---

## The filenames in `vba/` are load-bearing

None of the five modules carries an `Attribute VB_Name` line, so **the VBE takes
the module name from the filename**. And two calls in the code are
module-qualified:

```vb
modAccess.Access_AutoSaveAfterLoad      ' in modPNL
modPNL.PnlLayout                        ' in modDashboard
```

So a module imported under a different name fails to compile, pointing at a line
that has nothing to do with the rename. Import these exactly as they are:

| File | Module it becomes |
|---|---|
| `vba/modPNL.bas` | `modPNL` |
| `vba/modAccess.bas` | `modAccess` |
| `vba/modDashboard.bas` | `modDashboard` |
| `vba/modEconFormulas.bas` | `modEconFormulas` |
| `vba/modImpRepo.bas` | `modImpRepo` |

**Remove any older copy before importing.** Importing beside an existing `modPNL`
gives you `modPNL1` and an ambiguous-name error that names neither.

**Do not open them in an editor that rewrites line endings.** They are pure CRLF.
A bare LF makes the VBE join or truncate lines, and the compile then fails
somewhere unrelated to the damage. Verified pure CRLF as shipped:

| | Lines |
|---|---|
| `modPNL.bas` | 13,902 |
| `modDashboard.bas` | 3,300 |
| `modAccess.bas` | 2,606 |
| `modEconFormulas.bas` | 795 |
| `modImpRepo.bas` | 585 |

---

## The two commands you need

Check the DDL — runs anywhere, needs only Python 3:

```
python3 tools/build_access_db.py --check
```

Expect `454 statements`, `16 tables, 10 indexes, 92 alters, 331 seed rows`.

Create a run database — **Windows only**, needs the Access Database Engine
matching Excel's bitness:

```
python3 tools/build_access_db.py --create "PNL_Run_20260902_173000.accdb"
```

`--name-for 2026-09-02T17:30:00` prints the correct filename for a retrieval
timestamp. In normal use the run creates its own database; do it by hand once to
prove the engine works before anything depends on it.

---

## Before you start

| | |
|---|---|
| Windows with Excel | the Bloomberg add-in is worksheet-function only; there is no API route |
| Access Database Engine | **matching Excel's bitness** — the one failure that gives no useful message |
| `M:\P_Pires\TRADING\` | the two Hedge Risco coverage books |
| OPICS on `LOKI` / `OPICSMAIN` | the bond query |
| Python 3 | only for the two commands above, not for the daily run |

**Back up the workbook first.** Loading a past run overwrites `Bonds`, `Swaps`
and `Futures`.

---

## Not yet executed

The VBA here is statically verified — lint clean, argument counts checked,
`Option Explicit` clean, pure CRLF — but it has **not been run**, because the
machine it was prepared on had neither Excel nor LibreOffice. Treat your first
pass through the runbook as the test, and read the *When a step fails* table at
the end of it before you start.

Five known defects are open and listed there. None of them announces itself.
