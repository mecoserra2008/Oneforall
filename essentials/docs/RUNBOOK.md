# Runbook — setting up and running the Access-backed pipeline

Step by step, from a machine with nothing on it to a stored run you can load back.

Do the steps in order. Each ends with something you can check, so a failure
surfaces at the step that caused it rather than three steps later.

> **This supersedes [SETUP_ACCESS_STORE.md](SETUP_ACCESS_STORE.md)**, which
> describes the earlier single-database design. Three things in it are now wrong:
> `Config!B41` is a **folder** and not a file path, each run gets its **own**
> `.accdb`, and there **is** DDL to run — the store is 16 tables now, not the six
> that `AccEnsureSchema` creates by itself.

> **Not yet executed anywhere.** The VBA changes behind this are statically
> verified — lint clean, argument counts checked, `Option Explicit` clean, pure
> CRLF — but no Excel or LibreOffice was available to run them. Treat the first
> pass through this runbook as the test.

- [What you need](#what-you-need)
- [Step 1 · The Access engine, and its bitness](#step-1--the-access-engine-and-its-bitness)
- [Step 2 · Build the run database](#step-2--build-the-run-database)
- [Step 3 · Import the five modules](#step-3--import-the-five-modules)
- [Step 4 · Compile](#step-4--compile)
- [Step 5 · Lay out the workbook](#step-5--lay-out-the-workbook)
- [Step 6 · Set the Config cells](#step-6--set-the-config-cells)
- [Step 7 · Test the connection](#step-7--test-the-connection)
- [Step 8 · Run the day](#step-8--run-the-day)
- [Step 9 · Check what was stored](#step-9--check-what-was-stored)
- [Step 10 · Load a past run back](#step-10--load-a-past-run-back)
- [Every day after that](#every-day-after-that)
- [When a step fails](#when-a-step-fails)

---

## What you need

| | |
|---|---|
| Windows with Excel | the Bloomberg add-in is a worksheet-function add-in; there is no API route |
| The Access Database Engine | **matching Excel's bitness** — Step 1 |
| Network access to `M:\P_Pires\TRADING\` | the two coverage books |
| An OPICS connection to `LOKI` / `OPICSMAIN` | the bond query |
| Python 3 with PyYAML | only for the tooling in `tools/`, not for the daily run |

**Back up the workbook first.** Copy the `.xlsm` somewhere. Step 10 overwrites
`Bonds`, `Swaps` and `Futures`, and you want to be able to go back without
thinking about it.

---

## Step 1 · The Access engine, and its bitness

This is the one failure that gives you no useful message, so it goes first.

**Excel's bitness:** File → Account → About Excel. The version line ends `32-bit`
or `64-bit`.

**The engine:** Settings → Apps → Installed apps, search *Access*. You want
"Microsoft Access Database Engine 2016 Redistributable" (or a full Access
install) **in the same bitness as Excel**. A 64-bit Excel cannot talk to a 32-bit
engine, and the error it raises names neither.

If it is missing, install the redistributable matching Excel. If you have 32-bit
Office and the 64-bit engine, that is the mismatch — uninstall and reinstall the
right one.

**Check:** in the VBE Immediate window (`Ctrl+G`):

```vb
?CreateObject("ADOX.Catalog") Is Nothing
```

`False` means the engine is there. An error means it is not.

---

## Step 2 · Build the run database

**This step is new.** The store is 16 tables. `AccEnsureSchema` in the VBA
creates six of them — the position tables — and nothing else, so the dimension,
raw-capture and computed-fact tables have to come from the DDL.

First, check the DDL parses and covers everything. This runs anywhere, including
CI:

```
python3 tools/build_access_db.py --check
```

Expect:

```
  schema.sql                     29 statements
  alter_run_scope.sql             5 statements
  alter_meta_column.sql           5 statements
  alter_fact_pnl.sql             84 statements
  seed_meta_column.sql          331 statements

  total                         454 statements

  16 tables, 10 indexes, 92 alters, 331 seed rows
```

Then create the database itself, **on the Windows machine**:

```
python3 tools/build_access_db.py --create "M:\path\to\runs\PNL_Run_20260902_173000.accdb"
```

It refuses to overwrite an existing file. The name is not free-form — Step 6
explains why, and `--name-for` gives you the right one:

```
python3 tools/build_access_db.py --name-for 2026-09-02T17:30:00
  -> PNL_Run_20260902_173000.accdb
```

**In practice you do not create these by hand.** The run creates its own database
when Button 1 runs. Do it manually once, here, to prove the engine works and the
DDL applies before anything depends on it.

**Check:** open the `.accdb`. Sixteen tables, `Meta_Column` holding 331 rows.

---

## Step 3 · Import the five modules

In the VBE: **File → Import File**, once per module. They are `.txt` because
that is what the VBE exports; import them as they are.

| File | Becomes |
|---|---|
| `modPNL_28th_Aug.txt` | `modPNL` |
| `modAccess.txt` | `modAccess` |
| `Dashboard.txt` | `modDashboard` |
| `Economic_formula_library.txt` | `modEconFormulas` |
| `ImpRepo.txt` | `modImpRepo` |

**Remove any older copy first.** Importing beside an existing `modPNL` gives you
`modPNL1` and an ambiguous-name error at compile time that names neither.

**Do not open these files in an editor that rewrites line endings.** They must be
pure CRLF; a bare LF makes the VBE join or truncate lines and the compile fails
somewhere unrelated to the damage. To check:

```
python3 tools/check_line_endings.py
```

---

## Step 4 · Compile

VBE → **Debug → Compile VBAProject**. It must complete with no dialog.

If it stops, the usual causes are a duplicated module from Step 3, or a missing
reference. `modAccess` uses late binding throughout (`CreateObject`), so it needs
no reference to ADO or DAO — if something asks for one, a module was imported
twice.

**Check:** `Debug → Compile` is greyed out afterwards.

---

## Step 5 · Lay out the workbook

Run once, from the VBE or Alt+F8:

```
Setup_Workbook_Layout
Access_SetupConfigCells
```

The first creates or repairs every sheet and its headers. The second labels the
Config cells the store owns (`A41:A46`) and sets their defaults.

**Check:** the workbook has `Bonds`, `Futures`, `Swaps`, `OIS_Curves`,
`PNL_Attribution`, `Dashboard`, `Config`, `FutMap`, `SpreadOverride`. Then:

```
python3 tools/check_layout.py     ->  335 column constants across 6 sheets, 0 problems
```

---

## Step 6 · Set the Config cells

| Cell | Meaning | Set it to |
|---|---|---|
| `B4` | T-1, the opening snapshot | the prior business date |
| `B5` | T0, the closing snapshot | the as-of date |
| `B6` | `AUTO` derives both from the clock | `AUTO` or `MANUAL` |
| `B41` | **the run database FOLDER** | blank, or a folder path |
| `B42` | autosave every run | `TRUE` |
| `B43` | last saved RunID | *(auto — do not edit)* |
| `B44` | last store message | *(auto — do not edit)* |
| `B45` | runs to keep when trimming | `0` = keep all |
| `B46` | current run stamp | *(auto — do not edit)* |

**`B41` changed meaning.** It used to name a single `.accdb`. It is now the
folder the per-run databases live in. Blank means "beside the workbook". A value
that still names a `.accdb` is read as *that file's folder*, so an existing
configured path keeps working.

### Why one database per run

Each run writes `PNL_Run_<yyyymmdd>_<hhnnss>.accdb`, stamped with the moment the
data was **retrieved**, not the moment it was saved. Three consequences:

- a stored run is immutable — nothing a later run does can edit it;
- the 2 GB per-file ceiling applies to one day, not to the whole history;
- **re-saving the same pull lands on the same file**, so pressing the button
  twice on one retrieval is idempotent rather than producing two databases that
  differ only in when somebody clicked.

`PNL_Runs_Index.accdb` sits beside them with one row per run, so "every run over
this slice in September" is a query rather than a folder scan.

### Also check these exist

Formulas → **Name Manager**. Two workbook names must be defined:

| Name | Without it |
|---|---|
| `FutMapTable` | `Futures!BBG_Ticker` is `#NAME?` and the **entire futures chain stays blank** |
| `SpreadOverrideTable` | per-bond framework overrides are ignored |

`Setup_Workbook_Layout` defines both. If either is missing, run it again.

---

## Step 7 · Test the connection

```
Access_TestConnection
```

**Check:** it names the database path and reports the tables it found, with no
`MISSING`. If the folder is not writable, or `B41` points somewhere that does not
exist, this is where you find out.

---

## Step 8 · Run the day

Three buttons, left to right. `InstallButtons` puts them on the Dashboard.

| | Button | Does |
|---|---|---|
| 1 | `Button1_LoadBonds` | refreshes the OPICS query, **opens the run**, writes the curve and bond formulas, waits for Bloomberg, autosaves |
| 2 | `Button2_LoadHedgesAndAttribute` | reads both coverage books, writes `Futures`/`Swaps`/`PNL_Attribution`, autosaves into the **same** run |
| 3 | `Button3_BuildDashboard` | full rebuild, then builds the Dashboard |

**Button 1 opens the run.** `AccBeginRun forceNew:=True` fires straight after the
query refresh, because that refresh *is* the retrieval. Button 2 joins the run
Button 1 opened, so both autosaves land in one database.

**Wait for each button to finish.** Bloomberg resolves asynchronously and each
button waits per section; a timeout names the section that did not answer.

> Between Buttons 2 and 3, `Application.CalculateFullRebuild` must have run —
> Button 3 does it. A plain `F9` is not enough: `InterpOIS`, `InterpGov`,
> `InterpSwap` and `BondPullToParPrice` read `OIS_Curves` through the object
> model, so Excel has no dependency edge from a curve cell to the bonds that use
> it and leaves every one holding the **previous run's** number.

---

## Step 9 · Check what was stored

`Config!B44` carries the outcome — `Run 12: 380 bonds, 41 swaps, 22 futures`, or
`Unchanged since run 11 - not saved again`.

Then look at the database named in `Config!B46`:

| Table | Should hold |
|---|---|
| `Run` | one row, `RunStatus = OK`, and the four `Scope*` columns filled |
| `Pos_Bond` / `Pos_Swap` / `Pos_Future` | one row per position |
| `Run_Issue` | anything the save refused, and why |

The `Scope*` columns record **which slice of the book this run covers** — branch
`01`, portfolio `PORT`, accounting category `A`, products `OBR`/`SECUR`. Only the
category surfaces in a sheet column (`Bonds!I`, constant on every row); the rest
exist only as filters. Two runs are comparable **only** when
`ScopePortfolio` and `ScopeAcctgCat` match.

To see the history across files:

```
Access_ShowRunHistory
```

---

## Step 10 · Load a past run back

```
Access_LoadRunIntoSheets        ' prompts for a RunID
```

**This overwrites `Bonds`, `Swaps` and `Futures`.** It restores the stored
positions and then rewrites the formulas, so the market columns **recompute at
today's market** — a loaded run is not a frozen snapshot of that day's numbers.

If you need the numbers as they were, that is what the `Fact_*` tables are for,
and nothing fills them yet.

---

## Every day after that

1. Set `Config!B4` and `B5`, or leave `B6 = AUTO`.
2. Button 1, wait.
3. Button 2, wait.
4. Button 3.
5. Glance at `Config!B44` and at the Dashboard's quarantine block.

Both autosaves go to the same run database. Nothing else is needed.

---

## When a step fails

| Symptom | Cause |
|---|---|
| `ADOX.Catalog` errors in Step 1 | engine missing, or the wrong bitness |
| Compile stops on an ambiguous name | a module was imported twice — remove the old copy |
| VBE joins or truncates lines on import | bare LFs; run `tools/check_line_endings.py` |
| `Access_TestConnection` cannot create the file | `B41` folder missing or not writable |
| `Futures!BBG_Ticker` is `#NAME?` | `FutMapTable` is not defined — rerun `Setup_Workbook_Layout` |
| "No bonds came back from the query" | the OPICS query returned nothing; check the connection and the scope filters |
| The load rejects the query table | it validates **hard-coded literals** — header row 3, first column 1, exactly 11 columns. Adding a 12th SELECT means editing those too |
| Hedges load but no bond links | `LinkedISIN` is resolving to ISINs not on `Bonds`; check `Link_Source` |
| Numbers plausible but a day old | something recalculated without `CalculateFullRebuild` |
| Everything blank after a run | Bloomberg was still answering; the per-section wait names the section |
| `Unchanged since run N - not saved again` | the fingerprint matched. **It hashes position keys and row counts only, not values** — a reload where a notional changed is silently not stored |

### Known defects, still open

These are in the code today and none of them announces itself:

- **The EUR OIS curve is seeded one tenor out.** 13 tickers `ESTRON..EESWE30`
  paired by index against 13 tenors `1M..40Y`, so the overnight rate sits on the
  1M node and the 30Y on the 40Y node. Affects a freshly set-up workbook.
- **EUR and GBP Gov use different Bloomberg fields for T0 and T-1**
  (`YLD_YTM_MID` vs `PX_LAST`), so `Delta_g_bp` is built from incommensurable
  quantities for those two currencies.
- **`PNL_Attribution!Bond_DV01_Credit_Spread` is never written** — permanently
  blank, and anything reading it gets nothing silently.
- **`RebuildPNLOnly` uses `Range.Calculate`**, so the curve UDFs keep the
  previous run's values. Use Button 3.
- **`Bonds!J` is Product, labelled Portfolio** everywhere downstream. A naming
  defect only — it is never an aggregation key.

---

## The tooling, for reference

None of it is needed for a daily run. All of it runs on any machine with Python 3.

| Command | What it checks |
|---|---|
| `python3 tools/build_access_db.py --check` | the DDL parses; 16 tables |
| `python3 tools/check_layout.py` | the column contract |
| `python3 tools/check_access_schema.py` | the VBA schema matches `schema.sql` |
| `python3 tools/check_line_endings.py` | pure CRLF |
| `python3 tools/vba_lint.py` | static analysis |
| `python3 tools/field_graph.py` | the dependency graph — 86 fields, 9 waves, no cycles |
| `python3 tools/check_pipeline_config.py` | `config/pipeline.yaml` agrees with the formula library and the graph |
| `python3 tools/gen_column_registry.py --write` | regenerates `Meta_Column` and the fact-table DDL |

[PIPELINE_SPEC.md](PIPELINE_SPEC.md) is the field-level specification behind all
of it; [config/pipeline.yaml](../config/pipeline.yaml) is the wiring.
