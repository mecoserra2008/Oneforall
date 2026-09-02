# The run store

The workbook holds one day. Load tomorrow and today is gone — the bond
population, the swap mapping, the futures coverage, all overwritten in place,
because a sheet is a single mutable value and there is nowhere else for it to
go. Every question that starts *"what did we hold on…"* is therefore
unanswerable, and every answer the workbook gives is unauditable, because the
inputs that produced it no longer exist.

`modStore` gives them somewhere else to go: a folder of CSV files per run,
beside the workbook, keyed so that a run can never contain the same position
twice and can never be confused with another run.

> **This replaces the Access store.** The design is unchanged — the same three
> position keys, the same fingerprint, the same duplicate guards, the same
> issue log — but a run is now a folder of files rather than rows in a
> database. The reason is narrow and practical: the Access engine is a separate
> install that must match Excel's **bitness**, and getting that wrong produces
> an error message that never mentions bitness. Nothing the store does needed a
> database. [ACCESS_ARCHITECTURE.md](ACCESS_ARCHITECTURE.md) is kept as the
> record of *why the shape is the shape*; it is the same shape.

- [What a run holds](#what-a-run-holds)
- [The three guarantees](#the-three-guarantees)
- [The position keys](#the-position-keys)
- [Numbers, and the locale trap](#numbers-and-the-locale-trap)
- [The two buttons](#the-two-buttons)
- [Getting it running](#getting-it-running)
- [What the store refuses, and why](#what-the-store-refuses-and-why)

---

## What a run holds

Beside the workbook:

```
PNL_Runs\
    run_20260831_143000\
        run.csv                 the manifest: dates, counts, fingerprint, who ran it
        Config.csv              the cells the formulas read (B4, B5, B18, B20-B24)
        Bonds.csv               90 columns x N, VALUES not formulas
        Swaps.csv               78 columns x N
        Futures.csv             42 columns x N
        PNL_Attribution.csv     84 columns x N
        OIS_Curves.csv          the tenor grid
        SpreadOverride.csv      manual: ISIN -> framework
        FutMap.csv              manual: OPICS code -> Bloomberg generic
        Run_Issue.csv           what the save refused, and why
    run_20260901_091500\
        ...
```

Every run is kept. Nothing is ever overwritten, because a folder name carries
the second it was written.

**Why these files and not others.** A snapshot has to be enough to reproduce
the run without asking Bloomberg or OPICS anything, and three findings decide
what that means:

* **PNL_Attribution is 100% formulas, all same-row** — a pure function of
  Bonds + Config + Futures + Swaps + SpreadOverride, with no cross-row
  aggregation and no circularity. So it snapshots cleanly as values.
* **The Dashboard reads `Futures!` and `Swaps!` directly.** The book Actual
  total and the whole FX-hedge block come off `Futures!`, not off
  PNL_Attribution. Restoring PNL alone would leave both wrong.
* **`OIS_Curves` is an invisible dependency.** No cell anywhere references it;
  `InterpOIS`, `InterpGov`, `InterpSwap` and `BondPullToParPrice` reach it
  through the object model, and `BondPullToParPrice` runs on every PNL row via
  `Carry_RollToPar`. Omit the curves and that column silently recomputes
  against whatever curve happens to be on the sheet later — no error, no trace.

`SpreadOverride` and `FutMap` go in because they are **manual, single-copy
inputs**: one chooses the attribution chain per bond, the other chooses which
Bloomberg generic each OPICS code resolves to, and therefore chooses the price.

Deliberately excluded: `Coverage support` and `Coverage Total support` (both
self-declared temporary, cleared every run), and `Dashboard` (fully derivable
from what is here).

### `run.csv` is written last

A folder without a `run.csv` is a folder whose write did not finish. The run
list and the duplicate guard both key off the manifest rather than off the
folder existing, so a save interrupted halfway leaves a folder that is ignored
rather than a run that is half true.

---

## The three guarantees

**Runs are separated.** A run is a folder. Two runs cannot interleave, because
neither can write into the other's folder.

**A position cannot appear twice in one run.** Every position gets a key; a key
already seen in this run is refused, and the refusal is written to
`Run_Issue.csv` with the sheet row that caused it. Not silently dropped — a
silently dropped hedge is a bond that looks unhedged.

**The same run cannot be stored twice.** Every run gets a fingerprint: the
FNV-1a hash of every position key in order. Save the same book twice and the
second save is refused by name, so pressing the button twice does not create
two runs that a later reader has to choose between.

---

## The position keys

What "the same position" means. Get it wrong in one direction and two real
positions collapse into one — a hedge silently disappears. Wrong in the other
and one position becomes two — its risk is double-counted.

| Kind | Key | Why |
|---|---|---|
| Bond | `ISIN + Portfolio + AcctgCat` | one ISIN can be held in two portfolios, and in HTC and AFS at once; those are different positions with different accounting |
| Swap | the `Swaps!A` mapping id (`source\|id\|LinkedISIN\|row`) | the sheet already carries what distinguishes a swap: which bond it hedges and whether it is the plain or synthetic alternative |
| Future | `HedgeSource + CoverageSourceRow + ContractCode + LinkedISIN` | **HedgeSource leads**, because the two coverage books number their rows independently — row 163 of Tx Juro is a different position from row 163 of Total |

Components are joined with `Chr(11)`, not `"|"`. `"|"` would be wrong for
swaps, whose id is itself built out of `"|"`: two different mappings could then
produce one key by accident. Nothing in an ISIN, a portfolio code or a contract
code can contain a vertical tab.

A key longer than 200 characters is **refused, not truncated** — truncation
would make two different positions share a key, and a silent merge is worse
than a visible refusal.

---

## Numbers, and the locale trap

The one thing a CSV store gets wrong quietly, and the reason
`StoreNumberText` / `StoreLooksNumeric` exist.

The desk runs Excel in Portuguese, where the decimal separator is a **comma**.
`CStr(1234.56)` is therefore `"1234,56"`; the writer sees a comma and dutifully
quotes the field; the reader hands back the text `"1234,56"`. The file is
well-formed, the round trip "works", and every DV01 in the restored book is
text that sums to zero.

So numbers are written with `Str$` and read with `Val`, both of which are
invariant — they only ever speak `"."`. Nothing goes through the machine's
regional settings in either direction, which also means a run saved on one desk
restores identically on another.

The read side is deliberately **strict**, because the looser rule is worse:
`IsNumeric` accepts `"0012"` and `"1234,56"` alike, so a padded portfolio code
restores as `12` and matches no key at all, and a wrongly written decimal
restores as `123456` — a hundredfold error, in a cell that looks fine.
`tests/test_file_store.bas` pins both.

Two fidelity rules on the restore, both load-bearing on the Dashboard:

* `Row_Valid` must land **numeric** — `DashValidMask` is `--(N(...)=1)`, and
  `N()` of the text `"1"` is `0`, so a Row_Valid restored as text masks every
  row out and the Dashboard reports a book of zero bonds.
* `ISIN`, `Attribution_Status` and `Spread_Framework_Auto` must land **text**.

Sheets are read and written with `.Value2`, never `.Value`, so a date is a
serial number rather than something rendered in the machine's short-date
format — `03/07/2026` means two different days on two desks.

**Encoding.** Files are UTF-8 with a BOM, through `ADODB.Stream` — ADO, which
ships with Windows, not ACE, which does not. CSV headers are the *contract
keys* (`Delta_Y_bp`), never the display labels (`Variação yield bond (bps)`),
so the header row stays ASCII and a downstream tool sees the machine-readable
name. Values carry accented bond names, which is what the BOM is for.

---

## The two buttons

| Button | Does |
|---|---|
| `Store_RunAndSave` | run the framework (`Button2_LoadHedgesAndAttribute`), then write the snapshot |
| `Store_LoadLatestAndBuild` | read the latest run onto the sheets, republish the `Pnl_*` names at the loaded row count, then build the Dashboard |

That is the whole normal working day: one button to run, one to look at it.

`PublishPnlColumnNames` **must** run after a restore, at the loaded row count.
The Dashboard resolves every PNL column through those names, and `DashKeyCol`
raises 9901 without them — or worse, a name still spanning the previous run's
600 rows sums 220 rows of whatever the clear left behind.

The dependency runs **one way**: `modStore` calls into `modPNL`, never the
reverse. `modPNL` compiles on its own, with `modStore` not imported at all.

### The rest

| Macro | Does |
|---|---|
| `Store_TestStore` | write a run folder, read it back, and report what happened — run this first |
| `Store_SaveCurrentRun` | store the sheets as a run, now, without running the framework |
| `Store_LoadRunByName` | put a named past run back on the sheets |
| `Store_ShowRunHistory` | rebuild the `Run_History` sheet — every run, newest first |
| `Store_OpenRunsFolder` | open `PNL_Runs` in Explorer |
| `Store_SetupConfigCells` | label the Config cells the store owns |

---

## Getting it running

There is no installation step. That is the point of the change.

1. Import `modStore.txt` into the VBA project (**File → Import File…**).
   `modPNL`, `modEconFormulas`, `modImpRepo` and `Dashboard` are the required
   four; `modStore` is the fifth and is optional in the sense that nothing else
   calls it.
2. **Save the workbook somewhere first.** The store writes beside the workbook,
   and an unsaved workbook has no path. The store refuses rather than guessing.
3. Run `Store_SetupConfigCells` once, to label `Config!B41:B44`.
4. Run `Store_TestStore`. It writes a folder, reads it back and tells you what
   it found. If this works, everything works.

### The Config cells the store owns

| Cell | Holds |
|---|---|
| `B41` | where the run folders go. Blank means `PNL_Runs` beside the workbook. A bare name is taken as relative to the workbook; a full path or a `\\server\share` UNC is taken as-is |
| `B42` | reserved for auto-save on every run |
| `B43` | the last run folder written or loaded — written by the store, read by nobody |
| `B44` | what the last operation did, in words |

---

## What the store refuses, and why

Every refusal is written to `Run_Issue.csv` in the run folder, with the sheet
row that caused it, and counted in `run.csv`. Nothing is dropped quietly.

| Refused | Because |
|---|---|
| a duplicate position key within one run | two rows claiming to be the same position means one of them is wrong, and picking either silently is guessing |
| a position key over 200 characters | truncating would merge two different positions |
| a second save of an identical book | two runs a later reader would have to choose between |
| a save from an unsaved workbook | there is no "beside the workbook" yet |
| a run folder with no `run.csv` | its write did not finish |

---

## What checks this

| Check | Guards |
|---|---|
| `tools/check_store_contract.py` | that `StoreSnapshotSheets`, `PnlSheetExtent` and the real header row still agree — a sheet listed for snapshot that the extent cannot describe restores **nothing**, silently |
| `tests/run_file_store.py` | the position keys, the fingerprint, the CSV round trip (comma, quote, newline, accent, empty field), the number/locale rules and the folder ordering — 72 assertions, none of which needs a filesystem |
| `tools/vba_lint.py` | that every procedure called exists and is called with the right number of arguments |

Not covered by the tests, because it needs Windows: `ADODB.Stream`, `MkDir`
and `Dir$`. Those fail loudly. The rest would not.
