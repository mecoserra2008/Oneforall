# The Access store

The workbook holds one day. Load tomorrow and today is gone — the bond
population, the swap mapping, the futures coverage, all overwritten in place,
because a sheet is a single mutable value and there is nowhere else for it to
go. Every question that starts *"what did we hold on…"* is therefore
unanswerable, and every answer the workbook gives is unauditable, because the
inputs that produced it no longer exist.

`modAccess` gives them somewhere else to go: an Access database, one row per
position per run, keyed so that a run can never contain the same position twice
and can never be confused with another run.

This page is the operating guide. The wider design — computed facts, curves,
Bloomberg points, and the phases that get there — is
[ACCESS_ARCHITECTURE.md](ACCESS_ARCHITECTURE.md).

- [What is stored](#what-is-stored)
- [The three guarantees](#the-three-guarantees)
- [The position keys](#the-position-keys)
- [The run fingerprint](#the-run-fingerprint)
- [Setting it up](#setting-it-up)
- [The buttons](#the-buttons)
- [Loading a stored run back](#loading-a-stored-run-back)
- [How a column move reaches the database](#how-a-column-move-reaches-the-database)
- [What is checked, and by what](#what-is-checked-and-by-what)
- [When something goes wrong](#when-something-goes-wrong)

---

## What is stored

Three tables of positions, one per sheet, plus the run spine and a log of what a
save had to refuse.

| Table | One row per | Fed from |
|---|---|---|
| `Pos_Bond` | bond position | `Bonds!A:K`, the OPICS query's own block |
| `Pos_Swap` | swap **mapping row** | the input columns of `Swaps` |
| `Pos_Future` | futures coverage row | the input columns of `Futures` |
| `Run` | execution | the reporting dates, the user, the counts, the fingerprint |
| `Run_Issue` | thing the save refused | the save itself |
| `Meta_Schema` | schema version applied | `modAccess` on connect |

**Only inputs are stored.** Everything a formula computes is left out on
purpose: it is derived from what is here, so keeping it would be keeping the
same fact twice, and restoring it would overwrite the formula that produces it.
`Fact_PnlAttribution` in [`access/schema.sql`](../access/schema.sql) is where
computed output lands when that phase arrives.

Blank is stored as `Null`, never as `0` and never as `""`. A bond with no book
value has no book value; a stored `0` would read as a position worth nothing.

---

## The three guarantees

**1 · Run separation.** Every stored row carries `RunID`. Runs never edit each
other — a save inserts and does nothing else. History is a property of the
schema rather than a discipline somebody has to keep.

**2 · No duplicate position within a run.** Each position gets a `PositionKey`,
and `(RunID, PositionKey)` is the primary key of all three tables. Access
refuses the second insert: the guard is in the database, not only in the loop
that writes it. A refused row is not silently dropped — it goes to `Run_Issue`
with its sheet row and the row it duplicated.

There are deliberately **two** duplicate guards. A dictionary catches a repeat
before it reaches the database, which is what lets the issue log say *which*
row was refused and where it sat. The primary key catches anything the
dictionary's idea of sameness missed. The second one cannot be bypassed.

**3 · No duplicate run.** Saving an unchanged workbook twice would otherwise
produce two runs holding identical positions. Each run carries a fingerprint —
a hash over every position key it holds — and a save whose fingerprint already
exists for the same reporting date is refused, returning the existing `RunID`
instead of writing a second copy.

---

## The position keys

These decide what *"the same position"* means, so they are pure functions with
no sheet in sight, and they are the most heavily tested thing in the module
(`tests/test_access_store.bas`).

### Bonds — `ISIN · Portfolio · AcctgCat`

`AcctgCat` is in the key because one ISIN can be held in two accounting
categories in the same portfolio, and they are two positions with two book
values. Keying on `(ISIN, Portfolio)` alone would keep whichever arrived first
and refuse the other as a duplicate.

### Swaps — the sheet's own mapping-row id

`Swaps!A` already holds a unique id, built by `AppendSwapMapRowsToSwaps` out of
`source | id | LinkedISIN | row`, precisely so that one swap hedging two bonds
is two rows and the plain and synthetic alternatives on one coverage row never
collide. Re-deriving a key here would be a second opinion on a question the
sheet has already answered.

### Futures — `HedgeSource · Coverage_SourceRow · ContractCode · LinkedISIN`

`HedgeSource` leads because it is what scopes the rest. **The two Hedge Risco
books number their coverage rows independently**, so row 163 of Tx Juro and row
163 of Total are different positions. Drop the book from the key and one of the
two is discarded as a duplicate of the other — silently, and the bond it hedged
reads as unhedged.

### The separator

Components are joined with `vbVerticalTab` (Chr 11), not `|`. The swap mapping
id is itself built out of `|`, so a `|` separator would let two different
mappings produce one key by accident. And joining without a separator at all
would let `("AB","C")` and `("A","BC")` collide — there is a test for exactly
that.

---

## The run fingerprint

FNV-1a over the run's position keys — one 32-bit word per kind, rendered as
eight hex digits, plus a fourth over the three counts, joined by `-`. Two saves
of an unchanged book produce the same string and the second is refused.

The counts are in it as well as the keys, so a run that *lost* a position cannot
fingerprint the same as one that never had it.

It is computed in `Double` arithmetic rather than `Long`. VBA's `Long` is
signed, 32-bit multiplication overflows it on the second character, and VBA's
`Mod` coerces to `Long` and would overflow too. Every intermediate stays under
2⁵³, where a `Double` holds integers exactly. The test pins it against the
published FNV-1a values (`AccHash32("foobar") = "bf9cf968"`), so a future
"tidy-up" of that arithmetic cannot quietly change what counts as the same run.

---

## Setting it up

**[SETUP_ACCESS_STORE.md](SETUP_ACCESS_STORE.md) is the step-by-step version**,
with what to check after each step and what each failure means. The short form:

Nothing to install beyond the engine, and nothing to create by hand.

1. **Import `modAccess.txt`** into the VBE alongside the other four modules
   (File → Import File…), renaming it `modAccess.bas` first so the module takes
   that name.
2. **Run `Setup_Workbook_Layout` once.** It now calls
   `Access_SetupConfigCells`, which labels `Config!A41:A45` and sets the
   defaults.
3. **Press `Access_TestConnection`.** The database file is created if it is not
   there, every table and index is created if missing, and the result box says
   what is now in it.

There is no `.sql` to run and no Access application needed — only the **Access
Database Engine**, which must match Excel's bitness. That is the one setup
failure worth knowing about, and it is [covered
below](#when-something-goes-wrong).

### The Config cells

| Cell | Meaning | Default |
|---|---|---|
| `B41` | database path | blank → `PNL_Data.accdb` beside the workbook |
| `B42` | autosave every run | `TRUE` |
| `B43` | last saved `RunID` | written by the module |
| `B44` | last store message | written by the module |
| `B45` | runs to keep when trimming | `0` = keep everything |

A relative path in `B41` is resolved against the workbook's folder, so a bare
file name works and a UNC path to a shared folder works.

---

## The buttons

| Macro | Does |
|---|---|
| `Access_TestConnection` | creates the database and any missing objects, then reports what is there |
| `Access_SaveCurrentRun` | stores the three sheets as a new run |
| `Access_ShowRunHistory` | rebuilds the `Run_History` sheet: every run, newest first |
| `Access_LoadRunIntoSheets` | puts a stored run back on Bonds, Swaps and Futures |
| `Access_PurgeRun` | deletes one run, after naming what it holds |
| `Access_PurgeOldRuns` | keeps the most recent `Config!B45` runs, deletes the rest |

**Saving is automatic.** Button 1 and Button 2 each call
`Access_AutoSaveAfterLoad` when they finish. It is silent and cannot fail the
button that called it — a database that is unreachable must not lose a load that
has already succeeded; `Config!B44` carries what happened either way.

Button 1 saves too, not only Button 2, because a bond load that is never
followed by a hedge load is still a state somebody may need back. The Button 2
save minutes later supersedes it rather than doubling it, because the
fingerprint covers all three sheets.

Set `Config!B42` to `FALSE` to save only when asked.

### The order a save happens in

It is the order of what can still be undone:

1. read all three sheets into memory and build every key
2. fingerprint them, and **stop here** if this exact set is already stored
3. open a transaction, insert the `Run` row, read back its `RunID`
4. insert the positions, and the issues the insert refused
5. commit, and only then write the `RunID` onto `Config`

Nothing reaches the database until the duplicate check has passed, and nothing
is visible to another user until the commit — so a failure halfway leaves no
half-run behind. That is why there is a transaction rather than three
autocommitted loops.

The positions go in through an ADO **recordset**, not `INSERT` statements: no
value is turned into text and back (so no locale question and no escaping
question), the field types come from the table rather than from a guess, and ACE
appends far faster this way than it parses a thousand statements.

Where SQL text *is* built — the `Run` row, the deletes, the issue log — the
literal builders are locale-proof by construction. On a machine with a comma
decimal separator `CStr(1.5)` is `"1,5"`, and `VALUES (1,5)` is a syntactically
valid `INSERT` **into the wrong columns**. It does not fail; it writes the wrong
number and says nothing. `Str$` always emits a dot, and the date literal is
assembled from `Year`/`Month`/`Day` rather than through `Format`.

---

## Loading a stored run back

`Access_LoadRunIntoSheets` writes **only the mapped input columns**. A formula
column is never touched, because everything in it is derived from what is being
written and will recompute. Columns are cleared to the sheet's full previous
extent first, so a smaller restored run cannot leave the tail of a larger one
behind.

The Bonds sheet is the awkward one: `A:K` belongs to the OPICS query's
ListObject, and writing 380 rows into a table sized for 600 leaves 220 rows of a
table that still claims to hold data. So the table is **resized** to the restored
count first. The next query refresh overwrites all of it, which is correct — a
restored run is something you are looking at, not a new source of truth.

Afterwards `PnlRewriteFormulasForRestoredRun` rewrites the curve, bond, futures
and swap formulas for the restored row counts, rebuilds `PNL_Attribution` and
does a `CalculateFullRebuild`. It is Button 2 without step 1: calling Button 2
itself would refetch from Hedge Risco and throw the restored population away.

**What comes back is the positions, not the prices.** Market data is live — BDP
for T0, BQL dated from Config for T-1 — so a restored run is priced at today.
Restoring the prices too is the `Raw_BloombergPoint` phase in
[ACCESS_ARCHITECTURE.md](ACCESS_ARCHITECTURE.md), not this one.

A run stored before a field existed still loads: the missing column is simply
absent from the header the database returns, and the restore leaves that sheet
column blank rather than failing.

---

## How a column move reaches the database

It doesn't have to. `modAccess` does not know where a column is.

The contract is one array in `modPNL`:

```vba
Public Function PnlPositionFieldMap(ByVal kind As String) As Variant
    ...
    "Portfolio|" & BCOL_PORTFOLIO, _
```

Each entry is `FieldName|ColumnLetter`, and the letter comes from the `BCOL_` /
`WCOL_` / `FCOL_` constant rather than a literal. Move a column by editing its
constant, as always, and the store follows it with **no other edit anywhere**.

Adding a stored field is two edits:

1. an entry in `PnlPositionFieldMap` for the sheet it comes from
2. an entry in `AccTableSpec` for the table it goes to

Every existing database picks the new column up on the next connect —
`AccAddMissingColumns` compares the table to the spec and issues the
`ALTER TABLE`. No migration script, no lost history. It creates what is missing
and never drops or retypes anything: a destructive migration is a decision, not
a side effect of opening a workbook.

---

## What is checked, and by what

| Check | Catches |
|---|---|
| `tools/check_access_schema.py` | drift between `access/schema.sql`, `AccTableSpec` and `PnlPositionFieldMap` |
| `tests/run_access_store.py` | the keys, the fingerprint, the SQL literals and the generated DDL — 45 assertions |
| `tools/vba_lint.py` | the usual, now across five modules |

The schema checker exists because there are **three** descriptions of the store
written in three languages, and drift between them does not fail loudly. A field
in the sheet map with no column in the table is refused by the provider at save
time — one row at a time, inside a transaction, on a desk machine. A column in
the table that no sheet field feeds is simply always `Null`, which reads as *"we
never had that position detail"* rather than as a bug. Its rules:

```
A001  a table modAccess creates is not in schema.sql
A002  a column differs between modAccess and schema.sql
A003  a stored field has no column in its table
A004  a Pos_ column nothing feeds
A005  a position table is missing the (RunID, PositionKey) primary key
```

`A005` is the one that matters most: that key *is* both the run separation and
the duplicate guard, and a table that lost it would keep working and quietly
accept duplicates.

What the tests do **not** cover, because they cannot without Windows and ACE:
the provider, the transaction and the recordset append. Those fail loudly; the
pure half would not.

---

## When something goes wrong

**"Provider is not registered on the local machine."** The Access Database
Engine is missing, or — far more often — it is the wrong bitness. 64-bit Excel
cannot load the 32-bit redistributable and the provider's own message says
nothing about bitness at all, which is why `modAccess` reports Excel's bitness
in the failure box unasked. Install the matching engine.

**"These are exactly the positions already stored as run *n*."** The duplicate
guard did its job: the sheets have not changed since that run. Load new
positions, or `Access_PurgeRun` *n* if it should be replaced.

**A save reports issues.** Look at `Run_Issue` for that `RunID`. Each row names
the table, the sheet row and what was refused — a duplicate key (and the row it
duplicated), a row with no identifying fields at all, or a key too long for the
column. The run is still saved; the issues are what did not go into it.

**A key over 200 characters.** Refused rather than truncated, because truncating
would make two different positions share a key — a silent merge is worse than a
visible refusal. If it ever happens, `PositionKey`'s `TEXT(200)` is what needs
raising, in both `AccTableSpec` and `access/schema.sql` (the checker will insist
on both).

**The database is on a slow share.** Connection timeout is 15 s and command
timeout 120 s, deliberately short: a failure that is a message beats one that
looks like a frozen Excel.
