# Setting up the Access store — step by step

> **SUPERSEDED by [RUNBOOK.md](RUNBOOK.md).** This page describes the earlier
> single-database design. Three things below are no longer true:
>
> - `Config!B41` named a **file**; it is now the **folder** the per-run databases
>   live in.
> - There was one `PNL_Data.accdb` for everything; each run now writes its own
>   `PNL_Run_<timestamp>.accdb`, with a `PNL_Runs_Index.accdb` beside it.
> - "There is no `.sql` to run" — there is now. The store is 16 tables and
>   `AccEnsureSchema` creates six of them, so `tools/build_access_db.py` applies
>   the rest.
>
> Kept for the Step 1 bitness diagnosis and the failure table, which are still
> accurate and still the two things most likely to cost you an afternoon.

Twenty minutes, once. After that the store runs itself: Buttons 1 and 2 save
every run without being asked.

Do the steps in order. Each one ends with something you can check, so if
something is wrong you find out at the step that caused it rather than three
steps later.

- [Before you start](#before-you-start)
- [Step 1 · Check the Access engine, and its bitness](#step-1--check-the-access-engine-and-its-bitness)
- [Step 2 · Get the five modules into the VBE](#step-2--get-the-five-modules-into-the-vbe)
- [Step 3 · Compile](#step-3--compile)
- [Step 4 · Run Setup_Workbook_Layout once](#step-4--run-setup_workbook_layout-once)
- [Step 5 · Point it at a database](#step-5--point-it-at-a-database)
- [Step 6 · Press Access_TestConnection](#step-6--press-access_testconnection)
- [Step 7 · Store your first run](#step-7--store-your-first-run)
- [Step 8 · Prove the duplicate guard works](#step-8--prove-the-duplicate-guard-works)
- [Step 9 · Look at the history](#step-9--look-at-the-history)
- [Step 10 · Load a run back](#step-10--load-a-run-back)
- [From then on](#from-then-on)
- [If a step fails](#if-a-step-fails)

---

## Before you start

**Back up the workbook.** Copy the `.xlsm` somewhere before importing anything.
Step 10 overwrites what is on Bonds, Swaps and Futures, and you want to be able
to go back without thinking about it.

Nothing else is needed. There is no `.sql` to run and no Access application to
install — the database creates itself in Step 6.

---

## Step 1 · Check the Access engine, and its bitness

This is the one thing that can fail for a reason nothing else will tell you
about, so it goes first.

**Find Excel's bitness.** In Excel: **File → Account → About Excel**. The
version line ends in `32-bit` or `64-bit`. Write it down.

**Check the engine.** Windows **Settings → Apps → Installed apps**, search for
`Access`. You are looking for either *Microsoft Access* itself or *Microsoft
Access Database Engine*.

| What you see | Do |
|---|---|
| Microsoft Access, same bitness as Excel | nothing — you already have the engine |
| Microsoft Access Database Engine, same bitness as Excel | nothing |
| Neither, or the wrong bitness | install the **Access Database Engine 2016 Redistributable** matching **Excel's** bitness |

> The bitness must match **Excel**, not Windows. 64-bit Windows running 32-bit
> Excel needs the **32-bit** engine. This is the single most common setup
> failure, and the provider's own error message says nothing about bitness at
> all — which is why `modAccess` reports Excel's bitness in the failure box
> unasked.

You can skip ahead and let Step 6 tell you. It will name the problem correctly.

---

## Step 2 · Get the five modules into the VBE

**Import, do not paste.** Pasting into an existing module leaves the old code
underneath, and you get "Ambiguous name detected" on the first thing that is now
declared twice.

### 2a · Rename the files first

The VBE takes the module name from the **filename**, so rename each `.txt` to
the module name with a `.bas` extension before importing:

| Repository file | Rename to |
|---|---|
| `modPNL_28th_Aug.txt` | `modPNL.bas` |
| `Dashboard.txt` | `modDashboard.bas` |
| `Economic_formula_library.txt` | `modEconFormulas.bas` |
| `ImpRepo.txt` | `modImpRepo.bas` |
| `modAccess.txt` | `modAccess.bas` |

### 2b · Remove the old modules

Open the VBE (**Alt+F11**). In the Project Explorer, for each module that is
already there: **right-click → Remove**, and answer **No** when it offers to
export.

**Remove all five together, then import all five together.** Half-new and
half-old does not compile, and the error it gives points at whichever module
happens to be first — not at the mismatch.

### 2c · Import

**File → Import File…**, pick one `.bas`, repeat for all five.

**Check:** the Project Explorer shows exactly `modPNL`, `modDashboard`,
`modEconFormulas`, `modImpRepo`, `modAccess` under *Modules*, and nothing else.

---

## Step 3 · Compile

In the VBE: **Debug → Compile VBAProject**.

**Check:** nothing happens, and the *Compile* menu item goes grey. That is
success — VBA says nothing when a project compiles.

If it stops on a line, read [If a step fails](#if-a-step-fails) before changing
anything. Two of the errors it can give point at a line that is not the problem.

---

## Step 4 · Run Setup_Workbook_Layout once

In the VBE, put the cursor inside `Setup_Workbook_Layout` and press **F5**.
(Or in Excel: **Alt+F8**, pick it, **Run**.)

It creates or repairs every sheet and its headers, and — new — labels the five
Config cells the store uses.

**Check:** on the `Config` sheet, `A41:A45` now read:

```
A41   Access DB path (blank = PNL_Data.accdb beside this workbook)
A42   Autosave every run to Access (TRUE/FALSE)
A43   Last saved RunID (auto - do not edit)
A44   Last store message (auto - do not edit)
A45   Runs to keep when trimming (0 = keep all)
```

and `B42` says `TRUE`, `B45` says `0`.

---

## Step 5 · Point it at a database

**Leave `Config!B41` blank** unless you have a reason not to. Blank means
`PNL_Data.accdb` in the same folder as the workbook, which is what you want:
the two travel together and nobody has to type a path.

Fill it in only if the database should live somewhere else:

| `B41` | Means |
|---|---|
| *(blank)* | `PNL_Data.accdb` beside the workbook |
| `PNL_Data_2026.accdb` | that file, beside the workbook |
| `\\server\share\risk\PNL_Data.accdb` | that exact path |

A path without a drive letter or a `\\` prefix is treated as relative to the
workbook's folder.

> **If you put it on a share**, everyone who runs the workbook writes to the
> same history — which is usually the point. Access handles concurrent writers,
> and each save is one transaction, so two people saving at once produce two
> runs rather than a tangle. Just make sure the folder is writable by all of
> them: Access needs to create a lock file (`.laccdb`) next to the database.

**Set `Config!B42`.** Leave it `TRUE` so every run is stored automatically. Set
it to `FALSE` if you would rather press `Access_SaveCurrentRun` yourself.

---

## Step 6 · Press Access_TestConnection

**Alt+F8 → `Access_TestConnection` → Run.**

This creates the database file if it is not there, creates every missing table
and index, and then reports what is actually in the database. A good answer here
means a save will work — which is the only useful thing a connection test can
say.

**Check:** you get a box like

```
Access store is reachable.

Database: C:\...\PNL_Data.accdb
Excel:    64-bit
Schema:   version 1
Runs:     0

The database did not exist and has just been created.

Tables:
  OK   Run
  OK   Pos_Bond
  OK   Pos_Swap
  OK   Pos_Future
  OK   Run_Issue
  OK   Meta_Schema
```

Six `OK`s, no `MISSING`. And there is now a `PNL_Data.accdb` next to the
workbook.

If instead you get *"Could not open the Access database"*, that is Step 1 — the
box tells you Excel's bitness so you know which engine to install.

---

## Step 7 · Store your first run

Do a normal daily run: **Button 1**, then **Button 2**. Nothing looks different
— the save is silent and happens after each button has finished its own work.

**Check:** `Config!B43` holds a `RunID` (`1` on the first ever run), and
`Config!B44` reads something like

```
2026-08-31 09:14  Run 1: 384 bonds, 212 swaps, 190 futures
```

If `B44` says a save failed, the load still worked — the store never interrupts
a button. Read the message; it names the reason.

> **You can also store a run at any time** with `Access_SaveCurrentRun`, which
> reports what it did in a message box rather than only in `B44`.

---

## Step 8 · Prove the duplicate guard works

Worth doing once, so you trust it.

Press **`Access_SaveCurrentRun`** again straight away, without loading
anything. You should get:

```
These are exactly the positions already stored as run 1 for the same
reporting date, so nothing was saved.
```

That is the fingerprint doing its job: a hash over every position key in the
book plus the three counts. An unchanged book cannot become a second run — which
is also why Button 1 saving and Button 2 saving minutes later is one run's worth
of history rather than two.

---

## Step 9 · Look at the history

**Alt+F8 → `Access_ShowRunHistory` → Run.**

A `Run_History` sheet appears — every run, newest first, with the dates, who ran
it, the three counts, how many issues it hit and its fingerprint.

This sheet is rebuilt from scratch each time and nothing reads it, so you can
delete it, sort it or scribble on it without consequence.

**Check the `Issues` column is 0.** If it is not, open `Run_Issue` in Access for
that `RunID`: each row names the table, the sheet row and what was refused — a
duplicate key (and the row it duplicated), a row with no identifying fields, or
a key too long for its column. The run is still saved; the issues are what did
not go into it.

---

## Step 10 · Load a run back

Do this once, on purpose, so you know what it does before you need it.

**Alt+F8 → `Access_LoadRunIntoSheets` → Run.** It asks for a `RunID`, defaulting
to the last one saved. Enter one and confirm.

It replaces Bonds, Swaps and Futures with that run's positions, then rewrites
every formula for the restored row counts and does a full rebuild.

**What comes back is the positions, not the prices.** Market data is live — BDP
for today, BQL dated from Config for T-1 — so a restored run is priced at today.
Storing the prices too is a later phase.

**To get back to live positions:** press **Button 1** and **Button 2** as
normal. The OPICS query overwrites the restored bonds and Hedge Risco overwrites
the hedges. Nothing about the restore is sticky.

---

## From then on

Nothing. Press the three buttons as you always have; every run is kept.

| When you want to | Run |
|---|---|
| store a run right now, with a report | `Access_SaveCurrentRun` |
| see what is stored | `Access_ShowRunHistory` |
| put a past run on the sheets | `Access_LoadRunIntoSheets` |
| delete one run | `Access_PurgeRun` |
| keep only the most recent *N* | set `Config!B45` to *N*, then `Access_PurgeOldRuns` |
| check the store is healthy | `Access_TestConnection` |

**Back up the `.accdb` the way you back up anything else.** It is the only copy
of the history — the workbook no longer holds it.

**Watch the file size** if you save every day. A run of ~800 positions is
roughly 200 KB, so a year is around 50 MB against an Access limit of 2 GB. There
is room for many years, but `Config!B45` plus `Access_PurgeOldRuns` is there
when you want to trim.

---

## If a step fails

### "Compile error: Constant expression required"

A `Const` is being initialised with something VBA cannot work out at compile
time — `= ChrW$(1)` rather than `= vbVerticalTab`. The line it names is the
right one.

### "Compile error: Statements and labels invalid outside Sub/Function"

A module-level `Const` or `Dim` has ended up **after** the first `Sub` or
`Function`. VBA has a declarations *section*, not merely declarations: every
module-level declaration must precede the first procedure. Move it up to the top
of the module, with the others.

This is easy to do and invisible on review, because appending a new block of
code to the end of a module is the natural way to add a feature and the constant
it needs travels down with it. `tools/vba_lint.py` now reports it as **E011**,
so running the suite catches it before Excel does.

### "Ambiguous name detected: *something*"

Two modules define the same `Public` name — almost always because an old module
was left in place when the new one was imported. Go back to **Step 2b** and
remove all five before importing all five.

### "Sub or Function not defined: Access_SaveCurrentRun"

`modAccess` is not imported. `modPNL` calls into it from Buttons 1 and 2, so the
five modules are one set — importing four of them does not compile.

### "Sub or Function not defined: PnlPositionsForStore"

The reverse: `modAccess` is imported but `modPNL` is the old version, without
the store bridge. Re-import `modPNL` from this branch.

### "Provider is not registered on the local machine"

Step 1. The engine is missing, or it is the wrong bitness — 64-bit Excel cannot
load the 32-bit redistributable. The failure box gives you Excel's bitness;
install the engine that matches it.

### "Could not create the Access database file"

The folder is not writable, or `Config!B41` points somewhere that does not
exist. Try a local folder first to separate a permissions problem from a path
problem.

### A save reports issues, or `Config!B44` mentions them

The run saved; some rows did not. `Run_Issue` in Access, filtered to that
`RunID`, says which and why. See [Step 9](#step-9--look-at-the-history).

### `Config!B44` says the autosave failed and you never saw a box

That is deliberate. The autosave is silent in **both** directions: a database
that is unreachable must not interrupt a load that has already succeeded, and a
machine without the engine would otherwise pop a modal provider error twice a
day about something the buttons do not depend on. Run `Access_TestConnection`
to get the loud version.
