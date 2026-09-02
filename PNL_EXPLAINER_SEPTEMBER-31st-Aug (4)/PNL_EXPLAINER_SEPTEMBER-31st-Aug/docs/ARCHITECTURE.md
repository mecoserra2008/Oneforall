# Architecture

How the workbook is put together, what depends on what, and where to stand when
you want to change something.

- [The shape of it](#the-shape-of-it)
- [The sheets](#the-sheets)
- [The three buttons](#the-three-buttons)
- [Where the numbers come from](#where-the-numbers-come-from)
- [The two interfaces](#the-two-interfaces)
- [The modules](#the-modules)
- [What is checked, and by what](#what-is-checked-and-by-what)
- [Failure modes worth knowing](#failure-modes-worth-knowing)

Related: [BUTTONS.md](BUTTONS.md) · [COLUMNS.md](COLUMNS.md) ·
[DASHBOARD.md](DASHBOARD.md) · [FORMULAS.md](FORMULAS.md) ·
[DECOMPOSITIONS.md](DECOMPOSITIONS.md)

---

## The shape of it

Positions come from OPICS and two hedge coverage books. Market data comes from
Bloomberg, as **live formulas** rather than captured values. Everything in
between is an Excel formula in a cell, so any number on the Dashboard can be
traced back to its inputs by clicking on it.

```mermaid
flowchart LR
    subgraph sources["Sources"]
        OPICS[("OPICS<br/>positions")]
        HRTJ[("Hedge_Risco<br/>Tx Juro")]
        HRT[("Hedge_Risco<br/>Total")]
        BBG[("Bloomberg<br/>BDP + BQL")]
    end

    subgraph sheets["Workbook"]
        BONDS["Bonds"]
        FUT["Futures"]
        SWAPS["Swaps"]
        CURVES["OIS_Curves"]
        PNL["PNL_Attribution<br/><i>one row per bond</i>"]
        DASH["Dashboard"]
    end

    OPICS --> BONDS
    HRTJ --> SWAPS
    HRTJ --> FUT
    HRT --> FUT
    BBG --> CURVES
    BBG --> BONDS
    BBG --> FUT
    BBG --> SWAPS

    CURVES --> BONDS
    BONDS --> PNL
    FUT -- "SUMIFS on LinkedISIN" --> PNL
    SWAPS -- "SUMIFS on LinkedISIN" --> PNL
    PNL --> DASH
```

Two things about that picture are load-bearing.

**PNL_Attribution has one row per bond, and the hedges are aggregated onto it.**
A future or a swap does not get its own attribution row. Its DV01 and its PnL
are summed onto the bond it hedges, by `LinkedISIN`. That is what makes the
bridge close per bond — and it is also why a hedge with a blank `LinkedISIN`
appears on no row at all and is missing from every total. The Dashboard reports
those separately for exactly that reason.

**Nothing is frozen.** The T-1 columns are BQL point-in-time queries dated from
`Config!B4`; the T0 columns are BDP. Re-running reproduces both exactly. There is
no snapshot to capture, no stamp to check, and no un-freeze path to get wrong.

---

## The sheets

| Sheet | Owned by | Holds |
|---|---|---|
| `Config` | you | dates, the base currency, the spread-framework fallback, counts written back by the loaders |
| `Bonds` | query owns `A:K`, macro owns `L:CM` | one row per bond position, its market data and its derived risk |
| `Futures` | macro | one row per futures hedge, from both coverage books |
| `Swaps` | macro | one row per swap hedge |
| `OIS_Curves` | macro | the OIS, government and swap curves, T-1 and T0 |
| `PNL_Attribution` | macro | **the answer** — one row per bond, every leg of the decomposition |
| `Dashboard` | macro | the read-only presentation of `PNL_Attribution` |
| `SpreadOverride` | you | per-ISIN spread-framework overrides |
| `SwapMap`, `CoverageFutures`, `FutMap` | macro | mapping tables built during the hedge load |
| `Coverage support`, `Coverage Total support` | macro | scratch copies of the two coverage books |

The `Bonds` split matters. **The Excel query owns `A:K` and the macro owns
everything from `L` rightwards.** `BONDS_QUERY_LAST_COL` states the boundary and
`LoadOPICS_Bonds` refuses to load if the live query is the wrong width, rather
than misaligning every row. If you add a column to the OPICS query, the boundary
constant has to move with it — see [COLUMNS.md](COLUMNS.md).

---

## The three buttons

```mermaid
flowchart TD
    B1["<b>1 · Load bonds</b>"] --> B1a["refresh the OPICS query · A:K"]
    B1a --> B1b["clear + rewrite the macro columns L:CM<br/>for exactly the rows that came back"]
    B1b --> B1c["write the curve formulas"]
    B1c --> B1d["ask Bloomberg, wait, CalculateFullRebuild"]

    B1d --> B2["<b>2 · Load hedges + attribute</b>"]
    B2 --> B2a["load Futures and Swaps rows<br/>from both coverage books"]
    B2a --> B2b["write the futures formulas<br/><i>and</i> the swap formulas"]
    B2b --> B2c["rebuild PNL_Attribution · one row per bond"]
    B2c --> B2d["publish the column names"]
    B2d --> B2e["ask Bloomberg, wait, CalculateFullRebuild"]

    B2e --> B3["<b>3 · Build dashboard</b>"]
    B3 --> B3a["CalculateFullRebuild"]
    B3a --> B3b["read the published names, draw the sheet"]
```

There is deliberately no "write formulas" button and no "refresh market data"
button. [BUTTONS.md](BUTTONS.md) explains why, and what moved where.

---

## Where the numbers come from

One bond's row on `PNL_Attribution`, end to end:

```mermaid
flowchart TB
    subgraph inputs["Inputs"]
        POS["position<br/>notional, ISIN, portfolio"]
        PX["prices T-1 and T0<br/>dirty, clean, yield"]
        RISK["risk<br/>mod duration, convexity, spread duration"]
        CRV["curves T-1 and T0<br/>OIS · gov · swap"]
        HDG["hedges matched on LinkedISIN<br/>futures BPV+PnL · swap BPV+PnL"]
    end

    POS --> DV01["Bond_DV01_Opening<br/><i>T-1 price, T-1 FX</i>"]
    RISK --> DV01
    PX --> DV01

    PX --> DY["Delta_Y_bp"]
    CRV --> DR["Delta_r · Delta_Gov · Delta_g<br/>Delta_Swap · Delta_q · Delta_i"]
    PX --> DR

    DV01 --> DUR["PnL_Duration_Total<br/><i>chain chosen by the framework</i>"]
    DR --> DUR
    DV01 --> CONV["PnL_Convexity"]
    DY --> CONV

    POS --> CARRY["Carry_Total<br/>coupon accrual + pull to par"]
    CRV --> CARRY
    PX --> FX["PnL_FX"]

    HDG --> HM["Hedge_Curve_Model_PnL"]
    DR --> HM
    HDG --> HB["Hedge_Model_Residual_PnL<br/><i>actual − model</i>"]

    DUR --> EXP["Total_Model_Explained"]
    CONV --> EXP
    CARRY --> EXP
    FX --> EXP
    HM --> EXP
    HB --> EXP

    PX --> OFF["Official_Total_PnL<br/>Δ dirty MV + coupon cash + actual hedge"]
    HDG --> OFF

    EXP --> RES["Unexplained_Residual_PnL"]
    OFF --> RES
```

Every formula in that picture is written out in full in
[FORMULAS.md](FORMULAS.md), generated from the source. The economics behind the
shape are in [DECOMPOSITIONS.md](DECOMPOSITIONS.md).

---

## The two interfaces

Most of the workbook is ordinary: procedures call procedures. Two boundaries are
different, because they are crossed by *data* rather than by a call, and both
have broken in ways that were invisible until the whole thing stopped.

### 1 · The OPICS query boundary, on `Bonds`

The query writes `A:K`. The macro writes `L` onwards. Nothing enforces this at
runtime except `BondsLayoutIsSane`, which is checked on every load and refuses
rather than misaligning.

### 2 · The column contract, between `modPNL` and `modDashboard`

`modDashboard` needs to find `PNL_Attribution`'s columns. It used to find them
by **searching row 4 for the header text**, which quietly made every string in
that row part of the interface:

```mermaid
flowchart LR
    subgraph before["Before · the header row WAS the interface"]
        H4["row 4<br/>“Bond_DV01_Current”"] -->|"Rows(4).Find(text)"| D1["modDashboard"]
        R["retitle to<br/>“Bond BPVs”"] -.->|breaks| H4
        D1 --> X["error 9901<br/>no Dashboard at all"]
    end
```

Retitling `L4` to `Bond BPVs` — correct, obvious and the right name for the desk
to read — made the search return `Nothing` and killed the entire build before it
drew a cell. Nothing in the sheet could tell you that would happen.

```mermaid
flowchart LR
    subgraph after["After · the names are the interface"]
        T["<b>PnlLayout</b><br/>column · key · label"] --> HDR["row 4<br/>the LABEL<br/><i>“Bond BPVs”</i>"]
        T --> NM["workbook names<br/>the KEY<br/><i>Pnl_Bond_DV01_Current</i>"]
        NM --> D2["modDashboard"]
        HDR -.->|"read by nothing"| D2
    end
```

One table declares each column once. The header writer takes the **label** from
it; `PublishPnlColumnNames` publishes the **key** as a workbook name pointing at
the column letter; `modDashboard` uses only the names. Row 4 can now say anything,
in any language, and columns can move.

Full detail and the change procedure: [COLUMNS.md](COLUMNS.md).

---

## The modules

| Module | Exported as | Contains |
|---|---|---|
| `modPNL` | `modPNL_*.txt` | the buttons, the sheet layouts, every formula writer, the PnL model, the curve UDFs |
| `modDashboard` | `Dashboard.txt` | the Dashboard build, and nothing else |
| `modEconFormulas` | `Economic_formula_library.txt` | named builders for the economically meaningful formula strings |
| `modImpRepo` | `ImpRepo.txt` | day counts, accrued interest, coupon schedules, implied repo and basis |

They depend on each other, so importing only some of them will not compile.

The export filenames change (`modPNL_25th_Aug` → `modPNL_26th_Aug` → …). Nothing
in the tooling matches on a filename: `tools/modules.py` identifies each module
by a procedure only it defines, so a rename cannot quietly take the checks
offline — which it had been doing.

### Why `modEconFormulas` exists

A formula that means something economically — a residual, a hedge gap, a
market-value change — is defined **once**, as a builder returning the formula
string, and called wherever that quantity is needed. The alternative is the same
expression typed inline in four places, three of which get updated.

`tests/test_formula_equivalence.bas` pins each builder's output against the
literal the inline code used to produce, character for character, so moving a
formula into the library provably changed nothing. Because the pinned side is
built from the real `PCOL_`/`BCOL_` constants, those cases **also fail when a
column moves** — which is what makes removing a column safe.

---

## What is checked, and by what

There is no Excel here, so the suite works two ways.

```mermaid
flowchart LR
    subgraph static["Static · reads the .bas text"]
        L1["vba_lint.py<br/><i>the compile errors that actually<br/>stop this workbook</i>"]
        L2["check_layout.py<br/><i>the column contract</i>"]
        L3["check_line_endings.py<br/><i>CRLF, no BOM</i>"]
        L4["check_dashboard.py<br/><i>every formula the Dashboard writes,<br/>rendered and checked</i>"]
        L5["check_diagrams.py<br/><i>the documentation diagrams parse</i>"]
    end
    subgraph live["Executed · headless LibreOffice Basic"]
        R1["test_pull_to_par.bas<br/><i>20 assertions on the maths</i>"]
        R2["test_formula_equivalence.bas<br/><i>formula strings, character for character</i>"]
    end
    static --> OK["bash tests/run_all.sh"]
    live --> OK
```

`tools/vba_run.py` lifts pure procedures out of the modules and runs them
through LibreOffice Basic for real. That is why the model code is split into a
sheet-facing wrapper and a numeric core: the core touches no worksheet, so it
can be executed and asserted against. Procedures that touch `Range` or
`Worksheet` cannot be run this way and are covered by the static checks only.

Both linters are self-checked — `tests/test_lint_selftest.py` feeds them one
deliberately broken module per diagnostic, plus a clean one that must stay
silent. A linter that quietly stops detecting things is worse than none.

---

## Failure modes worth knowing

These are the ones that do not announce themselves.

**A stale full rebuild.** `InterpOIS`, `InterpGov`, `InterpSwap` and
`BondPullToParPrice` read `OIS_Curves` through the object model, not through cell
references. Excel therefore has **no dependency edge** from a curve cell to the
bonds that use it, and `Application.Calculate` leaves every one of them holding
the previous run's number — silently, and looking entirely plausible. Only
`Application.CalculateFullRebuild` re-evaluates them. Buttons 1, 2 and 3 all do
it; nothing else should need to.

**Asking Bloomberg and not waiting.** BDP and BQL resolve asynchronously.
Recalculating immediately after asking computes the book against
`#N/A Requesting Data`, and the result is not an error — it is blanks, which read
as a *small* PnL rather than a missing one.

**A hedge with no `LinkedISIN`.** It reaches no bond row, so it is in no total.
The Dashboard's quarantine block reports these; without that they simply vanish.

**Two bonds sharing a hedge ISIN.** Each row's `SUMIFS` claims the *full* hedge
PnL and DV01, so it is counted twice. The Dashboard reports rows sharing an ISIN.

**A column that is declared but never filled.** The name resolves, the `SUMIFS`
runs, and the Dashboard reports a confident zero. `check_layout.py` fails on
this now; the one known case is listed in `KNOWN_EMPTY_COLUMNS`.
