# The process map

Every process needed to load bonds, load hedges, and calculate — mapped against
the source, for the move to an Access-backed store.

**How this was built.** Twelve agents read the modules; 429 individual claims
were then re-opened and adversarially checked against the files, and 35 were
corrected. The dependency graph in [§2](#2--input-and-output-are-the-same-thing)
is derived mechanically from `docs/FORMULAS.md`, which `tools/dump_formulas.py`
generates from the VBA — so it cannot drift from the code.

Where a claim here disagrees with another document in `docs/`, this one was
checked against source and the other was not. Several are stale; they are named.

- [0 · Three corrections to start from](#0--three-corrections-to-start-from)
- [1 · The line between Excel and Access](#1--the-line-between-excel-and-access)
- [2 · Input and output are the same thing](#2--input-and-output-are-the-same-thing)
- [3 · Load Bonds](#3--load-bonds)
- [4 · Load Hedges](#4--load-hedges)
- [5 · Calculate](#5--calculate)
- [6 · The run database](#6--the-run-database)
- [7 · Defects found](#7--defects-found)
- [8 · Order of work](#8--order-of-work)

---

## 0 · Three corrections to start from

**The VBA does not compute values.** `WritePNLRow`
(`modPNL_28th_Aug.txt:8043-8940`) makes exactly **86 `.formula = "=..."`
assignments and zero `.value =` assignments** per bond row. modPNL is a *formula
generator*: the economics live in Excel formula strings that Excel evaluates,
plus five VBA UDFs. There is no version of this workbook in which the macro
calculates the attribution.

**The only thing that computes in code is `pnlx/`.** A pure-numpy engine,
CSV in / CSV+JSON+XLSX out. Its entire third-party import set is numpy, yaml and
openpyxl — **nothing in it touches OPICS, SQL, ADO or Access.** It reproduces
**70 of the 86** contract columns; of the 16 it lacks, five are VBA self-check
duplicates that need not exist at all.

**The Access store is positions-only.** `modAccess.txt` creates six tables:
`Run`, `Pos_Bond`, `Pos_Swap`, `Pos_Future`, `Run_Issue`, `Meta_Schema`.
`access/schema.sql` declares **ten more that no code creates, reads or writes**:

```
Instrument   Portfolio   CoverageRelation   Raw_CoverageRow   Raw_BloombergPoint
Curve_Point  Fact_BondRisk   Fact_HedgePosition   Fact_PnlAttribution   Meta_Column
```

The computed-fact tables this migration needs **already exist as DDL**. Nothing
fills them. A run loaded back from the store today is *recomputed* from live
market data — yesterday's PnL cannot be reproduced from the database.

---

## 1 · The line between Excel and Access

`grep` over all three modules finds **no BLPAPI, no COM, no Bloomberg object** —
`BDP`, `BDH` and `BQL` are worksheet functions and nothing else. Excel is
therefore mandatory as the Bloomberg acquisition layer. That single fact fixes
the architecture, and it puts the line exactly where `tools/classify_columns.py`
already computes it across all 330 columns:

| | Columns | Moves to Access? |
|---|---|---|
| Bloomberg | 56 | **No** — staging cells, values read out |
| OPICS query | 56 | No — a query table, not formulas |
| Derived | 139 | Yes |
| UDF | 46 | Mappings become reference tables; numerics stay procedural |
| Aggregate | 14 | Yes — per-bond `SUMIFS` becomes one indexed `GROUP BY` |
| Unwritten | 19 | Resolve before freezing the schema |

**112 fetched, 199 computed.** `PNL_Attribution` has **zero** fetched columns —
all 86 derive from `Bonds`, `Futures`, `Swaps` and `Config`. That is why it is
the sheet to serve from the store first.

The split, by responsibility:

| Excel | Access |
|---|---|
| Bloomberg retrieval (staging sheet) | the store, and history |
| Bond analytics — DV01, convexity, accrued, yield solve, pull-to-par | joins across bonds, hedges, curves, prior day |
| Curve interpolation and other array work | hedge DV01 roll-up per bond |
| Inspection and controlled overrides | framework classification, spread PnL |
| The Dashboard | reconciliation, aggregates, audit, validation flags |

---

## 2 · Input and output are the same thing

**A field is an output of one wave and an input to the next.** Labelling a
column "input" or "output" mislabels most of the sheet. The contract has to be a
dependency graph.

Derived from `docs/FORMULAS.md`: **9 waves over 86 columns, no cycles.**

| Wave | n | Contains |
|---|---|---|
| 0 | 26 | identity, every `Delta_*_bp`, `Bond_DV01_Opening`, `PnL_FX` |
| 1 | 23 | `YearFrac`, the four hedge DV01 roll-ups, single-leg spread PnLs, match counts |
| 2 | 10 | `Actual_Hedge_DV01`, `Target_Hedge_DV01`, `Carry_Coupon`, model PnLs, **`Spread_Framework_Auto`** |
| 3 | 16 | `PnL_Duration_Total`, `SpreadPnL_Used`, `Carry_RollToPar`, ratios, residual checks |
| 4 | 4 | `Carry_Total`, `Hedge_Model_Residual_PnL`, **`Official_Total_PnL`**, `Duration_Identity_Check` |
| 5 | 3 | **`Total_Model_Explained`**, `Hedge_Model_Residual_PnL_Check`, `Row_Exclusion_Reason` |
| 6 | 2 | `Unexplained_Residual_PnL`, `Row_Valid` |
| 7 | 1 | `Unexplained_Residual_Pct` |
| 8 | 1 | `Attribution_Status` |

The fields that are most emphatically both:

| Field | Produced in wave | Feeds |
|---|---|---|
| `ISIN` | 0 | **19** columns |
| **`Bond_DV01_Opening`** | 0 | **12** — every duration and every spread leg |
| `Spread_Framework_Auto` | 2 | 6 |
| `FuturesRTJ_DV01`, `FuturesRT_DV01` | 1 | 6 each |
| `PlainSwap_DV01` | 1 | 5 |
| `Actual_Futures_PnL`, `Actual_PlainSwap_PnL`, `Futures_Match_Count`, `PlainSwap_Match_Count` | 1–2 | 5 each |

### The naming rule

Never `input_` or `output_`. Name a field for **what it is**, and carry
`ProducedInWave`, `ProducedBy` and `ConsumedBy` as metadata in `Meta_Column`.
"Is this an input?" is then answered per-consumer from the edge list rather than
guessed from a prefix.

### Wave gating

Each wave is a table write followed by a completeness check, and the next wave
does not start until it passes:

1. row count equals the bond count for that run;
2. every column the wave owns is non-null, except where a documented guard
   permits null (`ISNUMBER` guards become `IS NOT NULL` — a database has nulls,
   a spreadsheet has `""`, and this gets *simpler* in SQL);
3. failures are written to `Run_Issue` with the wave number and the field;
4. a failed wave stops the run rather than letting wave *n+1* consume a hole.

---

## 3 · Load Bonds

Entry: `Button1_LoadBonds` (`modPNL_28th_Aug.txt:3803`).

The load itself is **one call**:

```vb
lo.QueryTable.Refresh BackgroundQuery:=False      ' :4195
```

against `wsBnd.ListObjects(1)`. No ADO, no recordset, no cell addressing. The
SQL lives in the workbook's Power Query definition and **is not in this
repository** — no `Sql.Database`, `LOKI` or `OPICSMAIN` string exists in any
tracked file. `GetBondSelectQuery` (`:10332`) *looks* like the bond SQL and is
**dead code with zero callers**, returning 16 aliases rather than 11.

Guards before the refresh (`:4173-4193`) — all **hard-coded literals**, not the
layout constants: a ListObject exists; header row `= 3`; starts in column `1`;
exactly `11` ListColumns.

| # | Step | Where | Produces |
|---|---|---|---|
| B1 | Validate query-table geometry | `:4173` | raises 1000–1003 |
| B2 | Refresh the query in place | `:4195` | `Bonds!A4:K{n}` |
| B3 | Measure rows, `End(xlUp)` on col A | `:10871` | `lastBondRow` |
| B4 | **Curves first** | `:4066` | `OIS_Curves` 7–19 |
| B5a | 10 Bloomberg id candidates from the ISIN | `:5777` | `BT:CC` |
| B5b | Resolve the ticker | `:10138` | `AO`, `AP` |
| B5c | T0 BDP block | `:6095` | `Q,S,T,U,V,W,Y,Z,AQ,AT,AV:AZ,BN,BO,CJ:CL` |
| B5d | T-1 BQL block | `:7078` | `R,AA:AE,AR,BP` |
| B5e | *(superseded — both writes overwritten by B5f)* | `:7040` | — |
| B5f | All derived / UDF columns | `:6310` | `L:P, AF:AN, BB:BM, BQ, CM` |
| B5g | Convexity by bump-and-reprice | `:11821` | `CD:CI`, then `X` |
| B6 | Ask Bloomberg, wait per section, `CalculateFullRebuild` | `:9956` | — |
| B7 | Persist the run | `Access_AutoSaveAfterLoad` | `Pos_Bond` |

**B4 is an ordering constraint, not a preference.** Every bond spread column is
an `InterpOIS`/`InterpGov`/`InterpSwap` call against `OIS_Curves`. Against an
empty curve sheet those return blank rather than wrong — "no PnL" instead of
"a plausible wrong PnL".

**B5a → B5b → B5c/B5d is also mandatory.** `BBGParseFormulaR1C1` (`:5817`) nests
`BDP(RC72..RC81, "PARSEKYABLE_DES")` over exactly the candidate columns, and
every market pull keys off the resolved ticker.

### New step B8 — publish the fetched fields into Access

The query stays a Power Query connection, editable in Excel. Its 11 fields are
then **written into the run database** and everything downstream reads the
store. The sheet is staging; it is not the record.

---

## 4 · Load Hedges

Entry: `LoadHedges_Step2` (`:4241`), via `Button2_LoadHedgesAndAttribute`.

Two desk books, **different layouts**, copied values-only into scratch sheets,
mapped, then materialised onto `Swaps` and `Futures`.

| # | Step | Where |
|---|---|---|
| H1 | Clear `Bonds!BR:BS` (the two link columns) | `:4368` |
| H2 | Copy book 1 → `Coverage support` | `:12551` |
| H3 | Build `SwapMap` from it | `:13036` |
| H4 | Build `CoverageFutures` from it | `:13160` |
| H5 | Discard `Coverage support` | `:2733` |
| H6 | Copy book 2 → `Coverage Total support` | `:12669` |
| H7 | **Append** to `CoverageFutures` | `:12733` |
| H8 | Discard `Coverage Total support` | `:12650` |
| H9 | Materialise `Swaps` rows from `SwapMap` | `:3030` |
| H10 | Link swap → bond (**PLAIN only**) | `:10953` |
| H11 | Materialise `Futures` rows from `CoverageFutures` | `:3336` |
| H12 | Link future → bond | `:10990` |
| H13 | OPICS swap query → PayFixed candidates | `:10380` |
| H14 | PayFixed pass 1 — PLAIN | `:4638` |
| H15 | PayFixed pass 2 — SYNTHETIC inherit from PLAIN | `:4828` |
| H16 | OPICS futures enrichment (fill-if-blank only) | `:3481` |
| H17 | Report counts to `Config` | `:4964` |
| H18 | Write the formula-owned `Futures` columns | `:4108` |
| H19 | Write the formula-owned `Swaps` columns | `:4129` |
| H20 | `PNL_Attribution` consumes the rows by `SUMIFS` | `:8409` |
| H21 | Harvest sheet rows for the store | `:13453` |
| H22 | Persist as an Access run | `modAccess.txt:1188` |

### The two source books

| | Hedge_Risco **Tx Juro** | Hedge_Risco **Total** |
|---|---|---|
| Path const | `:180` | `:217` |
| Sheet | `Resumo` | `Resumo` |
| Rows | 9 – 220 | 8 – 34 |
| Columns copied | `A:AI` | `A:W` |
| Column A formula stashed in | `AJ` | `X` |
| Portfolio filter column | **L** | **M** |
| Tagged | `Hedge_Source = "RTJ"` | `Hedge_Source = "RT"` |
| Column map | **hard-coded string literals inside two functions** | named constants |

Both ranges are scanned **in full** — the blocks are not contiguous, so stopping
at the first blank row would silently drop everything below it.

The Tx Juro book is the higher-volume one (212 candidate rows against 27) and is
the one with **no auditable column map**. A silent layout change there produces
zero rows or wrong-field rows with no error.

### What is loader-owned versus formula-owned

`Swaps` — loader writes `A, B, C, G, H, J, L, M, AC, AG, AH, AI, AJ, AK, AL,
AM, AN, AO, AP`, and `I` later from OPICS. Everything else is a formula:
`D, E, F, K, N:AB, AQ:BO`.

`Futures` — loader writes `A, C, D, I, J, AG:AP`, plus `B, E, F, G, H, K` from
OPICS **only where the cell is blank**. Everything else is a formula:
`L:AF`.

### The `#RC` coverage relation

`docs/ACCESS_ARCHITECTURE.md:189-193` says the relation is "reconstructed by a
forward-fill on every run". **It is not.** There is no normalisation, no
forward-fill and no `Link_Source` provenance anywhere in this branch — the
relation is read cell by cell. Any coverage group whose continuation rows carry
a blank `#RC` or a blank ISIN never resolves to its bond.

Building that is **new behaviour here, not a migration.** (It exists in the
other branch's `Pnl_Explain` and can be lifted.)

### New: capture the books raw

The desk overwrites both files in place, so a run is the only chance to record
what they said. `access/schema.sql:244` already declares `Raw_CoverageRow`
(`RunID, SourceBook, SourceRow, RCKeyRaw, FormulaA, ColValues`). Fill it
**before** mapping.

---

## 5 · Calculate

### Where each wave runs

| Wave | Computes | Engine |
|---|---|---|
| C0 | curve derivatives `g = Gov − OIS`, `q = Swap − Gov` | Access, once `Curve_Point` is filled |
| C1 | bond analytics — DV01, convexity, accrued, pull-to-par, interpolation | **Excel** |
| C2 | hedge economics — futures DV01 (CF × CTD), swap DV01 | Excel now, Access later |
| C3 | **hedge roll-up per bond** | **Access** — move this first |
| C4 | the attribution chain (waves 0–7 of §2) | **Access** |
| C5 | quarantine and status — `Row_Valid`, `Attribution_Status` | **Access** |
| C6 | book aggregates for the Dashboard | **Access** |

**C3 is the one to move first.** It is 14 aggregate columns computed today as 15
whole-column `SUMIFS`/`COUNTIFS` **per bond row**. In Access it is a single
indexed `GROUP BY` over the hedge table. Biggest win, smallest blast radius, and
independently checkable against the current `SUMIFS` result.

C4 and C5 must follow C3: `Hedge_DV01_Gap`, `Hedge_Ratio` and `Hedge_Efficiency`
all read the rolled-up hedge DV01, and `Row_Valid` reads the whole chain.

### UDFs that stop being code

`BloombergDayCountToDCC`, `ExcelPriceBasisFromBondDCC`, `CouponFreqNum`,
`SwapFloatFamily`, `SwapFloatTenor` are **lookup tables pretending to be
functions**. In Access they become reference tables and the call becomes a join
— mapping you can see, correct and version.

The numerics stay procedural: `InterpOIS`/`Gov`/`Swap`, `BondPullToParPrice`,
`AccruedInterest`, `NextCouponDate`, `ImpliedRepoBloomberg`. They are already
split into sheet-facing wrappers and pure cores, which is the shape that lets
them live either side of the line.

### The curve sheet is not laid out as commonly described

| Block | Often stated | **Actual (`:702-747`)** |
|---|---|---|
| EUR | `A:P` | `A:P` ✓ |
| USD | `R:AB` | **`S:AG`** |
| GBP | `AI:AX` | **`AJ:AX`** |

`Q/R` and `AH/AI` are declared spacers (`tools/check_layout.py:53`). Header row
**4**, data rows **7–19** — both confirmed. No constant anywhere is named
`Swap_Gov_spread`; the columns are `<CCY>_g_TM1` / `_g_T0` / `_q_TM1` / `_q_T0`.

A reader built to the `R:AB` spec reads the spacer column and the wrong half of
each block.

---

## 6 · The run database

**One `.accdb` per run, named for the retrieval timestamp** —
`PNL_Run_YYYYMMDD_HHMMSS.accdb`.

Keep the `Run` table *inside* each file, holding one row, so `Run_Issue` still
has a parent and a later history step can attach and `UNION` across files. Add a
small index database recording `{timestamp, path, AsOfT0, status, fingerprint}`.

> **Trade-off worth stating.** Per-run files sidestep Access's 2 GB ceiling and
> give an immutable audit trail. The cost is that cross-run history becomes an
> attach-and-union over many files rather than one query; the index database is
> what keeps that tractable.

Tables to activate, all already declared in `access/schema.sql`:
`Raw_CoverageRow`, `Raw_BloombergPoint`, `Curve_Point`, `Fact_BondRisk`,
`Fact_HedgePosition`, `Fact_PnlAttribution`, `Meta_Column`, and `Instrument` as
the ticker cache.

**`Instrument` as a cache is worth its own note.** Today the ISIN → Bloomberg
ticker resolution costs up to ten `BDP(..., "PARSEKYABLE_DES")` probes **per
bond per run**, and the answer essentially never changes. Resolved once and
stored, a run probes only for ISINs the database has never seen — turning a
per-run cost into a per-new-instrument cost, and turning a silently recurring
`"UNKNOWN"` into a row with a null ticker and a date.

---

## 7 · Defects found

Verified against source. Ordered by consequence.

### These move money

| Defect | Evidence | Consequence |
|---|---|---|
| **Hedge_Risco Total notional is consumed as a contract count.** Column `P` (notional) is loaded into the field used as `Contracts` | `:240`, `:211`, `AppendCoverageFuturesMapFromCoverageTotal` | Every RT futures row carries a DV01 inflated by roughly the contract face value — €50m notional treated as 50m contracts. `PNL_Attribution!Q` is wrong for every bond hedged out of the Total book. |
| **RT swaps are recognised, tagged, then silently discarded.** `IsCoverageSwapTypeLabel` forces `futureCode = ""`; `If Len(futureCode) > 0` then skips the row | `:12800`, `:3426`, `:13060` | Any IRS booked in the Total file carries zero risk and zero PnL anywhere in the model — and the status text written to explain the skip is never read. |
| **EUR and GBP Gov use different Bloomberg fields for T0 and T-1** — `YLD_YTM_MID` for T0, `PX_LAST` via BQL for T-1 | `:5331`, `:5400` vs `:9478` | `g_T0` and `g_T-1` are computed from incommensurable quantities, so `Delta_g_bp` and every leg built on it is wrong for EUR and GBP. |
| **`RebuildPNLOnly` uses `Range.Calculate`, not `CalculateFullRebuild`** | `:7586` vs the rule at `:9952` | The curve UDFs keep the **previous run's** values. Every `Delta_r` / `Delta_Gov` / `Delta_Swap` is stale — silently, and plausibly. |
| **Coupon unit ambiguity** — `COUPRATE_8 AS Coupon_Dec` reads as decimal, but every consumer of `Bonds!D` divides by 100 | `:11878`, `:8250`, `pnlx/instruments.py:123` | If it really is decimal, convexity and both pull-to-par carry legs are wrong by 100×. **Settle this against live data before anything else.** |
| **Seeded EUR ESTR tickers are offset one row against the seeded tenors** — 13 tickers ON…30Y paired by index against 13 years 1M…40Y | `:12015` vs `:12019`, paired at `:12040` | On a freshly set-up workbook the whole EUR OIS curve is mislabelled in tenor. |

### These corrupt meaning quietly

| Defect | Evidence |
|---|---|
| **`Product` is stored as `Portfolio`.** SELECT alias 10 is `sm.PRODUCT AS Product`; `Bonds!J` is `BCOL_PORTFOLIO`, and it is the Dashboard's grouping dimension. Only the column *count* is ever checked, never the names | `:361`, `:8073`, `schema.sql:133`, `Dashboard.txt:656` |
| **`PNL_Attribution` is 86 columns (`A..CH`), not 85.** `Risk_Timing_Bias = CG`, `Coupon_Paid_EUR = CH` | `:547-562`, 86 `AddPnlCol` calls |
| **Column `M Bond_DV01_Credit_Spread` is declared, headered, published as a name, and never written** | `:473`, `:7691` |
| **The duplicate-run guard hashes only position keys and row counts, never values.** A reload where a notional changed is silently *not stored* | `modAccess.txt:1330` |
| `Futures!AM CoverageInfo_D` holds different things depending on the source book | `:13207` vs `:12818` |
| `Swaps!AP Notional_Source` is written twice; the second write destroys the provenance, storing the constant `"COVERAGE_SUPPORT"` for every row of every run | `:3321` then `:6778` |
| `Coverage_BPV` is loaded for RT only, then never used — the desk's own BPV, which would fix the count bug above, sits unread | `:12826` |
| OPICS `DelivDate` is fetched then unconditionally overwritten by a Bloomberg formula | `:3597` vs `:6616` |

### Latent and cosmetic

| Defect | Evidence |
|---|---|
| `PNL_LAST_COL = "CI"` but the layout ends at `CH` | `:283` vs `:562` |
| `Actual_Hedge_DV01` is written twice with byte-identical formulas | `:8693`, `:8756` |
| Nothing clears `Bonds!L:CM` when the query returns fewer rows | no clear in `:4153-4223` |
| The load's guards are literals, so `BONDS_QUERY_LAST_COL` is not enforced | `:4180`, `:4190` |
| `Bonds!H` (the AvgCost slot) is loaded and read by nothing | constants jump `G`→`I` at `:359` |
| `tools/build_access_db.py`, named in `schema.sql:5` as the way to apply the DDL, **does not exist** | — |
| `AccIndexSpecs` omits `IX_Run_AsOf_Started`, which `schema.sql:47` creates | `:263` |
| `AccSwapPositionKey` does not upper-case; the bond and future keys do | `:598` vs `:590`, `:620` |
| `Run.SupersedesRunID` is declared and filtered on, never written | `:154`, `:1398` |
| `alter_fact_pnl.sql` emits 84 ALTERs where the docs claim 86 measures; the true total is 87 columns | `:7-90` |
| `Meta_Column` is missing `Bonds!D` — the generator's regex needs exactly one space before `As String` | `gen_column_registry.py:109` |
| Every navigational comment in `WritePNLRow` cites a superseded layout | `:8206`, `:8398`, `:8628` |

### Documents that are stale

- **`docs/analysis/column_sources.csv`** — omits `Bonds!D`, mislabels the whole
  T-1 market block `UNWRITTEN`, and misses `Futures!Q`. **Regenerate before
  using it as the column contract.**
- **`docs/ACCESS_ARCHITECTURE.md`** — claims a `#RC` forward-fill that does not
  exist; lists `Raw_OpicsBond` / `Raw_OpicsHedge`, which the schema does not
  have; its fact-table key `(RunID, ISIN)` is the exact bug the schema comment
  says it fixed.
- **`docs/MIGRATION.md`** — points at a `src/` directory that does not exist.
- The `KNOWN LIMITATIONS` block above `WritePNLRow` still lists "no coupon
  handling"; `Coupon_Paid_EUR` was added and is live.

---

## 8 · Order of work

1. **Settle the two data questions.** Is `COUPRATE_8` decimal or percent? Is
   `Bonds!J` portfolio or product? Both need one look at live OPICS output, and
   everything downstream is wrong by 100× or mis-grouped until they are answered.
2. **Freeze the field contract** in `Meta_Column` — every field with
   `ProducedInWave`, `ProducedBy`, `ConsumedBy`. Named fields, no letters.
3. **Per-run database.** Timestamped filename; create from the full DDL (write
   the missing `tools/build_access_db.py`); add the index database.
4. **Publish the fetched 112** — `Pos_Bond`, `Raw_CoverageRow` (before mapping),
   `Curve_Point`, `Raw_BloombergPoint`. Fix the fingerprint to cover values.
5. **Fix the two money bugs** — RT notional-as-count, discarded RT swaps —
   *before* storing, or the store faithfully records wrong risk.
6. **Move wave C3**, the hedge roll-up, into `Fact_HedgePosition`.
7. **Move C4–C6** into `Fact_PnlAttribution`, wave by wave, each gated.
8. **Point the Dashboard at the store.**

`pnlx/` stays as the independent check: it computes 70 of the 86 columns from
the same inputs, which makes it a second opinion on the Access results rather
than dead code. Two seams would need new work — `loaders.load_inputs` and the
report writer.
