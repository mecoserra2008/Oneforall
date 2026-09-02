# PnL Explainer — the two workflows, step by step

Two things happen in this workbook, and they are completely separate.

| | Workflow | Runs | Writes to |
|---|---|---|---|
| **1** | [Hedge attribution](#part-1--hedge-attribution) | `Load_Hedges` | `Futures`, `Swaps`, `Bonds!SwapLink`/`FutLink` |
| **2** | [Dashboard](#part-2--the-dashboard) | `BuildDashboard_Step6` | `Dashboard` only |

Between them sits `PNL_Attribution`, which **no macro writes any more**. Its
formulas are the desk's. Workflow 1 feeds the cells those formulas read;
workflow 2 reads what those formulas produce.

```
  Hedge_Risco Tx Juro.xlsm ─┐
                            ├─► Load_Hedges ─► Futures + Swaps ─┐
  Hedge_Risco Total.xlsm  ──┘                                   │
                                                                ▼
                                              PNL_Attribution (formulas only)
                                                                │
                                                                ▼
                                              BuildDashboard_Step6 ─► Dashboard
```

---

## Contents

- [Part 1 — Hedge attribution](#part-1--hedge-attribution)
  - [Before you press it](#before-you-press-it)
  - [The nine steps](#the-nine-steps)
  - [What lands in which column](#what-lands-in-which-column)
  - [The #RC join](#the-rc-join-the-thing-that-makes-it-work)
  - [When it goes wrong](#when-it-goes-wrong)
- [Part 2 — The Dashboard](#part-2--the-dashboard)
  - [How it addresses PNL_Attribution](#how-it-addresses-pnl_attribution)
  - [**Two reasons it will not work today**](#two-reasons-it-will-not-work-today)
  - [The 45 names it needs](#the-45-names-it-needs)
  - [What it reads off Futures and Swaps directly](#what-it-reads-off-futures-and-swaps-directly)
  - [The build steps](#the-build-steps)
- [The full running order](#the-full-running-order)

---

# Part 1 — Hedge attribution

**Button: `Load_Hedges`** — the only macro this workbook needs.

Assign it with right-click → *Assign Macro* → `Load_Hedges`. Do **not** look
for `LoadHedges_Step2`: Excel's Assign Macro dialog lists only `Public Sub`s
with **no parameters at all**, and an `Optional` one still disqualifies a
procedure. That is why a button could not be attached before.

Put the button on any sheet **except** `Dashboard` — or, if it must go there,
name the shape `btnPnlLoadHedges`. The Dashboard build deletes every shape on
that sheet whose name does not start with `btnPnl`.

## Before you press it

| Check | Why |
|---|---|
| `Bonds` has rows | The load aborts if `Bonds!A4` is empty. There is nothing to attach a hedge to. |
| Both coverage books are reachable on `M:\P_Pires\TRADING\` | Book 1 unreachable = nothing is changed at all. Book 2 unreachable = partial load, reported. |
| `Futures` row 4 and `Swaps` row 4 are intact | Checked automatically in step 2. Or run `Check_Hedge_Layout` on its own first — it reads two header rows and writes nothing. |

## The nine steps

### 1 · Freeze the application state
`Calculation = Manual`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts` and
`AskToUpdateLinks` all off, and the previous values saved. Every exit path —
success, handled failure, unhandled error — restores them before any message
box appears, so a modal dialog can never be waiting while events are off.

### 2 · Assert the layout — **before anything is read or written**
`AssertHedgeSheetHeaders` reads row 4 of `Futures` and `Swaps` and compares all
**42** and **78** headers against the layout the macro compiles against.
Comparison ignores case and surrounding whitespace, so retitling for
readability is fine; only a genuinely different column fails.

A mismatch **stops the run and changes nothing**, naming the sheet, the column
letter, what was expected and what is actually there.

> This step exists because of what went wrong before. Every column is addressed
> by letter. When those letters drifted from the sheet, the load did not fail —
> it succeeded against the wrong cells. Hedge text landed on market data, and
> the attribution reported a fully hedged book as unhedged. **A moved column
> must produce a message, never a wrong number.**

### 3 · Read book 1 — `Hedge_Risco Tx Juro.xlsm`
Rows 9–220 of `Resumo`, columns A:AI, copied to the hidden `Coverage support`
sheet. The whole range is scanned: the block is **not** contiguous, so stopping
at the first blank row would silently drop everything below it.

**If this book fails, the run stops and nothing is touched.** It owns the swaps
and most of the futures; carrying on would clear `Swaps` and `Futures` and
refill them from an empty map — losing every hedge because a mapped drive
blipped.

### 4 · Build the `#RC` index for book 1
Two passes over the copy:

1. `BuildCoverageRelationByRow` — forward-fill the `#RC` down each group,
   resetting at a genuinely blank row so a fill cannot leak into the next group.
2. `BuildCoverageRelationIndex` — register `#RC → ISIN` from the rows that
   actually carry an ISIN.

### 5 · Build `SwapMap` and `CoverageFutures`
`BuildSwapMapFromCoverageSupport` then
`BuildCoverageFuturesMapFromCoverageSupport`. Both are scratch sheets the macro
owns outright. The futures one also **clears** `CoverageFutures`, which is why
it must run before book 2 appends to it — otherwise book 2's rows pile up under
yesterday's book 1 rows on every run.

### 6 · Read book 2 — `Hedge_Risco Total.xlsm`
Rows 8–34 of `Resumo`, columns A:W, to `Coverage Total support`. Different
layout, its own scan, **its own separate `#RC` index** — `#RC 12` in one book is
not the same relation as `#RC 12` in the other, so merging the indexes would
attach the wrong bond to a hedge whenever both books used the same number.

Failure here is **not** fatal: it becomes a partial load, and the futures count
from each book is reported separately so a book that stopped arriving is
visible rather than netted away inside one total.

### 7 · Write `Swaps`, then `Futures`
Nothing above this point has changed a data sheet. Every early exit therefore
leaves all three exactly as they were.

For each sheet: clear **only the owned columns**, write **only the owned
columns**, then extend the sheet's own formulas onto any new rows.

> **The formula-destruction bug.** Both routines used to clear the full sheet
> width and write back a full-width buffer that was empty in the formula
> columns — wiping every BDP and BQL formula on `Futures` and `Swaps` on every
> run, including `Swaps!O` (`DV01 BBG`) and the whole BQL block. The old
> three-button order rewrote them immediately afterwards. Nothing does now.

Row growth is handled by `FillFormulaColumnsDown`: it reads the **R1C1 form of
the formula already in the column**, from the last row that has one, and
assigns it to the rows below that have none. R1C1 re-bases relative references
per row exactly as a fill-down would, and assigning `.FormulaR1C1` writes the
formula and nothing else — no format travels with it the way it would through
`.Copy`. Rows that already hold a formula are never touched.

### 8 · Fill futures contract detail
`Exchange` and `FaceValue` from the `FutureSpec` table, keyed on the contract
root (`RX`, `OE`, `DU`, `UB`, `OAT`, `IK`, `BTS`, `TY`, `FV`, `US`, `TU`, `WN`,
`UXY`, `G`). Written only into a cell that is **empty**, on a column that is
**not formula-driven** — decided once per column, so a BDP column is out of
bounds for its whole length even where Bloomberg has not answered yet.

`DelivDate`, `CTD_ISIN`, `CTD_CF`, `FUT_VAL_PT`, `CTD_Ticker` and
`HedgeUnitDV01` are **not written** — they are already Bloomberg formulas fed
from `BBG_Ticker = VLOOKUP(ContractCode, FutMapTable, 2, FALSE)`.

To add a contract: one line in `FutureSpec`. An unknown root is not an error —
the row loads, its two spec cells stay empty, and Bloomberg still prices it.

### 9 · Stamp the bond links, report
`Bonds!SwapLink` (**BQ**) and `Bonds!FutLink` (**BR**), accumulated in memory
against a one-pass ISIN index and written once per touched row.

`SwapLink` carries **plain swaps only**. A synthetic is an alternative to the
plain swap on the same relation, not a second hedge, so listing it would read
as double cover. Synthetics stay fully mapped in `SwapMap`, on `Swaps`, and
through `Swaps!LinkedISIN` + `Swaps!Swap_ID_Source`.

The completion message reports rows written per sheet, map rows, each book's
futures count separately, bonds linked, and the two known gaps (below).

## What lands in which column

**Everything not listed is out of bounds and is never touched.**

### `Futures` — 16 owned of 42

| Col | Header | Written |
|---|---|---|
| `A` | ContractCode | contract root from the coverage label |
| `C` | CCY | |
| `D` | Contracts | |
| `I` | Portfolio | the `#RC` |
| `J` | **LinkedISIN** | the covered bond — **the join to PNL_Attribution** |
| `AF` | Status | import status |
| `AG`–`AL` | Coverage_SourceRow … CoverageInfo_D | provenance |
| `AM` | **Hedge_Source** | `RTJ` or `RT` — drives the RTJ/RT split |
| `AN` | Coverage_BPV | |
| `AO` | **Hedge_Class** | `RATES` or `FX` |
| `AP` | Link_Source | how the ISIN was decided |

Filled if blank and non-formula: `B` Exchange, `E` FaceValue.

### `Swaps` — 27 owned of 78

| Col | Header | Written |
|---|---|---|
| `A` | DealID | composite mapping id |
| `B` `C` | Ccy, Notional | |
| `G` `H` | StartDate, EndDate | |
| `I` | PayFixed | **left empty** — see gaps below |
| `J` | Portfolio | the `#RC` |
| `K` | **FloatCurve_Type** | `ESTR` / `SOFR` / `EURIBOR` / `UNKNOWN` |
| `L` | **LinkedISIN** | the covered bond — **the join** |
| `M` | Cpty | |
| `AJ` `AK` `AM` | Status, PnL_Source, AF_Status | |
| `AN`–`AP` | BBG ids | direct / fixed leg / float leg |
| `AQ` | **Swap_ID_Source** | `PLAIN` or `SYNTHETIC` |
| `AR`–`AZ` | CoverageRelation … Notional_Source | mapping block |
| `BZ` | Link_Source | |

### `Bonds` — 2 owned of 90
`BQ` SwapLink, `BR` FutLink.

## The `#RC` join, the thing that makes it work

A coverage relation — the desk writes it `#RC` — ties **one** covered bond to
**one or more** hedge instruments. In both books the bond is named **once**, on
the row that opens the group. The rows underneath carry only the instrument:
their ISIN cell is blank, and in `Hedge Risco Total` their `#RC` cell is blank
too, because the file is read by a human who can see the group.

Taking the bond from each instrument row's **own** ISIN cell works on the first
row of a group and fails on every continuation row — which is then stamped
"Missing LinkedISIN", never linked, and its risk and PnL vanish from the
attribution. The bond then reads as under-hedged and gets the wrong spread
framework.

**The relation is the join.** Resolution order, recorded per row in
`Link_Source`:

| `Link_Source` | Meaning |
|---|---|
| `RC` | from the relation; the row carried no ISIN |
| `RC=ROW` | relation and row agreed |
| `RC<>ROW` | they disagreed — **the relation won** |
| `ROW` | relation unknown; fell back to the row's own ISIN |
| `RC_AMBIGUOUS` | one relation named several ISINs |
| `NONE` | nothing to link to |

Scan `Link_Source` after a load. `RC<>ROW` and `RC_AMBIGUOUS` are the rows
worth a human glance.

## When it goes wrong

| Symptom | Cause |
|---|---|
| "not laid out the way this macro expects" | A column moved on `Futures`/`Swaps`. Fix the header, not the macro. |
| "No bonds loaded" | Load bonds first. |
| "Tx Juro could not be read… sheets are exactly as they were" | Drive or file. Nothing was changed. |
| Futures count fine, swaps 0 | Book 1 read but no `PLAIN` rows resolved — check `Link_Source`. |
| Bonds linked = 0 with rows loaded | `LinkedISIN` resolving to ISINs not on `Bonds`. |
| `Futures!BBG_Ticker` = `#NAME?` | `FutMapTable` is not defined. See below. |

### The two known gaps

**`PayFixed` is not written.** It came only from OPICS, and neither coverage
book has a pay/receive column in the current constants. A guessed sign flips
the hedge, so the cell is left empty and the gap is reported. Name the column
in the Tx Juro book and it gets read.

**`FloatCurve_Type` can be `UNKNOWN`.** Read from the Bloomberg leg ids where
they say (`ESTR`/`EONIA`, `SOFR`, `EURIB`), then `USD → SOFR`, else `UNKNOWN`.
EUR is deliberately **not** defaulted: a plain-vanilla EUR IRS floats on
EURIBOR, not ESTR, and picking wrong changes the attribution silently.
`Swap_Curve_Model_PnL` tests for exactly `"UNKNOWN"` and **blanks** its result,
so an unresolved swap drops its bond out of the headline rather than
contributing a quiet zero. The count is in the completion message.

---

# Part 2 — The Dashboard

**Macro: `BuildDashboard_Step6`** (in `modDashboard`), preceded by
`Application.CalculateFullRebuild`.

The rebuild comes **first**. `InterpOIS`, `InterpGov`, `InterpSwap` and
`BondPullToParPrice` read `OIS_Curves` through the object model, not through
cell references, so Excel has **no dependency edge** from a curve cell to the
bonds that use it. A plain `F9` leaves every one of them holding the previous
run's number — silently, and looking entirely plausible. Only
`CalculateFullRebuild` re-evaluates them.

**Design rule: VBA decides layout, Excel computes every number.** The module
writes formulas, never computed values. So the sheet recalculates when
`PNL_Attribution` changes instead of going stale, `Ctrl+[` from any figure
jumps to the rows behind it, and a change to a PnL formula cannot leave the
Dashboard disagreeing with it. The only values written are text labels and the
source-row indices behind the ranked tables — both layout decisions.

## How it addresses `PNL_Attribution`

**Never by column letter. Never by reading row 4.**

Each column is published as a workbook name `Pnl_<Key>` pointing at that
column's data range, and every formula reads the name:

```
=SUMIFS(Pnl_Official_Total_PnL,Pnl_ISIN,"<>",Pnl_ISIN,"<>TOTAL")
```

Two consequences, and both are the point:

- **Move a column** and nothing in the Dashboard changes — the name is
  republished against the new letter.
- **Rename a header** and nothing changes either. It used to search row 4 for
  the header text, which made row 4 a machine interface wearing the costume of
  a label row: retitling `L4` to "Bond BPVs" for the desk to read — correct,
  obvious, harmless-looking — made the search return `Nothing` and killed the
  whole build on error 9901 before the first cell was drawn.

Geometry: header row **4**, data from row **5**.

## Two reasons it will not work today

### 1 · Nothing publishes the names

`PublishPnlColumnNames` lives in **`modPNL`**, and was called by Button 2.
Under the one-macro rule `modPNL` is not imported, so nothing defines
`Pnl_*` any more.

**Symptom:** the Dashboard is `#NAME?` everywhere, or the build stops on error
**9901** naming the first key it could not resolve.

### 2 · Two of the names would point one column too far left

`modPNL`'s layout table declares **84** columns. Your sheet has **85** — it
carries `Bond_DV01_Opening` at **CE**, which that table omits (its comment even
asserts "the real sheet does not have it"). Everything after CE therefore
shifts:

| Key | `modPNL` publishes | Your sheet | Would actually read |
|---|---|---|---|
| `Risk_Timing_Bias` | `CE` | **`CF`** | `Bond_DV01_Opening` |
| `Coupon_Paid_EUR` | `CF` | **`CG`** | `Risk_Timing_Bias` |

The other **43** of the 45 keys the Dashboard requires are correct.

**So do not run Button 2 to publish the names.** `WritePNLSectionSafe` would
rebuild `PNL_Attribution` to its own 84-column layout, deleting
`Bond_DV01_Opening` — the column your `PnL_OIS`, `PnL_GovBasis`,
`PnL_Credit_Ispread` and whole duration chain reference as `CE5`.

**Define the 45 names by hand instead** (Formulas → Name Manager → New), or ask
for a small publisher macro that writes names only — no formulas, no headers.

## The 45 names it needs

Scope: **Workbook**. Refers to, against your actual layout:

| Name | Col | Refers to |
|---|---|---|
| `Pnl_ISIN` | `A` | `='PNL_Attribution'!$A$5:$A$50004` |
| `Pnl_Name` | `B` | `='PNL_Attribution'!$B$5:$B$50004` |
| `Pnl_CCY` | `C` | `='PNL_Attribution'!$C$5:$C$50004` |
| `Pnl_Portfolio` | `D` | `='PNL_Attribution'!$D$5:$D$50004` |
| `Pnl_Official_Total_PnL` | `BB` | `='PNL_Attribution'!$BB$5:$BB$50004` |
| `Pnl_Total_Model_Explained` | `BA` | `='PNL_Attribution'!$BA$5:$BA$50004` |
| `Pnl_Unexplained_Residual_PnL` | `BC` | `='PNL_Attribution'!$BC$5:$BC$50004` |
| `Pnl_PnL_Duration_Total` | `AG` | `='PNL_Attribution'!$AG$5:$AG$50004` |
| `Pnl_PnL_OIS` | `AH` | `='PNL_Attribution'!$AH$5:$AH$50004` |
| `Pnl_PnL_GovBasis` | `AI` | `='PNL_Attribution'!$AI$5:$AI$50004` |
| `Pnl_PnL_SwapGovBasis` | `AJ` | `='PNL_Attribution'!$AJ$5:$AJ$50004` |
| `Pnl_SpreadPnL_Used` | `AU` | `='PNL_Attribution'!$AU$5:$AU$50004` |
| `Pnl_PnL_Convexity` | `AL` | `='PNL_Attribution'!$AL$5:$AL$50004` |
| `Pnl_Carry_Coupon` | `AM` | `='PNL_Attribution'!$AM$5:$AM$50004` |
| `Pnl_Carry_RollToPar` | `AN` | `='PNL_Attribution'!$AN$5:$AN$50004` |
| `Pnl_Funding_Carry_Memo` | `AO` | `='PNL_Attribution'!$AO$5:$AO$50004` |
| `Pnl_Carry_Total` | `AP` | `='PNL_Attribution'!$AP$5:$AP$50004` |
| `Pnl_PnL_FX` | `AV` | `='PNL_Attribution'!$AV$5:$AV$50004` |
| `Pnl_Futures_Gov_Model_PnL` | `AW` | `='PNL_Attribution'!$AW$5:$AW$50004` |
| `Pnl_Swap_Curve_Model_PnL` | `AX` | `='PNL_Attribution'!$AX$5:$AX$50004` |
| `Pnl_Hedge_Curve_Model_PnL` | `AY` | `='PNL_Attribution'!$AY$5:$AY$50004` |
| `Pnl_Hedge_Model_Residual_PnL` | `AZ` | `='PNL_Attribution'!$AZ$5:$AZ$50004` |
| `Pnl_Actual_Futures_PnL` | `BJ` | `='PNL_Attribution'!$BJ$5:$BJ$50004` |
| `Pnl_Actual_PlainSwap_PnL` | `BM` | `='PNL_Attribution'!$BM$5:$BM$50004` |
| `Pnl_Bond_DV01_Current` | `L` | `='PNL_Attribution'!$L$5:$L$50004` |
| `Pnl_Hedge_DV01_Gap` | `R` | `='PNL_Attribution'!$R$5:$R$50004` |
| `Pnl_PlainSwap_DV01` | `O` | `='PNL_Attribution'!$O$5:$O$50004` |
| `Pnl_SyntheticSwap_DV01` | `S` | `='PNL_Attribution'!$S$5:$S$50004` |
| `Pnl_FuturesRTJ_DV01` | `P` | `='PNL_Attribution'!$P$5:$P$50004` |
| `Pnl_FuturesRT_DV01` | `Q` | `='PNL_Attribution'!$Q$5:$Q$50004` |
| `Pnl_Actual_Hedge_DV01` | `N` | `='PNL_Attribution'!$N$5:$N$50004` |
| `Pnl_Residual_DV01` | `BE` | `='PNL_Attribution'!$BE$5:$BE$50004` |
| `Pnl_Target_Hedge_DV01` | `T` | `='PNL_Attribution'!$T$5:$T$50004` |
| `Pnl_Hedge_Efficiency` | `BG` | `='PNL_Attribution'!$BG$5:$BG$50004` |
| `Pnl_Hedge_Ratio` | `BF` | `='PNL_Attribution'!$BF$5:$BF$50004` |
| `Pnl_Delta_Y_bp` | `V` | `='PNL_Attribution'!$V$5:$V$50004` |
| `Pnl_Attribution_Status` | `BH` | `='PNL_Attribution'!$BH$5:$BH$50004` |
| `Pnl_Spread_Framework_Auto` | `BY` | `='PNL_Attribution'!$BY$5:$BY$50004` |
| `Pnl_Spread_Framework_Reason` | `BZ` | `='PNL_Attribution'!$BZ$5:$BZ$50004` |
| `Pnl_Duration_Identity_Check` | `CA` | `='PNL_Attribution'!$CA$5:$CA$50004` |
| `Pnl_Row_Valid` | `CB` | `='PNL_Attribution'!$CB$5:$CB$50004` |
| `Pnl_Row_Exclusion_Reason` | `CC` | `='PNL_Attribution'!$CC$5:$CC$50004` |
| `Pnl_FX_Exposure_EUR` | `CD` | `='PNL_Attribution'!$CD$5:$CD$50004` |
| `Pnl_Risk_Timing_Bias` | `CF` | `='PNL_Attribution'!$CF$5:$CF$50004` |
| `Pnl_Coupon_Paid_EUR` | `CG` | `='PNL_Attribution'!$CG$5:$CG$50004` |

## What it reads off `Futures` and `Swaps` directly

Some reconciliation tiles measure what `PNL_Attribution` **cannot** see —
hedge PnL on rows whose `LinkedISIN` reaches no bond. Those read the hedge
sheets by letter. All eight verified correct against your headers:

| Constant | Col | Header |
|---|---|---|
| `DASH_FUT_KEY_COL` | `A` | ContractCode |
| `DASH_FUT_LINK_COL` | `J` | LinkedISIN |
| `DASH_FUT_NOTIONAL_COL` | `T` | NotionalValue_EUR |
| `DASH_FUT_PNL_COL` | `U` | FuturesPnL_EUR |
| `DASH_FUT_CLASS_COL` | `AO` | Hedge_Class |
| `DASH_SWAP_KEY_COL` | `A` | DealID |
| `DASH_SWAP_LINK_COL` | `L` | LinkedISIN |
| `DASH_SWAP_DV01_COL` | `O` | **DV01 BBG** |
| `DASH_SWAP_PNL_COL` | `AL` | PnL |
| `DASH_SWAP_SOURCE_COL` | `AQ` | Swap_ID_Source |
| `DASH_SWAP_FAMILY_COL` | `BL` | FloatIndex_Family |

> `DASH_SWAP_DV01_COL` is `DV01 BBG`, **not** `Swap_DV01_EUR_Theoretical`. The
> attribution sums the Bloomberg DV01 into `PlainSwap_DV01`; the internal model
> DV01 is deliberately not used. This block once said "Swaps!BN" and "Swaps!X",
> which were true of an older sheet — `BN` is `FloatCurve_T0` today. A formula
> built from that prose sums a curve level as if it were a risk number.

Row extents are **read from the sheet**, not frozen. They used to be
`$U$5:$U$224` and `$AB$5:$AB$204`, which quietly assumed at most 220 futures and
200 swaps — past those counts the "unlinked hedge PnL" tiles stopped counting
the very rows most likely to be unlinked.

## The build steps

| # | Step | Notes |
|---|---|---|
| 1 | `CalculateFullRebuild` | before anything is read |
| 2 | `DashValidateSourceHeaders` | `PNL_Attribution` exists and has data |
| 3 | `DashClear` | contents, formats, charts; **keeps shapes named `btnPnl*`** |
| 4 | `DashDefinePnlNames` | resolve every required key or raise 9901 |
| 5 | Header block | as-of dates from `Config!B4`/`B5`, last refresh `Config!B24` |
| 6 | KPI tiles | official, explained, residual, coverage |
| 7 | **Factor bridge** | see the note below |
| 8 | Carry summary | coupon, pull-to-par, funding memo |
| 9 | Hedge summary | bond vs hedge DV01, ratio, gap |
| 10 | Efficiency buckets | from `Pnl_Hedge_Efficiency` |
| 11 | Status / framework counts | text-count summaries |
| 12 | Rank stage (hidden, col `BH`+) | tie-broken sort keys so Top-N uses plain `LARGE`/`MATCH` — no CSE, no VBA sorting |
| 13 | Top-N tables | 10 rows each |
| 14 | Hedge-efficiency detail | |
| 15 | Risk-factor / shared-framework / drift views | |
| 16 | Error locator | |
| 17 | FX hedge view | `Hedge_Class = "FX"` |
| 18 | **Quarantine block** | up to 40 rows, with `Row_Exclusion_Reason` |
| 19 | Charts | failures do not stop the build |
| 20 | Formatting | |

**The bridge.** Only `PnL_Duration_Total` enters the bridge sum. The OIS /
Gov-basis / Swap-Gov / Credit lines below it are **"of which" memo lines and are
deliberately not added in.** Summing those legs directly was double counting:
*which* legs make up the duration total depends on each row's spread framework,
so a G-framework row contributed a Swap-Gov leg it does not actually contain. A
tie-out row proves the bridge closes rather than asking you to trust it.

**Hedge efficiency** is read from `PNL_Attribution!Hedge_Efficiency`. It is
**not** recomputed here — this module used to apply its own definition, a
different target and no clamping, so the same bond scored differently depending
on which sheet you looked at.

---

# The full running order

1. **Load bonds** (however you do that today) — `Bonds` must have rows.
2. **Check `FutMapTable` and `SpreadOverrideTable` exist** — Formulas → Name
   Manager. Without `FutMapTable`, `Futures!BBG_Ticker` is `#NAME?` and the
   whole futures chain stays blank no matter what else is right.
3. **Press `Load_Hedges`.** Read the completion message; scan `Link_Source`.
4. **Paste the 14 corrected formulas** from `PNL_Attribution_formulas.txt`.
   Do `CE` (`=Bonds!CL4`) first — the entire duration chain hangs off it.
5. **`Ctrl+Alt+F9`** — full rebuild, not `F9`.
6. **Define the 45 `Pnl_*` names** (once — they persist in the file).
7. **Build the Dashboard.**

## Verifying after a run

| Check | Expect |
|---|---|
| `Futures!AO` | `RATES` on every rates row — was blank on every row ever loaded |
| `Futures!AM` | `RTJ` or `RT` |
| `Futures!L:AE` | still formulas |
| `Swaps!AQ` | `PLAIN` / `SYNTHETIC` |
| `Swaps!O` and `P:X` | still `DV01 BBG` numbers and BQL formulas |
| `Bonds!BQ`, `BR` | SwapLink, FutLink — and `BS` (`BBG_Cand_ISIN`) untouched |
| `PNL_Attribution!CE` | a number, not blank |
| `PNL_Attribution!P`, `Q` | non-zero on a bond with a futures hedge |
| `PNL_Attribution!O`, `S` | non-zero on a bond with a swap |
| `PNL_Attribution!CA` | near zero |
| `BA` vs `BB` | explained ≈ official on a clean row |
