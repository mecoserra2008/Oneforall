# The pipeline specification

Every step from source system to Dashboard, stated as **source, field,
condition, unit**.

```
Fetch OPICS ─┐
             ├─► T0 / T-1 tables ─► decompose size into impact ─► PNL_Attribution ─► Dashboard
Fetch Hedges ┤
Fetch BBG   ─┘
```

The point of writing it this way is that the **wiring** then lives in
[`config/pipeline.yaml`](../config/pipeline.yaml) and the **maths** stays in
`Economic_formula_library.txt`. Change where a step runs — Excel today, Access
tomorrow — and neither the formulas nor the workflow move.

Every claim below is quoted from source with a `file:line`. Where a document in
`docs/` disagrees, the code was checked and the document is wrong.

- [1 · Fetch OPICS — bonds](#1--fetch-opics--bonds)
- [2 · Fetch Hedges — two books](#2--fetch-hedges--two-books)
- [3 · Fetch Bloomberg](#3--fetch-bloomberg)
- [4 · The T0 and T-1 tables](#4--the-t0-and-t-1-tables)
- [5 · Decompose size into impact](#5--decompose-size-into-impact)
- [6 · Feed PNL_Attribution](#6--feed-pnl_attribution)
- [7 · Feed the Dashboard](#7--feed-the-dashboard)
- [8 · What is configurable](#8--what-is-configurable)

---

## 1 · Fetch OPICS — bonds

### Connection

- Server = `LOKI`
- Database = `OPICSMAIN`, context `OpicsMain`
- Branch = `01`
- Portfolio = `PORT`
- Accounting category = `A`
- Eligible products = `OBR`, `SECUR`
- Security alternative identifier type = `ISIN`

### Tables

| Table | Role | Join |
|---|---|---|
| `SPOS` | the position | — |
| `SECM` | security master | `SPOS.SECID = SECM.SECID` — **INNER** |
| `ASID` | alternative ids | `SPOS.SECID = ASID.SECID` **and** `ASID.SECIDTYPE = 'ISIN'` — **LEFT** |
| `TPOS` | book values | `BR` + `COST` + `PORT` + `SECID` + `INVTYPE` — **LEFT** |

`SECM` is INNER: a position whose `SECID` is not in the security master is
dropped. `ASID` and `TPOS` are LEFT: a position with no ISIN, or no book-value
row, still comes through.

### Filters

```
SPOS.QTY      <> 0
SPOS.PORT      = 'PORT'
TRIM(SPOS.BR)  = '01'
TRIM(SPOS.INVTYPE) IN ('A')
TRIM(SECM.PRODUCT) IN ('OBR','SECUR')
```

### Fields

| Output | Source | Unit |
|---|---|---|
| `ISIN` | `ASID.SECALTID` | text |
| `Name` | `SECM.DESCR` | text |
| `Currency` | `SPOS.CCY` | text |
| `Coupon_Dec` | `SECM.COUPRATE_8` | **see the unit note below** |
| `CouponFreq` | `SECM.INTPAYCYCLE` | text |
| `MaturityDate` | `SECM.MDATE` | date |
| `Notional` | `SPOS.QTY` | nominal |
| `AvgCost` | computed — the ladder below | price per 100 |
| `AcctgCat` | `TRIM(SPOS.INVTYPE)` | text |
| `Product` | `SECM.PRODUCT` | text |
| `BookVal_EUR` | `TPOS.TDBOOKVALUE`, `0` when NULL | EUR |

`AvgCost`, in order — the division is never performed when the absolute notional
is zero:

```
if |SPOS.QTY| > 0 and |TPOS.TDBOOKVALUE|  > 0 :  |TDBOOKVALUE|  / |QTY| * 100
if |SPOS.QTY| > 0 and |TPOS.ADJTDBOOKVAL| > 0 :  |ADJTDBOOKVAL| / |QTY| * 100
otherwise                                     :  SPOS.PURCHAVGCOST
```

Ordered by `CCY`, `TRIM(INVTYPE)`, `COST`, `MDATE`.

> **Unit note — `Coupon_Dec` is a PERCENTAGE.** Settled: many coupon rates in the
> book are above 1, which a decimal rate cannot be. The alias name is misleading
> and the code is right — `Carry_Coupon` uses `Bonds!D6/100` (`docs/FORMULAS.md`,
> column `AM`) and the convexity bump block does the same
> (`modPNL_28th_Aug.txt:11878`). No change needed.

> **Naming defect — `Bonds!J` is PRODUCT, surfaced everywhere as "Portfolio".**
> Settled: it carries `SECM.PRODUCT`, i.e. `OBR` / `SECUR`. Its constant is
> `BCOL_PORTFOLIO` (`:361`), it copies to `PNL_Attribution!D` named `Portfolio`
> (`:8081`), it is stored as `Pos_Bond.Portfolio`, and the Dashboard heads a
> column with it.
>
> **No total is wrong.** It is referenced once in `Dashboard.txt` (`:2283`), as a
> display cell in the Top-N tables, and is never an aggregation key. It is also
> the middle component of `AccBondPositionKey`, which is harmless — the query
> filters `sp.PORT = 'PORT'`, so the real portfolio is that constant on every
> row, and Product discriminates strictly more than a portfolio would have.
>
> The fix is to rename the label, not to re-point it at `Bonds!J`:
> `BCOL_PORTFOLIO` → `BCOL_PRODUCT`, the PnL field and the store column to
> `Product`, and the Dashboard heading to match. Only the column *count* is
> validated on load, never the names, so nothing catches this.
>
> **The portfolio dimension is real, but it is in the filters, not in a column.**
> See "Scope" immediately below.

### Mechanism

The query is a Power Query connection **inside the workbook**; its text is not in
this repository. The load is one call:

```vb
lo.QueryTable.Refresh BackgroundQuery:=False     ' modPNL_28th_Aug.txt:4195
```

against `wsBnd.ListObjects(1)`, guarded first by four **hard-coded literals**
(`:4173-4193`): a ListObject exists, header row `= 3`, first column `= 1`,
exactly `11` ListColumns. `BONDS_QUERY_LAST_COL` is *not* what is enforced, so
widening the query means editing the literals too.

Row count is then measured by `End(xlUp)` on column A (`LastBondDataRow :10871`),
and every downstream writer is sized to `[4 .. lastBondRow]`.

### Scope — the dimension that lives in the WHERE clause

The extraction does not fetch the book. It fetches **one slice** of it, and the
slice is defined by four filters rather than by any column:

| Dimension | Value | Where it is visible afterwards |
|---|---|---|
| Branch | `01` | nowhere |
| Portfolio | `PORT` | nowhere |
| Accounting category | `A` | `Bonds!I` `AcctgCat` — the only one that survives |
| Product | `OBR`, `SECUR` | `Bonds!J`, mislabelled `Portfolio` |

`Bonds!I` carries `TRIM(SPOS.INVTYPE)`, and because the query filters
`INVTYPE IN ('A')` it is the constant `A` on every row. So the accounting
category *is* a real dimension — the book has others — and what the run holds is
the `A` slice of the `PORT` portfolio on branch `01`.

**Two consequences the architecture has to carry.**

*The scope is a property of the run, not of the data.* Nothing in `Bonds`,
`PNL_Attribution` or the store says which slice a run captured: `Run`
(`access/schema.sql:26`) records dates, counts, user, version and status, and
**no branch, portfolio, category or product**. Two runs over different slices are
indistinguishable except by inspecting their contents. Widening `INVTYPE` to
`('A','T')` next month would produce a run that looks exactly like today's and
means something else.

*The real dimension is the pair.* `Portfolio` in the schema
(`access/schema.sql:71`) is already keyed `(PortfolioCode, AcctgCat)` — the right
shape, and unused. A total is only comparable with another total over the same
pair.

So the run database records the scope explicitly (`Run.Scope*`,
`Run_Index.Scope*`), the config declares it in one place, and any cross-run
comparison joins on it rather than assuming it.

---

## 2 · Fetch Hedges — two books

Both are `M:\P_Pires\TRADING\`, sheet `Resumo`, and both are scanned **in full** —
the blocks are not contiguous, so stopping at the first blank row silently drops
everything below it.

**RTJ carries swaps and futures. RT carries futures only.**

| | **RTJ** — `Hedge_Risco Tx Juro.xlsm` | **RT** — `Hedge_Risco Total.xlsm` |
|---|---|---|
| Rows | 9 – 220 | 8 – 34 |
| Columns copied | `A:AI` (formula of A stashed in `AJ`) | `A:W` (formula of A stashed in `X`) |
| Filter | `L = HtC&S` | `M = HtC&S` |
| `#RC` | `B` | `A` |
| Covered ISIN | `C` | `B` |
| Bond notional | `F` — **store only** | `E` — **store only** |
| Type | `H` | `L` |
| Counterparty | `M` | `N` |
| Contracts | `O` | `P` |
| Coverage info | `D` | — |
| Start / End | `P` / `Q` | `Q` / — |
| Currency | `T` | `R` |
| Desk BPV | — | `S` |
| Swap ids | `I` deal, `J` fixed, `K` float | n/a |
| Synthetic ids | `AC` / `AD` / `AE`, notional `AF`, end `AI` | n/a |

The contract-count columns genuinely differ between the two books: **`O` for RTJ**
(`:13215`), **`P` for RT** (`HEDGE_TOTAL_NOTIONAL_COL`, `:240`). Both are correct.

### The row-type rule

```
if Type == "Plain Vanilla Swap"  -> a SWAP;    read I, J, K
otherwise                        -> a FUTURE;  Type is the contract ticker
                                               (RX1, OE1, UB1, UXY1, TU1, FY1, ...)
```

One RTJ source row can yield **two** swap rows — a `PLAIN` from `I/J/K` and a
`SYNTHETIC` from `AC/AD/AE` — distinguished downstream by `Swap_ID_Source`.

### Where the rows land

Loader-owned columns only; everything else on those sheets is a formula.

- **`Swaps`** — `A` DealID, `B` Ccy, `C` Notional, `G` Start, `H` End,
  `J` Portfolio (the `#RC`), `L` **LinkedISIN**, `M` Cpty, `AC` AF_Status,
  `AD:AF` BBG ids, `AG` **Swap_ID_Source**, `AH:AP` the map block. `I` PayFixed
  is filled later from OPICS.
- **`Futures`** — `A` ContractCode, `C` CCY, `D` Contracts, `I` Portfolio,
  `J` **LinkedISIN**, `AG:AP` the coverage block including `AN` **Hedge_Source**
  (`RTJ`/`RT`) and `AP` **Hedge_Class** (`RATES`/`FX`).
- **`Bonds`** — `BR` SwapLink, `BS` FutLink, as CSV lists. SwapLink carries
  `PLAIN` swaps only: a synthetic is an alternative to the plain swap on the same
  relation, not a second hedge, so listing it would read as double cover.

`LinkedISIN` is the join to everything downstream. A hedge with a blank one
reaches no bond row and is in no total.

---

## 3 · Fetch Bloomberg

**This layer cannot leave Excel.** There is no BLPAPI, no COM, no Bloomberg
object anywhere in the modules — `BDP`, `BDH` and `BQL` are worksheet functions
and nothing else. Everything downstream can move; this cannot.

Two snapshots, both reproducible from `Config`: **T0 via `BDP`**, **T-1 via `BQL`
point-in-time** dated from `Config!B4`. Nothing is frozen, so there is no captured
value to refresh.

### 3.1 Resolve the identity first

Nothing can be priced until the security is named, and naming it is itself a
Bloomberg question. Ten candidates are built from the ISIN into `Bonds!BT:CC`
(`:5777`) — bare ISIN, `/isin/`, then `Corp`, `@BVAL Corp`, `@BGN Corp`, `Govt`,
`@BVAL Govt`, `@BGN Govt`, `Mtge`, `M-Mkt` — and resolved by nesting
`BDP(candidate,"PARSEKYABLE_DES")` down the list until one parses (`:5817`,
`:10138`), falling back to `"UNKNOWN"`.

This is re-asked for every bond on every run, and the answer essentially never
changes — which is why `Instrument` in the schema is worth filling as a cache.

### 3.2 Curves — `OIS_Curves`

Header row **4**, data rows **7–19**, thirteen tenors. Three blocks side by side,
`Q/R` and `AH/AI` being declared spacers (`tools/check_layout.py:53`):

| | EUR | USD | GBP |
|---|---|---|---|
| Years | `B` | `S` | `AJ` |
| OIS ticker / T-1 / T0 | `C` `D` `E` | `T` `U` `V` | `AK` `AL` `AM` |
| Gov ticker / T-1 / T0 | `F` `G` `H` | `W` `X` `Y` | `AN` `AO` `AP` |
| Swap ticker / T-1 / T0 | `I` `J` `K` | `Z` `AA` `AB` | `AQ` `AR` `AS` |
| `g` T-1 / T0 | `L` `M` | `AC` `AD` | `AT` `AU` |
| `q` T-1 / T0 | `N` `O` | `AE` `AF` | `AV` `AW` |
| Status | `P` | `AG` | `AX` |

Derived, per currency and per snapshot:

```
g = Gov  - OIS        ' the govie basis
q = Swap - Gov        ' the swap/gov basis
```

Tenors: `0.0833, 0.25, 0.5, 1, 2, 3, 5, 7, 10, 15, 20, 30, 40` (`:12018`).
Seeded EUR OIS tickers: `ESTRON Index`, then `EESWEA/EC/EF` (1M/3M/6M) and
`EESWE1/2/3/5/7/10/15/20/30 Curncy` (`:12022-12036`).

> **Defect — the EUR OIS curve is seeded one tenor out.** The two arrays are
> paired by index (`:12040-12051`), so `ESTRON` (overnight) lands on the `0.0833`
> (1M) node and `EESWE30` on the `40` node. Every EUR OIS node is mislabelled in
> tenor on a freshly set-up workbook. Either prepend an overnight tenor or drop
> the `40`.

> **Defect — T0 and T-1 are not commensurable for EUR and GBP Gov.** `H` and `AP`
> are `BDP(..., "YLD_YTM_MID")` while their T-1 partners come through BQL as
> `PX_LAST`. `g_T0 = Gov_T0 - OIS_T0` and `g_T-1` are then built from different
> quantities, so `Delta_g_bp` and every leg above it is wrong for those two
> currencies.

`InterpOIS` / `InterpGov` / `InterpSwap` (`:10452-10460`) locate the currency
block, linearly interpolate on `Years`, extrapolate flat outside the node range,
and return `#N/A` when the curve cannot be resolved — which every call site wraps
in `IFERROR(..., "")`, so a missing curve reads as *no spread* rather than as an
error.

### 3.3 Bonds

Field constants at `:840-850`; BQL expressions at `:863-875`.

| Bonds column | T0 — `BDP` | T-1 — `BQL` point-in-time |
|---|---|---|
| CleanPx | `PX_LAST` | `PX_LAST` |
| DirtyPx | `PX_DIRTY_MID` | `price(price_type='dirty')` |
| YTM | `YLD_YTM_MID` | `YIELD(YIELD_TYPE='YTM')` |
| ZSprd | `Z_SPRD_MID` | `SPREAD(SPREAD_TYPE='Z')` |
| ASW | `ASSET_SWAP_SPD_MID` | `SPREAD(SPREAD_TYPE='ASW')` |
| OAS | `OAS_SPREAD_MID` | `SPREAD(SPREAD_TYPE='OAS')` |
| ModDur | `DUR_ADJ_MID` | — |
| OAS duration | `DUR_ADJ_OAS_MID` | — |
| OAS convexity | `CONVEXITY_OAS` | — |
| Pricing source | `PRICING_SOURCE` | — |
| FundRate | `PX_LAST` | `PX_LAST` (BDH last point over `[T0-7, T0]`) |

The T0 block is `WriteBonds_T1_BDP_Efficient` (`:6095`) — the name says T1 and it
writes T0. The T-1 block is `WriteBondT0BQLFormulas` (`:7078`) — the name says T0
and it writes T-1. Both write through `PutBondF` (`:6088`) and the T-1 writer
returns its union range so the refresh can wait on it.

Convexity is not taken from Bloomberg. It is bump-and-repriced (`:11821`): the
bump in bp comes from `Config!B36`, three prices are computed (`Price_Base`,
`Price_Up`, `Price_Down`), and `Convexity` follows from them.

### 3.4 Futures

`BBG_Ticker = VLOOKUP(ContractCode, FutMapTable, 2, FALSE)` (`:6609`) — the
contract root from the coverage label resolved through the hand-maintained
`FutMap` sheet. Then `FUT_VAL_PT` (point value), `FUT_PX_VAL_BP` (**unit DV01,
point value already included**), `LAST_TRADEABLE_DT` falling back to
`FUT_DLV_DT`, `PX_LAST` for both snapshots, and the CTD chain
`CTD_ISIN -> CTD ticker -> CTD dirty price -> implied repo / gross basis`.

That CTD chain is the **one genuine ordering constraint inside the fetch**: the
futures basis cannot be computed until the cheapest-to-deliver bond has been
priced, and which bond that is comes from OPICS.

### 3.5 Swaps

BQL: `SW_MARKET_VAL`, `SW_MARKET_VAL_PRIOR`, `SW_MARKET_VAL_CHG`, `notional`,
`dv01` (`:870-875`), pulled direct and per leg for both snapshots.

**`DV01_BBG` is the figure the attribution sums** — not `Swap_DV01_EUR_Theoretical`,
the in-house annuity model, which is deliberately not used.

---

## 4 · The T0 and T-1 tables

One snapshot pair per run.

| | Dated from | Fetched by |
|---|---|---|
| **T-1** — the opening position | `Config!B4` | BQL point-in-time |
| **T0** — the closing position | `Config!B5` | BDP |

Every `Delta_*` on `PNL_Attribution` is the difference of the two, in basis
points:

```
Delta_Y_bp   = Bonds!Delta_y_bp                     ' bond yield
Delta_r_bp   = (OIS_T0  - OIS_T-1)  x 100           ' risk-free
Delta_Gov_bp = (Gov_T0  - Gov_T-1)  x 100
Delta_Swap_bp= (Swap_T0 - Swap_T-1) x 100
Delta_g_bp   = Bonds!Delta_g_bp                     ' govie basis
Delta_q_bp   = Bonds!Delta_q_bp                     ' swap/gov basis
Delta_i_bp   = Bonds!Delta_i_bp                     ' I-spread
Delta_Z_bp   = ZSprd_T0 - ZSprd_T-1
Delta_ASW_bp = ASW_T0   - ASW_T-1
Delta_OAS_bp = Bonds!DeltaOAS
```

Curve rates arrive as percentages, hence the `x 100` to reach basis points.
Spreads arrive already in basis points and are differenced directly. **Getting
this wrong by a factor of 100 is the single most likely error in a port.**

> After any recalculation, `Application.CalculateFullRebuild` is required, not
> `Calculate`. `InterpOIS`/`InterpGov`/`InterpSwap` read `OIS_Curves` through the
> object model, so Excel has **no dependency edge** from a curve cell to the bonds
> that use it and a plain recalculation leaves every one of them holding the
> previous run's number. `RebuildPNLOnly` (`:7586`) uses `Range.Calculate` and so
> has exactly this bug.

---

## 5 · Decompose size into impact

Every leg is **a size times a move**. The size is a DV01 or a market value; the
move is a delta in basis points or a price change.

### 5.1 Size — rolling the hedges onto the bond

All four by `SUMIFS` keyed on `LinkedISIN`:

```
PlainSwap_DV01     = SUMIFS(Swaps!DV01_BBG,            LinkedISIN=ISIN, Swap_ID_Source="PLAIN")
SyntheticSwap_DV01 = SUMIFS(Swaps!DV01_BBG,            LinkedISIN=ISIN, Swap_ID_Source="SYNTHETIC")
FuturesRTJ_DV01    = SUMIFS(Futures!Futures_DV01_EUR,  LinkedISIN=ISIN, Hedge_Source="RTJ", Hedge_Class="RATES")
FuturesRT_DV01     = SUMIFS(Futures!Futures_DV01_EUR,  LinkedISIN=ISIN, Hedge_Source="RT",  Hedge_Class="RATES")

Actual_Hedge_DV01  = (FuturesRTJ_DV01 + FuturesRT_DV01) + PlainSwap_DV01
Target_Hedge_DV01  = SyntheticSwap_DV01 if non-zero, else -Bond_DV01_Current
```

This is the step that makes the bridge close per bond, and the reason a hedge
with a blank `LinkedISIN` disappears from every total.

### 5.2 Size — a contract count becoming EUR risk

`Economic_formula_library.txt:321` and `:306`:

```
Futures_DV01_EUR = Contracts x HedgeUnitDV01 x FX
FuturesPnL_EUR   = Contracts x FUT_VAL_PT x (FutPx_T0 - FutPx_T-1) x FX
```

`HedgeUnitDV01` is `FUT_PX_VAL_BP`, which **already carries the point value** —
hence `unitDv01IncludesPointValue = True` and no second multiplication by
`FUT_VAL_PT` in the DV01 line. Multiplying by it again is the obvious way to get
this wrong by the contract size.

### 5.3 Impact — each leg

`DV01` is **EUR per basis point** and every `Delta_*` is **in basis points**, so
each product is EUR. Verbatim from `docs/FORMULAS.md`:

```
PnL_OIS            = -Bond_DV01_Opening x Delta_r_bp
PnL_GovBasis       = -Bond_DV01_Opening x Delta_g_bp
PnL_SwapGovBasis   = -Bond_DV01_Opening x Delta_q_bp
PnL_Credit_Ispread = -Bond_DV01_Opening x Delta_i_bp
PnL_ZSpread        = -Bond_DV01_Opening x Delta_Z_bp
PnL_GSpread        = -Bond_DV01_Opening x Delta_GSpread_bp
PnL_ASW            = -Bond_DV01_Opening x Delta_ASW_bp
PnL_OAS            = -Bond_DV01_Opening x Delta_OAS_bp

PnL_Convexity      = 0.5 x DirtyMV_T-1_EUR x Convexity x (Delta_Y_bp / 10000)^2
Carry_Coupon       = Notional x (Coupon / 100) x YearFrac x FX_T-1
Carry_RollToPar    = Notional x FX_T-1 x BondPullToParPrice(...) / 100
Carry_Total        = Carry_Coupon + Carry_RollToPar
PnL_FX             = IFERROR(IF(CCY="EUR", 0,
                       (DirtyMV_T-1_EUR / FX_T-1) x (FX_T0 - FX_T-1)), 0)

Futures_Gov_Model_PnL = -(FuturesRTJ_DV01 + FuturesRT_DV01) x Delta_Gov_bp
Swap_Curve_Model_PnL  = -SUMPRODUCT(matched PLAIN swap DV01 x the curve move
                                    for that swap's FloatCurve_Type)
Hedge_Curve_Model_PnL = Futures_Gov_Model_PnL + Swap_Curve_Model_PnL

Risk_Timing_Bias   = -(Bond_DV01_Opening - Bond_DV01_Current) x Delta_Y_bp
```

The minus sign is the bond convention: yields up, price down.

`Swap_Curve_Model_PnL` returns **blank** — not zero — when any matched swap has
`FloatCurve_Type = "UNKNOWN"`, which drops the bond out of the headline rather
than understating it.

### 5.4 Why opening risk

Every leg uses `Bond_DV01_Opening` (the T-1 risk), never `Bond_DV01_Current`.
The move being attributed happened *to the position as it stood at T-1*, so that
is the risk it acted on. `Risk_Timing_Bias` measures exactly what using current
risk instead would have changed — it is a diagnostic, not a leg.

### 5.5 The framework switch

`Spread_Framework_Auto` resolves to one of
`G / I / ASW / Z / OAS / OIS / SOFR / MIXED / REVIEW`, in this precedence:

1. a per-bond override from `SpreadOverrideTable`;
2. otherwise automatic, from the hedge DV01 mix and which spread legs are
   numeric — `MIXED` when futures and swaps are both present and neither is less
   than 20% of the total, else the dominant side's framework;
3. otherwise the `Config!B18` fallback;
4. otherwise `REVIEW`, which suppresses attribution for the row.

That choice decides **which legs sum** into `PnL_Duration_Total` and **which
single leg** becomes `SpreadPnL_Used`:

| Framework | `PnL_Duration_Total` | `SpreadPnL_Used` |
|---|---|---|
| `G` | OIS + GovBasis + GSpread | `PnL_GSpread` |
| `I` | OIS + GovBasis + SwapGovBasis + Credit_Ispread | `PnL_Credit_Ispread` |
| `ASW` | OIS + GovBasis + SwapGovBasis + ASW | `PnL_ASW` |
| `Z` | OIS + GovBasis + SwapGovBasis + ZSpread | `PnL_ZSpread` |
| `OAS` | OIS + GovBasis + SwapGovBasis + OAS | `PnL_OAS` |
| `OIS` / `SOFR` | `-Bond_DV01_Opening x Delta_Y_bp` (no decomposition) | `-Bond_DV01_Opening x (Delta_Y_bp - Delta_r_bp)` |
| `MIXED` | DV01-weighted blend of the `G` and `I` chains | the same blend |
| `REVIEW` | blank | blank |

The `MIXED` weighting:

```
_f = ABS(FuturesRTJ_DV01 + FuturesRT_DV01)
_s = ABS(PlainSwap_DV01)
_t = _f + _s
SpreadPnL_Used = _f/_t x PnL_GSpread + _s/_t x PnL_Credit_Ispread
```

### 5.6 The identity

```
Total_Model_Explained = PnL_Duration_Total + PnL_Convexity + Carry_Total
                      + PnL_FX + Hedge_Curve_Model_PnL + Hedge_Model_Residual_PnL

Official_Total_PnL    = Delta_Dirty_MV_EUR + Coupon_Paid_EUR + Actual_Hedge_PnL

Unexplained_Residual  = Official_Total_PnL - Total_Model_Explained
```

`Duration_Identity_Check = PnL_Duration_Total - (-Bond_DV01_Opening x Delta_Y_bp)`
proves the decomposed duration chain still adds back to the whole yield move. A
row where it does not tie is quarantined by `Attribution_Status`.

---

## 6 · Feed PNL_Attribution

One row per bond; hedges aggregate onto it by `LinkedISIN`. 86 fields, `A..CH`.

They resolve in **9 waves with no cycles**, derived mechanically by
`tools/field_graph.py` from the generated formula record:

| Wave | n | What settles |
|---|---|---|
| 0 | 26 | identity, every `Delta_*_bp`, `Bond_DV01_Opening`, `PnL_FX` |
| 1 | 23 | the four DV01 roll-ups, single-leg PnLs, match counts |
| 2 | 10 | `Actual_Hedge_DV01`, `Target_Hedge_DV01`, `Carry_Coupon`, **`Spread_Framework_Auto`** |
| 3 | 16 | `PnL_Duration_Total`, `SpreadPnL_Used`, `Carry_RollToPar`, ratios |
| 4 | 4 | `Carry_Total`, **`Official_Total_PnL`**, `Duration_Identity_Check` |
| 5 | 3 | **`Total_Model_Explained`**, `Row_Exclusion_Reason` |
| 6 | 2 | `Unexplained_Residual_PnL`, `Row_Valid` |
| 7 | 1 | `Unexplained_Residual_Pct` |
| 8 | 1 | `Attribution_Status` |

**A field is an output of one wave and an input to the next** — `Bond_DV01_Opening`
is produced in wave 0 and read by 12 later fields; `ISIN` by 19. So the contract
carries `ProducedInWave` and `ConsumedBy` per field (`Meta_Column`), never an
input/output flag.

**Wave gating.** Wave *n* may not start until wave *n−1* is checked: row count
equals the bond count, and every field the wave owns is non-null except where a
documented guard permits null. Failures go to `Run_Issue` and stop the run.

> **Defect — `Bond_DV01_Credit_Spread` (`M`) is never written.** Two references
> in the whole module: the declaration (`:473`) and `AddPnlCol` (`:7697`). It is
> permanently blank, and anything reading `Pnl_Bond_DV01_Credit_Spread` gets
> nothing, silently.

---

## 7 · Feed the Dashboard

`PNL_Attribution` is addressed **only** through published workbook names
`Pnl_<Field>`, never a column letter and never by searching the header row. Move
a column or retitle a header and nothing on the Dashboard changes.

Eleven direct reads off the hedge sheets remain, for the tiles that measure what
`PNL_Attribution` cannot see — hedge PnL on rows whose `LinkedISIN` reaches no
bond: `Futures!A/J/T/U/AO` and `Swaps!A/L/O/AL/AQ/BL`.

**The bridge sums `PnL_Duration_Total` only.** The OIS / Gov-basis / Swap-Gov /
Credit lines beneath it are **memo**. Summing them double-counts, because which
legs make up the duration total depends on each row's framework — a `G`-framework
row would contribute a Swap-Gov leg it does not actually contain.

Hedge efficiency is **read** from `PNL_Attribution`, never recomputed, so the two
sheets cannot disagree about what the same bond scored.

---

## 8 · What is configurable

The split that makes the architecture swappable:

| | Lives in | Changes when |
|---|---|---|
| **Wiring** — paths, sheets, rows, columns, filters, which formula feeds which field, the wave order | [`config/pipeline.yaml`](../config/pipeline.yaml) | a source moves or a column is added |
| **Maths** — 53 named builders | `Economic_formula_library.txt` | the economics change |
| **Execution** — where a wave runs | the engine | you move a wave to Access |

A field entry names a formula from the catalogue and binds its arguments to other
fields. Changing architecture rewires; it never re-derives:

```yaml
- {name: PnL_OIS, wave: 1, formula: first_order,
   args: {dv01: Bond_DV01_Opening, delta_bp: Delta_r_bp}, units: EUR}
```

### Settled

- **`COUPRATE_8` is a percentage.** Coupon rates above 1 rule out a decimal. The
  existing `/100` is correct; nothing to change.
- **`Bonds!J` is Product.** A naming defect only — see §1. Rename the label.
- **The portfolio dimension is real and lives in the filters**, not in a column:
  the run holds the `INVTYPE = 'A'` slice of portfolio `PORT` on branch `01`.
  `Bonds!I` is the constant `A`. The scope must be recorded per run — see
  §1 "Scope".

### Still open

- **The EUR OIS tenor offset** — 13 tickers `ESTRON..EESWE30` are paired by index
  against 13 tenors `1M..40Y`, so every EUR OIS node is one tenor out on a
  freshly set-up workbook (§3.2).
- **EUR and GBP Gov T0 vs T-1 use different Bloomberg fields**, so `Delta_g_bp`
  is built from incommensurable quantities for those currencies (§3.2).
- **`Bond_DV01_Credit_Spread` is never written** (§6).
- **`RebuildPNLOnly` uses `Range.Calculate`**, leaving the curve UDFs a run stale
  (§4).
