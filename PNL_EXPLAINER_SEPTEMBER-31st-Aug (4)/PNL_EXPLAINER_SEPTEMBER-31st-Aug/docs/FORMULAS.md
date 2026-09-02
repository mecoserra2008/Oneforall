# The formula record

Every formula the workbook writes into `PNL_Attribution`, exactly as it is
built, for one bond.

**This file is generated.** Regenerate it after any change to `WritePNLRow` or
`PnlLayout`:

```
python3 tools/dump_formulas.py        # the table below
```

`tools/dump_formulas.py` reads the VBA and evaluates the string concatenation
the way VBA would, so this record cannot drift from the code. Hand-transcribed
formula documentation is right on the day it is written and wrong from the next
commit onwards, without anybody noticing — which is worse than having none,
because people trust it.

The economics behind these formulas — why each leg is shaped the way it is — are
in [DECOMPOSITIONS.md](DECOMPOSITIONS.md). How to change a column is in
[COLUMNS.md](COLUMNS.md).

## How to read it

Row numbers are arbitrary and only the substitution pattern matters:

- **`PNL_Attribution` row 5** — so `L5` means "this row's `Bond_DV01_Current`"
- **`Bonds` row 6** — so `'Bonds'!AF6` means "this bond's `DV01_EUR`"
- hedge lookup ranges are shown ending at row `999`; at run time they are sized
  from the hedge sheets themselves, because the number of futures and swaps is
  whatever the coverage books loaded today

The three columns are the three halves of the contract (see
[COLUMNS.md](COLUMNS.md)): where it sits, the **key** the Dashboard asks for, and
the **label** the desk reads in row 4.

Blank formulas are `""`, never `0`, throughout — a missing input must never
masquerade as a zero PnL.

## Sheet prefixes

| Prefix | Sheet |
|---|---|
| `'Bonds'!` | one row per bond position |
| `'Futures'!` | one row per futures hedge |
| `'Swaps'!` | one row per swap hedge |
| `'Config'!` | `B4` = T-1 date, `B5` = T0 date, `B18` = framework fallback |

## Cross-sheet lookups

Two patterns carry most of the model:

```excel
SUMIFS('Futures'!$AC$5:$AC$999, 'Futures'!$J$5:$J$999, A5, ...)
```
sum a futures column over the rows whose `LinkedISIN` (`Futures!J`) is this
bond's ISIN — this is how a hedge reaches a bond row at all.

```excel
BondPullToParPrice('Config'!B4, 'Config'!B5, 'Bonds'!F6, ...)
```
a VBA UDF. `InterpOIS`, `InterpGov`, `InterpSwap` and `BondPullToParPrice` read
`OIS_Curves` **through the object model**, so Excel has no dependency edge from a
curve cell to these formulas — which is why every button ends in
`Application.CalculateFullRebuild`.

---

| Col | Key | Sheet label | Formula |
|---|---|---|---|
| A | `ISIN` | ISIN | `=UPPER(TRIM('Bonds'!A6))` |
| B | `Name` | Name | `='Bonds'!B6` |
| C | `CCY` | CCY | `='Bonds'!C6` |
| D | `Portfolio` | Portfolio | `='Bonds'!J6` |
| E | `AcctgCat` | AcctgCat | `='Bonds'!I6` |
| F | `Notional` | Notional | `='Bonds'!G6` |
| G | `ModDur` | ModDur | `='Bonds'!W6` |
| H | `Convexity` | Convexity | `='Bonds'!X6` |
| I | `SpreadDuration` | SpreadDuration | `='Bonds'!AU6` |
| J | `Days` | Days | `=INT('Config'!B5)-INT('Config'!B4)` |
| K | `YearFrac` | YearFrac | `=IFERROR(J5/365,"")` |
| L | `Bond_DV01_Current` | Bond BPVs | `='Bonds'!AF6` |
| M | `Bond_DV01_Credit_Spread` | Bond DV01 Credit Spread | _(not written by WritePNLRow)_ |
| N | `Actual_Hedge_DV01` | Hedge (BPVs) | `=IF(AND(ISNUMBER(P5+Q5),ISNUMBER(O5)),(P5+Q5)+O5,"")` |
| O | `PlainSwap_DV01` | PlainSwap (BPVs) | `=IFERROR(SUMIFS('Swaps'!$BN$5:$BN$999,'Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"PLAIN"),0)` |
| P | `FuturesRTJ_DV01` | Futures RTJ (BPVs) | `=IFERROR(SUMIFS('Futures'!$AC$5:$AC$999,'Futures'!$J$5:$J$999,A5,'Futures'!$AN$5:$AN$999,"RTJ",'Futures'!$AP$5:$AP$999,"RATES"),0)` |
| Q | `FuturesRT_DV01` | Futures RT (BPVs) | `=IFERROR(SUMIFS('Futures'!$AC$5:$AC$999,'Futures'!$J$5:$J$999,A5,'Futures'!$AN$5:$AN$999,"RT",'Futures'!$AP$5:$AP$999,"RATES"),0)` |
| R | `Hedge_DV01_Gap` | Hedge_DV01_Gap | `=IF(AND(ISNUMBER(N5),ISNUMBER(T5)),N5-T5,"")` |
| S | `SyntheticSwap_DV01` | SyntheticSwap_DV01 | `=IFERROR(SUMIFS('Swaps'!$BN$5:$BN$999,'Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"SYNTHETIC"),0)` |
| T | `Target_Hedge_DV01` | Target_Hedge_DV01 | `=IF(AND(ISNUMBER(S5),S5<>0),S5,IF(AND(ISNUMBER(L5),L5<>0),-L5,""))` |
| U | `Delta_Dirty_MV_EUR` | Variação Bond (EUR) | `=IF(AND(ISNUMBER('Bonds'!AG6),ISNUMBER('Bonds'!AH6)),'Bonds'!AG6-'Bonds'!AH6,"")` |
| V | `Delta_Y_bp` | Variação yield bond (bps) | `='Bonds'!BM6` |
| W | `Delta_r_bp` | Variação yield rf (bps) | `=IF(AND(ISNUMBER('Bonds'!M6),ISNUMBER('Bonds'!N6)),('Bonds'!M6-'Bonds'!N6)*100,"")` |
| X | `Delta_Gov_bp` | Variação yield Gov (bps) | `=IF(AND(ISNUMBER('Bonds'!BB6),ISNUMBER('Bonds'!BC6)),('Bonds'!BB6-'Bonds'!BC6)*100,"")` |
| Y | `Delta_g_bp` | Variação Govi basis bp | `='Bonds'!BH6` |
| Z | `Delta_Swap_bp` | Variação Swap bp | `=IF(AND(ISNUMBER('Bonds'!BD6),ISNUMBER('Bonds'!BE6)),('Bonds'!BD6-'Bonds'!BE6)*100,"")` |
| AA | `Delta_q_bp` | Variação Gov/Swap basis bp | `='Bonds'!BK6` |
| AB | `Delta_i_bp` | Variação I spread bp | `='Bonds'!BL6` |
| AC | `Delta_Z_bp` | Variação Z spread bp | `=IF(AND(ISNUMBER('Bonds'!Y6),ISNUMBER('Bonds'!AD6)),'Bonds'!Y6-'Bonds'!AD6,"")` |
| AD | `Delta_GSpread_bp` | Variação G spread bp | `=IF(AND(ISNUMBER('Bonds'!AL6),ISNUMBER('Bonds'!AM6)),'Bonds'!AL6-'Bonds'!AM6,"")` |
| AE | `Delta_ASW_bp` | Variação ASW bp | `=IF(AND(ISNUMBER('Bonds'!Z6),ISNUMBER('Bonds'!AE6)),'Bonds'!Z6-'Bonds'!AE6,"")` |
| AF | `Delta_OAS_bp` | Variação OAS bp | `='Bonds'!AS6` |
| AG | `PnL_Duration_Total` | PnL_Duration_Total | `=SWITCH(BZ5,"G",IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AR5)),AH5+AI5+AR5,""),"I",IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AJ5),ISNUMBER(AK5)),AH5+AI5+AJ5+AK5,""),"ASW",IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AJ5),ISNUMBER(AS5)),AH5+AI5+AJ5+AS5,""),"Z",IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AJ5),ISNUMBER(AQ5)),AH5+AI5+AJ5+AQ5,""),"OAS",IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AJ5),ISNUMBER(AT5)),AH5+AI5+AJ5+AT5,""),"OIS",IF(AND(ISNUMBER(CF5),ISNUMBER(V5)),-CF5*V5,""),"SOFR",IF(AND(ISNUMBER(CF5),ISNUMBER(V5)),-CF5*V5,""),"MIXED",LET(_f,ABS(P5+Q5),_s,ABS(O5),_t,_f+_s,IF(_t=0,"",IFERROR(_f/_t*(IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AR5)),AH5+AI5+AR5,""))+_s/_t*(IF(AND(ISNUMBER(AH5),ISNUMBER(AI5),ISNUMBER(AJ5),ISNUMBER(AK5)),AH5+AI5+AJ5+AK5,"")),""))),"REVIEW","","")` |
| AH | `PnL_OIS` | PnL_OIS | `=IF(AND(ISNUMBER(CF5),ISNUMBER(W5)),-CF5*W5,"")` |
| AI | `PnL_GovBasis` | PnL_GovBasis | `=IF(AND(ISNUMBER(CF5),ISNUMBER(Y5)),-CF5*Y5,"")` |
| AJ | `PnL_SwapGovBasis` | PnL_SwapGovBasis | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AA5)),-CF5*AA5,"")` |
| AK | `PnL_Credit_Ispread` | PnL_Credit_Ispread | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AB5)),-CF5*AB5,"")` |
| AL | `PnL_Convexity` | PnL_Convexity | `=IF(AND(ISNUMBER('Bonds'!AH6),ISNUMBER(H5),ISNUMBER(V5)),0.5*'Bonds'!AH6*H5*(V5/10000)^2,"")` |
| AM | `Carry_Coupon` | Carry_Coupon | `=IF(AND(ISNUMBER(F5),ISNUMBER('Bonds'!D6),ISNUMBER(K5),ISNUMBER('Bonds'!R6)),F5*('Bonds'!D6/100)*K5*'Bonds'!R6,"")` |
| AN | `Carry_RollToPar` | Carry_RollToPar | `=IFERROR(F5*'Bonds'!R6*BondPullToParPrice('Config'!B4,'Config'!B5,'Bonds'!F6,'Bonds'!D6,'Bonds'!CE6,'Bonds'!CK6,C5,BZ5,BondSpreadTMinus1(BZ5,'Bonds'!AM6,'Bonds'!AK6,'Bonds'!AE6,'Bonds'!AD6,'Bonds'!AR6))/100,"")` |
| AO | `Funding_Carry_Memo` | Funding_Carry_Memo | `=IF(AND(ISNUMBER('Bonds'!AH6),ISNUMBER(K5)),LET(_r0,'Bonds'!BO6,_r1,'Bonds'!BP6,_fr,IF(AND(ISNUMBER(_r0),ISNUMBER(_r1)),AVERAGE(_r0,_r1),IF(ISNUMBER(_r1),_r1,IF(ISNUMBER(_r0),_r0,""))),IF(ISNUMBER(_fr),-'Bonds'!AH6*IF(ABS(_fr)>1,_fr/100,_fr)*K5,"")),"")` |
| AP | `Carry_Total` | Carry_Total | `=IF(AND(ISNUMBER(AM5),ISNUMBER(AN5)),AM5+AN5,"")` |
| AQ | `PnL_ZSpread` | PnL_ZSpread | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AC5)),-CF5*AC5,"")` |
| AR | `PnL_GSpread` | PnL_GSpread | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AD5)),-CF5*AD5,"")` |
| AS | `PnL_ASW` | PnL_ASW | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AE5)),-CF5*AE5,"")` |
| AT | `PnL_OAS` | PnL_OAS | `=IF(AND(ISNUMBER(CF5),ISNUMBER(AF5)),-CF5*AF5,"")` |
| AU | `SpreadPnL_Used` | SpreadPnL_Used | `=SWITCH(BZ5,"G",AR5,"I",AK5,"ASW",AS5,"Z",AQ5,"OAS",AT5,"OIS",IF(AND(ISNUMBER(CF5),ISNUMBER(V5),ISNUMBER(W5)),-CF5*(V5-W5),""),"SOFR",IF(AND(ISNUMBER(CF5),ISNUMBER(V5),ISNUMBER(W5)),-CF5*(V5-W5),""),"MIXED",LET(_f,ABS(P5+Q5),_s,ABS(O5),_t,_f+_s,IF(_t=0,"",IFERROR(_f/_t*(AR5)+_s/_t*(AK5),""))),"REVIEW","","")` |
| AV | `PnL_FX` | PnL_FX | _(not written by WritePNLRow)_ |
| AW | `Futures_Gov_Model_PnL` | Futures_Gov_Model_PnL | `=IF(BJ5=0,0,IF(AND(ISNUMBER(P5+Q5),ISNUMBER(X5)),-(P5+Q5)*X5,""))` |
| AX | `Swap_Curve_Model_PnL` | Swap_Curve_Model_PnL | `=IF(BM5=0,0,IF(COUNTIFS('Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"PLAIN",'Swaps'!$K$5:$K$999,"UNKNOWN")>0,"",-SUMPRODUCT(('Swaps'!$L$5:$L$999=A5)*('Swaps'!$AG$5:$AG$999="PLAIN")*('Swaps'!$BN$5:$BN$999)*IF(('Swaps'!$K$5:$K$999="ESTR")+('Swaps'!$K$5:$K$999="SOFR"),W5,IF('Swaps'!$K$5:$K$999="EURIBOR",Z5,0)))))` |
| AY | `Hedge_Curve_Model_PnL` | Hedge_Curve_Model_PnL | `=IF(AND(ISNUMBER(AW5),ISNUMBER(AX5)),AW5+AX5,"")` |
| AZ | `Hedge_Model_Residual_PnL` | Hedge_Model_Residual_PnL | `=IF(AND(ISNUMBER(BT5),ISNUMBER(BW5)),BT5+BW5,"")` |
| BA | `Total_Model_Explained` | Total_Model_Explained | `=IF(AND(ISNUMBER(AG5),ISNUMBER(AL5),ISNUMBER(AP5),ISNUMBER(AV5),ISNUMBER(AY5),ISNUMBER(AZ5)),AG5+AL5+AP5+AV5+AY5+AZ5,"")` |
| BB | `Official_Total_PnL` | Official_Total_PnL | `=IF(AND(ISNUMBER(U5),ISNUMBER(BX5)),U5+N(CH5)+BX5,"")` |
| BC | `Unexplained_Residual_PnL` | Unexplained_Residual_PnL | `=IF(AND(ISNUMBER(BB5),ISNUMBER(BA5)),BB5-BA5,"")` |
| BD | `Unexplained_Residual_Pct` | Unexplained_Residual_Pct | `=IF(AND(ISNUMBER(BC5),ISNUMBER(BB5),BB5<>0),BC5/ABS(BB5),"")` |
| BE | `Residual_DV01` | Residual_DV01 | `=IF(AND(ISNUMBER(L5),ISNUMBER(N5)),L5+N5,"")` |
| BF | `Hedge_Ratio` | Hedge_Ratio | `=IFERROR(-N5/L5,"")` |
| BG | `Hedge_Efficiency` | Hedge_Efficiency | `=IF(AND(ISNUMBER(N5),ISNUMBER(T5),T5<>0),1-ABS(N5-T5)/ABS(T5),"")` |
| BH | `Attribution_Status` | Attribution_Status | `=IF('Bonds'!BQ6<>"OK",'Bonds'!BQ6,IF(BZ5="REVIEW","Framework review required",IF(AND(BM5>0,NOT(ISNUMBER(BN5))),"Missing actual swap PnL",IF(AND(BJ5>0,NOT(ISNUMBER(BK5))),"Missing actual futures PnL",IF(AND(ISNUMBER(BF5),BF5<0),"Wrong-direction hedge",IF(AND(ISNUMBER(BF5),BF5>1+0.02),"Over-hedged",IF(NOT(ISNUMBER(AG5)),"Missing selected framework data",IF(AND(ISNUMBER(CB5),ABS(CB5)>MAX(1,0.01*ABS(AG5))),"Duration chain does not tie",IF(NOT(ISNUMBER(BB5)),"Missing official PnL",IF(AND(ISNUMBER(BD5),ABS(BD5)>0.1),"High residual","OK"))))))))))` |
| BI | `Synthetic_Alternative_DV01` | Synthetic_Alternative_DV01 | `=S5` |
| BJ | `Futures_Match_Count` | Futures_Match_Count | `=COUNTIFS('Futures'!$J$5:$J$999,A5,'Futures'!$AP$5:$AP$999,"RATES")` |
| BK | `Actual_Futures_PnL` | Actual_Futures_PnL | `=IF(BJ5=0,0,IF(COUNTIFS('Futures'!$J$5:$J$999,A5,'Futures'!$U$5:$U$999,">=-1E+307",'Futures'!$AP$5:$AP$999,"RATES")=0,"",SUMIFS('Futures'!$U$5:$U$999,'Futures'!$J$5:$J$999,A5,'Futures'!$AP$5:$AP$999,"RATES")))` |
| BL | `Swap_All_Match_Count` | Swap_All_Match_Count | `=COUNTIF('Swaps'!$L$5:$L$999,A5)` |
| BM | `PlainSwap_Match_Count` | PlainSwap_Match_Count | `=COUNTIFS('Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"PLAIN")` |
| BN | `Actual_PlainSwap_PnL` | Actual_PlainSwap_PnL | `=IF(BM5=0,0,IF(COUNTIFS('Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"PLAIN",'Swaps'!$AB$5:$AB$999,">=-1E+307")=0,"",SUMIFS('Swaps'!$AB$5:$AB$999,'Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"PLAIN")))` |
| BO | `SyntheticSwap_Match_Count` | SyntheticSwap_Match_Count | `=COUNTIFS('Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"SYNTHETIC")` |
| BP | `Actual_SyntheticSwap_PnL` | Actual_SyntheticSwap_PnL | `=SUMIFS('Swaps'!$AB$5:$AB$999,'Swaps'!$L$5:$L$999,A5,'Swaps'!$AG$5:$AG$999,"SYNTHETIC")` |
| BQ | `AllSwapRows_PnL_Diagnostic` | AllSwapRows_PnL_Diagnostic | `=SUMIFS('Swaps'!$AB$5:$AB$999,'Swaps'!$L$5:$L$999,A5)` |
| BR | `Actual_Futures_PnL_Check` | Actual_Futures_PnL_Check | `=BK5` |
| BS | `Futures_Gov_Model_PnL_Check` | Futures_Gov_Model_PnL_Check | `=AW5` |
| BT | `Futures_Model_Residual_PnL` | Futures_Model_Residual_PnL | `=IF(BJ5=0,0,IF(AND(ISNUMBER(BK5),ISNUMBER(AW5)),BK5-AW5,""))` |
| BU | `Actual_PlainSwap_PnL_Check` | Actual_PlainSwap_PnL_Check | `=BN5` |
| BV | `Swap_Curve_Model_PnL_Check` | Swap_Curve_Model_PnL_Check | `=AX5` |
| BW | `Swap_Model_Residual_PnL` | Swap_Model_Residual_PnL | `=IF(BM5=0,0,IF(AND(ISNUMBER(BN5),ISNUMBER(AX5)),BN5-AX5,""))` |
| BX | `Actual_Hedge_PnL` | Actual_Hedge_PnL | `=IF(AND(ISNUMBER(BK5),ISNUMBER(BN5)),BK5+BN5,"")` |
| BY | `Hedge_Model_Residual_PnL_Check` | Hedge_Model_Residual_PnL_Check | `=AZ5` |
| BZ | `Spread_Framework_Auto` | Spread_Framework_Auto | `=IF(A5="","",LET(_ovr,IFERROR(UPPER(TRIM(VLOOKUP(A5,SpreadOverrideTable,2,FALSE))),""),_gbl,IFERROR(UPPER(TRIM('Config'!B18)),""),_fut,ABS(P5+Q5),_swp,ABS(O5),_tot,_fut+_swp,_hasI,ISNUMBER(AK5),_hasG,ISNUMBER(AR5),_hasASW,ISNUMBER(AS5),_hasZ,ISNUMBER(AQ5),_hasOAS,ISNUMBER(AT5),_both,AND(_fut>0,_swp>0,_hasG,_hasI,MIN(_fut,_swp)>=0.2*_tot),_auto,IF(_both,"MIXED",_pick,IF(_ovr<>"",_ovr,IF(_auto<>"",_auto,IF(_gbl<>"",_gbl,"REVIEW"))),IF(ISNA(MATCH(_pick,{"G","I","ASW","Z","OAS","OIS","SOFR","MIXED","REVIEW"},0)),"REVIEW",_pick)))` |
| CA | `Spread_Framework_Reason` | Spread_Framework_Reason | `=IF(BZ5="","",LET(_ovr,IFERROR(UPPER(TRIM(VLOOKUP(A5,SpreadOverrideTable,2,FALSE))),""),_gbl,IFERROR(UPPER(TRIM('Config'!B18)),""),_fut,ABS(P5+Q5),_swp,ABS(O5),_anyLeg,OR(ISNUMBER(AK5),ISNUMBER(AR5),ISNUMBER(AS5),ISNUMBER(AQ5),ISNUMBER(AT5)),_src,IF(_ovr<>"","per-bond override (SpreadOverride sheet)",IF(_anyLeg,IF(OR(_fut>0,_swp>0),"automatic from hedge DV01 mix","automatic from available spread legs (no hedge attached)"),IF(_gbl<>"","Config!B18 fallback - automatic rule had no spread leg to pick","no framework resolvable - no spread leg and no fallback"))),_why,SWITCH(BZ5,"G","measured vs governments - futures/govie hedge or G-spread strongest available","I","measured vs swaps - swap hedge or default credit framework","ASW","asset-swap spread","Z","Z-spread fallback","OAS","option-adjusted spread","OIS","no decomposition - whole yield move taken as -DV01 x Delta_y","SOFR","no decomposition - whole yield move taken as -DV01 x Delta_y","MIXED","hedge split across futures and swaps - DV01-weighted blend of the govie and swap chains","REVIEW","attribution suppressed - unrecognised override code, or no usable spread leg on this row","framework fallback"),_why&" [" &_src& "]"))` |
| CB | `Duration_Identity_Check` | Duration_Identity_Check | `=IF(AND(ISNUMBER(AG5),ISNUMBER(CF5),ISNUMBER(V5)),AG5-(-CF5*V5),"")` |
| CC | `Row_Valid` | Row_Valid | `=IF(A5="","",IF(CD5="",1,0))` |
| CD | `Row_Exclusion_Reason` | Row_Exclusion_Reason | `=IF(A5="","",IF('Bonds'!BQ6<>"OK","Bond data: "&'Bonds'!BQ6,IF(BZ5="REVIEW","Spread framework unresolved",IF(AND(BJ5>0,NOT(ISNUMBER(BK5))),"Futures matched but actual futures PnL missing",IF(AND(BM5>0,NOT(ISNUMBER(BN5))),"Swaps matched but actual swap PnL missing",IF(AND(ISNUMBER(CB5),ABS(CB5)>MAX(1,0.01*ABS(AG5))),"Duration chain does not tie","")))))))` |
| CE | `FX_Exposure_EUR` | FX_Exposure_EUR | `=IF(A5="","",IF(C5="EUR",0,IFERROR('Bonds'!AH6,0)))` |
| CF | `Bond_DV01_Opening` | Bond_DV01_Opening | `='Bonds'!CM6` |
| CG | `Risk_Timing_Bias` | Risk_Timing_Bias | `=IF(AND(ISNUMBER(CF5),ISNUMBER(L5),ISNUMBER(V5)),-(CF5-L5)*V5,"")` |
| CH | `Coupon_Paid_EUR` | Coupon_Paid_EUR | `=IF(A5="","",IFERROR(F5*SumCouponsBetween('Config'!B4,'Config'!B5+1,'Bonds'!F6,'Bonds'!D6/100,'Bonds'!CE6)/100*'Bonds'!R6,0))` |

---

## Columns not written here

`Bond_DV01_Credit_Spread` (`M`) is declared in `PnlLayout` and published, but
nothing writes it. It resolves cleanly and sums to a confident zero. It is listed
in `KNOWN_EMPTY_COLUMNS` in `tools/check_layout.py` so the check does not fail on
it — what it should hold that `Bond_DV01_Current` does not is a desk question
rather than a coding one, since `Bond_DV01_Current` is already struck on
`SpreadDuration_Used` where that exists.

## The formula builders

Formulas whose economic meaning has a name are built by a named builder in
`modEconFormulas` rather than typed inline, so exactly one place defines each
quantity:

| Builder | Quantity |
|---|---|
| `MarketValueChangeFml` | change in dirty market value |
| `HedgeDv01GapFml` | actual hedge BPV less target |
| `HedgeEfficiencyFml` | `1 − abs(actual − target) / abs(target)` |
| `ResidualPnLFml` | official less explained |
| `ResidualPercentFml` | residual as a share of official |
| `RollToParFml` | pull to par, wrapping the `BondPullToParPrice` UDF |
| `DiffFormula` | a guarded difference, optionally scaled to basis points |
| `ChainSumFml` | a framework chain — guarded sum of its legs |
| `DV01BlendFml` | the `MIXED` blend, weighted by the hedge's own DV01 split |

`tests/test_formula_equivalence.bas` pins each builder's output against the
literal the inline code used to produce, character for character, so moving a
formula into the library provably changed nothing. Because the pinned side is
built from the real `PCOL_`/`BCOL_` constants, those cases **also fail when a
column moves** — which is what makes moving one safe.
