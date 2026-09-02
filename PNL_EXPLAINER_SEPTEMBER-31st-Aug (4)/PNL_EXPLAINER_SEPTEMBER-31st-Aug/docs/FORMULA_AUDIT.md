# The formula audit

The desk supplied the live workbook's four header rows and the complete
PNL_Attribution formula map. This is what the comparison found, and what was
done about it.

The short version: **the headers match, the formulas do not.** Every column
constant in the code lands on the right column of the real sheet. The formulas
in the live workbook are a stale generation, and one of them is fatal.

- [The dead attribution](#the-dead-attribution)
- [The one-column shift on Bonds](#the-one-column-shift-on-bonds)
- [The three wrong Swaps letters](#the-three-wrong-swaps-letters)
- [What was already right](#what-was-already-right)
- [What changed in the code](#what-changed-in-the-code)
- [Still open](#still-open)

---

## The dead attribution

The live sheet's `Bond_DV01_Opening` (PNL_Attribution!CE) is:

```
=Bonds!CM4
```

**Bonds ends at CL.** CM is one column past the end of a 90-column sheet — an
empty cell, which Excel reads as `0`.

Every duration and spread leg multiplies by that cell:

```
PnL_OIS            = -CE5*W5    ->  -0*W5  =  0
PnL_GovBasis       = -CE5*Y5    ->  0
PnL_SwapGovBasis   = -CE5*AA5   ->  0
PnL_Credit_Ispread = -CE5*AB5   ->  0
PnL_ZSpread / PnL_GSpread / PnL_ASW / PnL_OAS  ->  0
PnL_Duration_Total                             ->  0
```

So the entire duration and spread attribution reports **exactly zero**, and the
whole bond PnL lands in `Unexplained_Residual_PnL`.

Nothing errors. Every `ISNUMBER` guard passes, because `0` is a number. And the
one check built to catch a broken chain is defeated by the same fact:

```
Duration_Identity_Check = AG5 - (-CE5*V5)  =  0 - 0  =  0
```

The identity ties **because both sides are zero**. A check that passes hardest
exactly when the thing it checks is most broken.

This is the most likely cause of "the Dashboard doesn't work".

The macro does not have this bug. `WritePNLRow` references the original
directly — `dvOpen = 'Bonds'!CL<row>`, from `BCOL_DV01_OPENING_EUR` — with no
intermediate copy column. Rebuilding PNL_Attribution fixes it.

**L013** was added so this class cannot come back: a cross-sheet reference must
name its column through the right family's constant, never a literal. It was
proved against a reintroduced `Bonds!CM` before being left green.

---

## The one-column shift on Bonds

Every letter the live formulas get wrong on Bonds is **exactly one column
later** than the truth, and only from `Bond_Status` onward:

| Live formula reads | That letter really is | Should be |
|---|---|---|
| `Bonds!CM4` opening DV01 | *past the last column* | `CL` DV01_Opening_EUR |
| `Bonds!CE4` coupon frequency | `Price_Base` | `CD` CouponFreq_Num |
| `Bonds!CK4` day-count code | `BondDCC_Name` | `CJ` BondDCC_Code |
| `Bonds!BQ4` bond status | `SwapLink` | `BP` Bond_Status |
| `Bonds!BP4` funding rate T-1 | `Bond_Status` | *does not exist* |

The shape names the cause. The sheet once had a **`FundRate_T-1`** between
`FundRate_T0` and `Bond_Status`. It was dropped; everything after it moved one
column left; the formulas did not. `Funding_Carry_Memo` has been averaging a
**text status string** into a funding rate ever since.

The macro uses `CD`, `CJ`, `CL`, `BP`, and `BO` alone for funding — all
correct. Regenerating fixes all five.

---

## The three wrong Swaps letters

| Live formula reads | That letter really is | Should be |
|---|---|---|
| `Swaps!BN` swap DV01 | `FloatCurve_T0` | `O` DV01 BBG |
| `Swaps!AG` PLAIN/SYNTHETIC | `PV_T-1 da anuidade` | `AQ` Swap_ID_Source |
| `Swaps!AB` actual swap PnL | `DF_T-1` | `AL` PnL |
| `Swaps!L`, `Swaps!K` | LinkedISIN, FloatCurve_Type | correct |

A `SUMIFS` filtered on `Swaps!AG="PLAIN"` matches nothing, returns `0`, and
reads as *this bond has no swap hedge*. Summing `Swaps!BN` adds up curve levels
as if they were risk numbers.

**Every Futures reference was already correct** (`AB`, `J`, `AM`, `AO`, `U`).

### This one was our fault

Three comments in this repository named exactly those wrong letters:

* `Dashboard.txt` — "Swaps!BN DV01_BBG is the risk PNL_Attribution actually
  sums", sitting two lines above `DASH_SWAP_DV01_COL = "O"`
* `modPNL` — "PNL_Attribution never read Swaps!X - it uses the Bloomberg BPV in
  Swaps!BN"
* `modPNL` — "Swaps!AG = SYNTHETIC", two lines above code writing to
  `WCOL_SWAP_ID_SOURCE` (AQ)

Each was true of an older sheet. Anyone writing a formula from the prose rather
than the constant got the old layout. **All three now name the column by its
constant and carry no letter**, because a letter in a comment is a fact with no
check on it.

The same sweep found a geometry assertion that had been false since the Swaps
realignment:

```vba
Debug.Assert colNum(WCOL_FLOATFAMILY_STATUS) = colNum(WCOL_DV01_BBG) + 1
```

`FloatFamily_Status` is column 77; `DV01_BBG` is 15. Nothing joins them. Removed.

---

## What was already right

Worth stating, because it is most of the model:

* **The attribution chains tie.** `G` takes `r + g + G-spread` — three legs, and
  correctly so, because a G-spread is already measured against governments and
  never passes through swaps. `I` takes the full `r + g + q + i`. `ASW`, `Z` and
  `OAS` substitute their leg for `i`. `OIS`/`SOFR` take the yield move whole.
* `Total_Model_Explained` correctly does **not** add `SpreadPnL_Used`
  separately — the spread leg is already inside `PnL_Duration_Total`, so adding
  it would double-count.
* `Carry_Total` correctly excludes `Funding_Carry_Memo`. The memo is a memo.
* `Actual_Hedge_PnL` and `Hedge (BPVs)` both exclude the synthetic swap,
  consistently: the synthetic is the *alternative* to the plain swap, not an
  addition to it.
* **The DV01 sign convention**, confirmed with the desk: hedges carry the
  opposite sign to the bond, so unhedged risk is `Bond_DV01 + Hedge_DV01`.
  `Residual_DV01` already computes that. `Hedge_DV01_Gap` is `Hedge - Target`,
  a different and also useful quantity. The live sheet's `L-N` was neither.
* `Hedge_Efficiency` already builds `1-ABS(N-T)/ABS(T)`, where a matched hedge
  scores 1. The live sheet's version had a `#REF!` in its guard *and* compared
  the wrong pair.

No `#REF!` and no `[PNL_Explainer_v6_after_changes.xlsm]` external link appears
anywhere the macro emits. Both are workbook-side damage, and both go away on a
rebuild.

---

## What changed in the code

**A missing spread leg no longer deletes the bond from the book.**

`ChainSumFml` used to guard every leg with `ISNUMBER` and return `""` if any one
was absent. `PnL_Duration_Total` is on `RowValidRequiredCols`, so a single
missing spread dropped the **whole bond** out of the attribution — good yield
move, good carry, good hedge, and no line in the book.

Now:

* the legs go through `N()`, so the total still computes, understated by exactly
  the missing leg;
* the guard that remains is on the **anchor** — with no opening DV01 every leg is
  blank, `N()` of all of them is a confident zero, and the duration identity
  ties at `0 - 0 = 0`. That is precisely how the dead attribution above passed
  every check it had;
* `SpreadPnL_Used` was removed from `RowValidRequiredCols` — it is the leg the
  framework selected, it is already inside `PnL_Duration_Total`, and
  `Total_Model_Explained` never reads it;
* `ChainIncompleteFml` asks only about the legs the selected framework uses, and
  `Attribution_Status` reports **"Incomplete spread legs"**;
* `Row_Exclusion_Reason`'s "Duration chain does not tie" now fires only when
  every leg is present — otherwise it would report a tolerance breach when the
  real answer is "the OAS did not arrive".

**What this changes in the numbers:** a row with a missing leg now reports an
explained total short by that leg. The shortfall lands in
`Unexplained_Residual_PnL`, and `Attribution_Status` says why. The row stays in
the book and says what is wrong with it, which is the point of an attribution.

Both builders are pinned string-for-string in
`tests/test_formula_equivalence.bas`.

---

## Still open

**`Bond_DV01_Credit_Spread`** (PNL_Attribution!M) is declared, has a header, and
**has no writer** — an empty column in the code, not only on your sheet. Its
natural definition uses the spread duration already in column I rather than
modified duration:

```
=IF(AND(ISNUMBER(I5),ISNUMBER('Bonds'!AH4)),I5*'Bonds'!AH4/10000,"")
```

That is a modelling choice, so it is left empty pending the desk's call.

**`Bonds!J` is headed `Product` and holds the portfolio.** Confirmed with the
desk. There is no column headed `Portfolio` anywhere on Bonds. The header comes
from the OPICS refresh that owns Bonds A:L, which is **outside this module**, so
renaming it is a change to that query and not to any code here. The fixture
records what the sheet says, and the alias in `tools/check_layout.py` bridges
the two, both now carrying the explanation.

**`GetBondSelectQuery` is dead code** and disagrees with the real sheet — it
selects `EntryDate` and `Branch`, which Bonds does not have, in positions the
sheet gives to `AcctgCat` and `Product`. Nothing calls it. It is a plausible
source of a future wrong assumption about what Bonds A:L contains.

**The one test that settles all of it, in Excel:** rebuild PNL_Attribution, then
check `PnL_Duration_Total` is **non-zero** on a bond that moved, and that
`Unexplained_Residual_PnL` is small rather than equal to the whole bond PnL.
That single comparison is the difference between the sheet you have and a
working one.
