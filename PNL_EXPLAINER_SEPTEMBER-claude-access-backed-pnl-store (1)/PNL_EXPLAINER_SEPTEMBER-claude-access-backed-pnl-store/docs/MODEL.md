# The model: every formula, what it means, and why it is that way

This is the review of the economics that came out of the migration. It states
each formula, what it measures, and — where the Python differs from the VBA —
what changed and what it is worth.

Nothing here is stylistic. Every item in [What changed](#what-changed-and-what-it-costs)
moves a number on the report.

---

## 1. The two totals

Everything the report does is the gap between two numbers.

**Practical PnL** — what the book actually made, from marks and cash:

```
practical = Δ(dirty market value, base ccy)
          + coupon cash received in the period
          + actual hedge PnL (futures variation margin + swap NPV move)
```

**Theoretical PnL** — what the risk model says it should have made:

```
theoretical = duration        (framework chain; ties to -DV01 × Δy)
            + convexity       (second order in the same yield move)
            + carry           (coupon accrual + pull to par)
            + FX              (translation of the opening position + cross term)
            + hedge (model)   (hedge DV01 against the curve each hedge neutralises)
```

Then:

```
explained = theoretical + hedge basis          hedge basis = actual − model hedge
residual  = practical − explained
```

### Why the hedge basis is inside `explained`

`practical` contains the **actual** hedge PnL. Adding the basis back makes the
hedge leg of `explained` telescope to that same actual figure:

```
model hedge + (actual hedge − model hedge) = actual hedge
```

so the bridge closes by construction and `residual` isolates **bond-leg model
error**. That is what "unexplained" should mean.

Leave the basis out — which the legacy sheet did before it was corrected — and
the futures/CTD basis and the swap spread mismatch land in the residual *every
single day*, however good the model is. The basis is still reported on its own
line, because a desk running a cross-instrument hedge wants to watch it.

A useful consequence, and a check worth running: because the hedge legs
telescope, the **bond cash-leg residual equals the bridge residual exactly**.
`tests/test_engine.py` asserts it.

---

## 2. The decomposition the whole model rests on

```
y = r + g + q + i
```

| term | definition | economic meaning |
|---|---|---|
| `y` | bond yield to maturity | what the bond yields |
| `r` | OIS zero rate at the bond's tenor | the risk-free rate |
| `g` | `Gov − OIS` | sovereign / collateral basis |
| `q` | `Swap − Gov` | swap / government basis |
| `i` | `y − Swap` | the bond's own I-spread |

Every term is a **difference of two quoted rates**, so the chain telescopes
exactly. That is the entire reason several frameworks can split the same yield
move differently without any of them changing the total.

---

## 3. Spread frameworks

A framework decides **how the yield move is split**. It never changes the total.
With `D` the attribution DV01 and every delta in basis points:

| code | chain | ties |
|---|---|---|
| `G` | `−D·Δr + −D·Δg + −D·Δ(G-spread)` | **exactly** |
| `I` | `−D·Δr + −D·Δg + −D·Δq + −D·Δi` | **exactly** |
| `ASW` | `−D·Δr + −D·Δg + −D·Δq + −D·Δ(ASW)` | approximately |
| `Z` | `−D·Δr + −D·Δg + −D·Δq + −D·Δ(Z)` | approximately |
| `OAS` | `−D·Δr + −D·Δg + −D·Δq + −D·Δ(OAS)` | approximately |
| `OIS` / `SOFR` | `−D·Δy` (no decomposition) | trivially |
| `MIXED` | DV01-weighted blend of the `G` and `I` chains | approximately |
| `REVIEW` | attribution suppressed | — |

`G` and `I` tie **exactly** because G-spread and I-spread are *defined* as
differences from the same yield. On the sample book they tie to `1e-10` — see
`test_exact_frameworks_tie_to_machine_precision`.

`ASW`, `Z` and `OAS` are quoted on their **own** conventions — asset-swap spread
against a floating leg, Z-spread against the whole zero curve, OAS after
stripping optionality — so their chains tie approximately. `duration_identity_check`
is where that difference shows up.

**Read the identity check as data quality, not model quality.** A non-zero value
on a `G` or `I` row means an input is stale or T0 and T-1 came from different
snapshots. A small standing value on `ASW`/`Z`/`OAS` is expected.

### How a bond's framework is chosen

Precedence, highest first:

1. **per-bond override** — the overrides file, keyed on ISIN
2. **global override** — `model.spread_framework_override`, or `--framework`
3. **automatic** — from the mix of hedges actually attached

The automatic rule follows what the hedge actually neutralises:

- **futures-dominated → `G`.** Bond futures hedge the deliverable *government*
  curve through the CTD, so what is left unhedged is the bond's spread over
  governments.
- **swap-dominated → `I`.** A swap hedges the *swap* curve, so what is left
  unhedged is the bond's spread over swaps.
- **unhedged →** preference order `I`, `G`, `ASW`, `Z`, `OAS`, by availability.

Dominance is measured on **absolute** DV01, so it does not depend on the sign
convention of either hedge leg.

**Synthetic swaps do not vote.** A synthetic is the hedge the desk *could* have
put on, not the one whose risk is in the book. Letting it choose the framework
would measure the bond against a curve it is not exposed to.

A framework whose spread is missing is never chosen *automatically* — picking it
would produce a blank duration total and no attribution at all. An **explicit
override is honoured even when its spread is missing**: an override is a human
decision, and silently overruling it would hide the fact that its data is absent.
The row then reports a missing-framework status, which is the honest outcome.

An unrecognised override code resolves to `REVIEW`, so a typo suppresses
attribution *loudly* rather than quietly changing the answer.

---

## 4. Risk

```
DV01 per unit nominal = |duration| × |dirty price| / 100 × |FX| × 0.0001
position DV01         = notional × DV01 per unit          (signed by notional)
```

Duration preference: OAS duration where quoted (a spread move is what it
measures sensitivity to), then modified duration, then the model's own
reprice-derived duration.

Struck against the **dirty** price, because that is the price the position is
marked at — a duration struck against the clean price understates a high-coupon
bond's risk by the accrued fraction of its value.

### Convexity

```
convexity = (P₊ + P₋ − 2P₀) / (P₀ · Δy²)
PnL       = 0.5 × MV_prior × convexity × (Δy / 10000)²
```

The bump size is configurable and it is not cosmetic: at a large bump these are
the *average* duration and convexity across the bump — which is what a desk
wants when the day's move was large — and at a vanishing bump they converge on
the analytic derivatives. The two differ by around 0.1% on a ten-year bond.
`bump` at 100bp reproduces the legacy workbook.

---

## 5. Carry

```
carry = coupon + pull to par
```

**Coupon** (`coupon_carry: exact`, the default):

```
carry_coupon = notional/100 × (accrued(T0) − accrued(T-1) + coupons paid) × FX_prior
```

**Pull to par** — the part of the clean price change that is pure passage of
time, with the curve and the spread both held at their prior levels.

The naive version — reprice on the *same* curve at a shorter maturity — books
the entire roll-down of an upward-sloping curve as PnL the desk never earned.
The arbitrage-free answer discounts each remaining cash flow at the **forward**
rate the prior curve implies between the T0 horizon `h` and the flow at `t`:

```
(1 + f(h,t))^(t−h) = (1 + y(t))^t / (1 + y(h))^h
```

so

```
DirtyPx_fwd = (1 + y(h))^h × Σ_{t > h} C_t / (1 + y(t))^t
```

the classic "spot price less the PV of interim coupons, compounded to the
horizon". Coupons paid on or before the horizon **drop out on purpose**: that
cash is reported by the coupon leg, and counting it here as well is the single
easiest way to double the carry.

Result, per 100 nominal, with accrued stripped at both ends so coupon accrual
cannot leak in:

```
(DirtyPx_fwd − AI(T0)) − (DirtyPx_spot − AI(T-1))
```

Both legs are **model** prices, so neither ties to a quoted price. Their
*difference* is what is used, and the calibration error is common to both and
cancels to first order. **Neither leg is a price on its own.**

The time axis is ACT/365F, *not* the bond's own day count: the curve nodes are
decimal years against annually compounded zero rates, so the horizon and every
cash flow must be on the same axis for the forward identity to hold. The bond's
day count is still used where it belongs — accrued interest.

`tests/test_bondmath.py` checks the forward identity on flat, steep and inverted
curves, and checks that a steep curve gives *less* pull to par than a flat one —
which is exactly what a spot reprice at a shorter maturity would get backwards.

### Funding is outside the bridge

```
funding_carry_memo = −MV_prior × funding rate × year fraction
```

`practical` is a mark-to-market figure with **no financing leg**, so adding
funding to the explained side would open a residual of exactly the funding cost.
It is reported as an economic-carry memo. `model.include_funding_in_bridge`
turns it on — but only do that if the practical feed becomes funded too, in the
same change.

---

## 6. FX

```
Δ(MV_base) = MV_local_prior × ΔFX          ← pnl_fx
           + Δ(MV_local) × FX_prior        ← the rate / spread / carry legs
           + Δ(MV_local) × ΔFX             ← pnl_fx_cross
```

An **exact identity**, so nothing is left over. Every model leg is struck at the
**opening** fix, the FX leg is the translation of the opening position, and the
second-order term is booked explicitly rather than falling into the residual.
`test_the_fx_decomposition_is_an_exact_identity` asserts it.

---

## 7. Hedges

Both are matched to bonds **by ISIN only**.

### Futures

```
DV01     = contracts × FUT_PX_VAL_BP × FUT_VAL_PT × FX
model PnL = −DV01 × Δ(government curve at the CTD's tenor)
actual PnL = contracts × FUT_VAL_PT × Δprice × FX
```

`FUT_PX_VAL_BP` arrives in **price points** per basis point (about 0.0605 for a
Bund), so the cash figure needs the further multiply by `FUT_VAL_PT` (EUR 1000
per full point) — roughly EUR 60 per contract per bp.

> **Desk check.** `futures_dv01 / contracts` should read **60–90** for a Bund. A
> figure near 60,000 means the feed is already in cash and
> `futures.dv01_includes_point_value` should be `true`. The symptom of getting
> it wrong is every futures-hedged bond reading "Over-hedged".

Futures hedge the **government** curve through the CTD, so the move they respond
to is the government curve — not the bond's yield. The tenor comes from the
CTD if it is in the book, else the linked bond; if neither, there is **no model
PnL**, because a tenor guessed from the delivery month would produce a plausible
number from an assumption nobody made.

### Swaps

```
DV01      = the swap system's published BPV (fallback: annuity model)
model PnL = −DV01 × Δ(the curve its floating leg projects off)
actual PnL = NPV(T0) − NPV(T-1)
```

Which curve depends on the floating index:

| family | curve | why |
|---|---|---|
| ESTR / SOFR / SONIA | OIS | overnight-index swaps |
| EURIBOR | swap | IBOR-projection swaps |
| UNKNOWN | *none* | model leg suppressed, row flagged |

Splitting the DV01 by family and applying each against its own curve move is
what stops an ESTR swap's PnL being measured against EURIBOR — which would book
the OIS/IBOR basis as model error every day.

`UNKNOWN` is a real answer. Defaulting an unclassifiable index to the OIS curve
would produce a plausible hedge PnL from an assumption nobody made.

**Synthetic swaps are targets, not positions.** They carry risk in the hedge
comparison and PnL **nowhere** — asserted by
`test_synthetic_swaps_contribute_no_pnl`.

### Hedge quality

```
actual hedge DV01 = futures + plain swaps
target hedge DV01 = synthetic where one exists, else −bond DV01 (i.e. be flat)
hedge ratio       = −actual / bond DV01
hedge efficiency  = 1 − |actual − target| / |target|
```

Efficiency is **deliberately not clamped to [0, 1]**. Clamping reported a
wrong-way hedge and a triple-sized hedge both as 0%, hiding exactly the positions
worth looking at. A value of −1 means the hedge is as wrong as it could be at
that size, and that is worth seeing.

The synthetic-first target rule lives in **one place**. The legacy sheet applied
one rule and the dashboard another, so the same bond scored differently
depending on which you read.

---

## 8. What the report cannot see

Two consequences of matching hedges by ISIN alone. Neither can be fixed by
arithmetic — they are data questions — so they are **measured and reported**
rather than hidden.

1. **A hedge with a blank LinkedISIN sits on no bond row**, so its PnL is absent
   from every bond-level total. Reported as `unlinked_*` on Data Quality and as
   a bridge memo.
2. **If one ISIN appears on two rows** — the same bond in two portfolios — *both*
   claim its full hedge PnL and DV01. Flagged per row as `shared_isin`.

A total that quietly covers 80% of the book is the failure mode this guards
against, so coverage sits next to every headline figure.

---

## 9. The snail methodology

A snail trail is a **path**, not a bar. The shape is the information.

**Intraday** — walk one day leg by leg. Each step adds a leg and plots what is
still outstanding. A leg that barely moves the line is not explaining anything;
a long run means that leg is doing the work. The walk closes on the practical
total by construction.

**Historical** — one point per run, cumulative, at
`(cumulative practical, cumulative explained)`. A perfect model traces the
45° line.

| shape | reading |
|---|---|
| along the diagonal | the model is tracking |
| steady drift to one side | a systematic bias — a convention, a stale input, a missing leg |
| looping back and forth | noise, or marks arriving on different days |
| a sudden right angle | something changed on that date |

A second pair of axes plots cumulative residual against cumulative risk taken:
is the unexplained PnL growing *with* the risk being run, or independently of it?

The verdict is computed, not left to the eye. The test is the mean daily
residual against the **standard error of the mean** (`std / √n`), i.e. a
t-statistic — not against the standard deviation. Comparing mean to standard
deviation is not sample-size aware: a genuinely random 25-day series sits about
0.2 standard deviations from zero by chance, so a fixed threshold on that ratio
calls honest noise a drift, and gets *more* wrong the shorter the history.
Dividing by `√n` makes the reading mean the same thing after 10 days as after 200.

---

## What changed, and what it costs

Each of these moves a number.

### 1. Risk is struck at the start of the period

The workbook built DV01 from the **T0** dirty price, so the risk used to explain
a move was the risk *after* it. Attribution convention is to strike risk at the
start of the period being explained.

The workbook's own comments quantified the cost: on a self-consistent test book
a 10bp day left **+288 EUR** of residual against **−66,170 EUR** of duration
PnL, of which **+278 (96.5%)** was this convention alone. Scaled up, a 50bp day
carries roughly a **3% standing residual that is pure convention, not model
error**.

Attribution DV01 is now struck at T-1. `dv01_timing_bias` reports what the old
convention would have added, so the change is auditable rather than asserted, and
`model.attribution_risk_date: current` reproduces the old behaviour exactly.

**Hedging metrics still use current risk** — that is the risk to hedge tomorrow,
and the hedge ratio rightly measures it.

### 2. FX is consistent end to end

The workbook translated rate legs at `FX_T0` and the FX leg at `FX_T-1`, which
opens a standing residual on every non-base position. Now every model leg is at
the opening fix and the cross term is explicit — see §6.

### 3. Coupon payments are handled

The workbook accrued coupon smoothly and had **no term for a coupon actually
paid** inside the period. On an ex-coupon date the dirty price drops by the
coupon and nothing on the explained side put the cash back, so that bond showed
a one-day residual **the size of its coupon**. Carry is now
`accrued(T0) − accrued(T-1) + coupons paid`, and the cash appears on the
practical side too. `test_exact_carry_survives_an_ex_coupon_date` covers it.

### 4. The model swap DV01 was 100× too small

`SwapDV01Fml` built the annuity as `(1 − DF) / rate` with the rate in
**percentage points**, so the model DV01 was out by two orders of magnitude. It
went unnoticed because the sheet used the Bloomberg BPV and never read the model
figure. It is fixed rather than reproduced: a fallback that is only ever wrong is
worse than no fallback.

(The model swap **PV** was fine — `spread_pct / rate_pct` equals
`spread_dec / rate_dec`, so the units cancelled there.)

### 5. Gross basis is a clean-price concept

The workbook computed `CTD_DirtyPx − futures × CF`. Gross basis is
`CTD **clean** − futures × CF`; from the dirty price it is overstated by the
CTD's accrued, which is not basis. Computed from clean whenever the CTD is in the
book, and labelled `unavailable` when it is not, rather than silently using the
wrong one.

### 6. Month-end coupon schedules no longer drift

Stepping a schedule period by period compounds the day clamp: 31 March back six
months is 30 September, and six months on from *that* is 30 **March** — silently
moving a 31st bond onto the 30th and shifting every accrual after it by a day.
Every date is now one month-add from maturity, which cannot drift.
`test_month_end_semiannual_schedule_does_not_drift` pins it.

### 7. One sign driver

The workbook carried `PositionSide` in three places, and its own comments record
the cost: `DirtyMV` and `DV01_EUR` were *already* signed by it, so any formula
that multiplied by it again silently flipped every short and corrupted every
downstream DV01 comparison. Signed notional, and nothing multiplies by a side
flag a second time.

### 8. The residual test needs an amount as well as a share

A percentage test alone flags a 24 EUR break on a 65 EUR position and buries the
real ones under it. `tolerances.residual_eur` is an absolute floor, matching how
`identity_eur` already worked.

---

## Known limitations

Real, deliberate, and **not** bugs to be fixed locally.

1. **No T-1 position snapshot.** Positions are a single pull of *today's* book,
   so `MV_prior` is today's notional at yesterday's price. A position opened
   during the period has its whole MV change attributed to market moves; one
   closed during the period is absent entirely and its realised PnL is missing
   from every total; an intraday size change is mis-attributed in proportion.
   **Fixing this needs a second, dated position pull — not a formula change.**

2. **Hedges match on ISIN only.** See §8.

3. **Spreads are applied flat across tenors** in pull to par. This is the same
   calibration the rest of the model uses, but it is a calibration, not a fact.

4. **The annuity swap DV01 fallback ignores the real payment schedule.** It is a
   sanity check. Prefer the published BPV.

---

## Where to look

| question | file |
|---|---|
| day counts, schedules, accrued | `pnlx/daycount.py` |
| price / yield, duration, convexity, pull to par | `pnlx/bondmath.py` |
| curves and interpolation | `pnlx/curves.py` |
| framework chains and resolution | `pnlx/frameworks.py` |
| the attribution itself | `pnlx/engine.py` |
| the bridge and asset-class comparison | `pnlx/aggregate.py` |
| the snail trail | `pnlx/snail.py` |
| every column's unit and meaning | `pnlx/engine.py` (`BOND_SCHEMA`), or the workbook's **Definitions** sheet |
