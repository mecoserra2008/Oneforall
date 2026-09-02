# VBA → Python: where everything went

A map from the old workbook to the new package, so a figure someone remembers
from a sheet can be found in the code.

The economics are reviewed in [MODEL.md](MODEL.md). This file is purely
"where did X go".

---

## The shape of the change

| | VBA workbook | Python package |
|---|---|---|
| engine | Excel formulas written by VBA | `pnlx/`, numpy column arrays |
| inputs | OPICS over ADO, Bloomberg add-in | CSV files |
| config | `Config` sheet, cells B4–B50 | `config/pnl_explainer.yaml` |
| output | the workbook itself | Excel + CSV + JSON |
| tests | static lint + headless LibreOffice Basic | `python -m unittest` |
| lines | ~15,800 VBA | ~5,300 Python incl. tests and docs |

### Why the inputs became files

The workbook could only run next to a Bloomberg terminal with an OPICS
connection. That means it could not be tested, could not be re-run for a past
date, and could not be checked by anyone without a licence.

The boundary is now a set of flat files. Whatever produces them — an export of
the existing sheets, a scheduled query, a market-data job — is a separate
concern, and the engine gets a reproducible input it can be tested against.

### Why the output stopped being formulas

The legacy design rule was "VBA decides layout, Excel computes every number",
which was right when Excel *was* the engine. It is not any more. A sheet that
recomputes the model in a second language is a second model that can disagree
with the first.

What is preserved is **auditability**: every dashboard figure traces to a column
on the Bonds sheet, every column is defined on the Definitions sheet (generated
from the same schema the columns are built from, so it cannot drift), and the CSV
and JSON carry the same numbers for anything downstream.

---

## Modules

| VBA module | Python | notes |
|---|---|---|
| `modPNL` (12,105 lines) | `engine.py`, `bondmath.py`, `daycount.py`, `curves.py`, `frameworks.py`, `loaders.py`, `instruments.py` | the bulk of it was formula-string construction, which has no Python equivalent |
| `modImpRepo` | `daycount.py` | day counts, schedules, accrued |
| `modEconFormulas` | `engine.py`, `bondmath.py` | these returned formula *strings*; the maths is now the maths |
| `modDashboard` | `aggregate.py`, `report/excel.py` | layout split from aggregation |

---

## Config sheet → YAML

| cell | VBA constant | YAML |
|---|---|---|
| `B4` | `CFG_T0_DATE` | `run.prior` ⚠️ |
| `B5` | `CFG_T1_DATE` | `run.as_of` ⚠️ |
| `B18` | `CFG_SPREAD_FRAMEWORK` | `model.spread_framework_override` |
| `B27` | `CFG_BASE_CCY` | `run.base_ccy` |
| `B29` | `CFG_FUT_PRICE_FIELD` | *(gone — prices are an input)* |
| `B31` | `CFG_YEARSLEFT_BASIS` | `model.pull_to_par.day_basis` |
| `B34` | `CFG_ACTUAL_PNL_SOURCE` | *(gone — practical PnL is always built from marks + cash)* |
| `B36` | `CFG_CONVEXITY_BUMP_BP` | `model.convexity_bump_bp` |
| `B39` | `CFG_KEEP_FORMULAS` | *(gone — there are no formulas to keep)* |
| `B49`, `B50` | freeze timestamps | *(gone — nothing is frozen; a run is reproducible from its inputs)* |
| — | `HEDGE_RATIO_TOL` etc. | `tolerances.*` |

> ⚠️ **The date convention was inverted, and it is the single easiest way to flip
> the sign of every delta in the book.** The VBA carried `CFG_T0_DATE` (B4) as the
> *prior* date and `CFG_T1_DATE` (B5) as the *current* one, with a warning block
> at the top of `modPNL` about exactly this. The Python uses the report
> convention throughout: **`prior` = T-1 = opening**, **`as_of` = T0 = closing**.
> There is no ambiguity left anywhere in the package.

---

## Sheets → files

| sheet | file | notes |
|---|---|---|
| `Bonds` A:K (OPICS query) | `positions_bonds.csv` | plus the market columns that were `L:CL` |
| `Bonds` L:CL (VBA-owned) | computed | these were derived; they are now computed, not stored |
| `Futures` | `positions_futures.csv` | inputs only; derived columns are computed |
| `Swaps` | `positions_swaps.csv` | as above |
| `OIS_Curves` (wide) | `curves.csv` (long) | see below |
| `SpreadOverride` | `spread_overrides.csv` | |
| `FutMap`, `SwapMap`, `Coverage support`, `CoverageFutures`, `BBG_T0_Staging` | *(gone)* | ticker plumbing for the Bloomberg add-in |
| `PNL_Attribution` | `daily_pnl_explained.csv` + **Bonds** sheet | |
| `Dashboard` | the whole workbook | |
| `Diagnostics` | **Data Quality** sheet + `daily_pnl_metadata.json` | |

### The curve sheet went from wide to long

`OIS_Curves` gave each currency its own column block with a hard-coded tenor
column — EUR in `B`, USD in `S`, GBP in `AJ`. That layout is what forced
`CurveRanges` to exist at all, and adding a fourth currency meant the same edit
in two places (the resolver *and* the pull-to-par loader), one of which was easy
to miss.

`curves.csv` is long format:

```
currency,curve_type,tenor_years,rate_prior,rate_current
```

A new currency is new rows. There is no resolver.

---

## Column names

Every legacy header is accepted as an alias by the loaders, so an export of the
existing sheets loads without being renamed first. The canonical names are
lower-case with the report's date convention:

| legacy | canonical |
|---|---|
| `CleanPx_T0` | `clean_px_current` |
| `CleanPx_T-1` | `clean_px_prior` |
| `DirtyMV_T0_EUR` | *(computed)* `mv_current_base` |
| `DV01_EUR` | `dv01_current` |
| `DV01_Unit` | `dv01_unit_current` |
| `ISpread_T0` | `ispread_current` |
| `Delta_y_bp` | `delta_y_bp` |
| `PnL_Duration_Total` | `pnl_duration_total` |
| `Total_Model_Explained` | `total_explained` |
| `Official_Total_PnL` | `practical_pnl` |
| `Unexplained_Residual_PnL` | `residual_pnl` |
| `Spread_Framework_Auto` | `spread_framework` |
| `Attribution_Status` | `attribution_status` |
| `Hedge_Model_Residual_PnL` | `hedge_basis_pnl` |
| — *(new)* | `theoretical_pnl`, `dv01_timing_bias`, `pnl_fx_cross`, `coupon_cash`, `attribution_dv01` |

---

## Procedures

| VBA | Python |
|---|---|
| `SetupWorkbookFinalLayout` | *(gone — no layout to maintain)* |
| `LoadOPICS_Bonds` / `LoadOPICS_Hedges` | `loaders.load_inputs` |
| `WriteAllModelFormulas` | *(gone — the model is the code)* |
| `RefreshMarketData` | *(gone — market data is an input)* |
| `WritePNLRow` | `AttributionEngine._bond_legs` / `_bond_frame` |
| `BuildDashboard_Step6` | `report.excel.write_workbook` |
| `RebuildPNLOnly` | `python -m pnlx` |
| `BondPullToParCore` | `bondmath.pull_to_par` |
| `BondPullCurveNodes` / `BondPullInterp` | `curves.ZeroCurve` |
| `InterpOIS` / `InterpGov` / `InterpSwap` | `CurveSet.rates_for` |
| `CurveRanges` | *(gone — see above)* |
| `AccruedInterest`, `PrevCouponDate`, `NextCouponDate` | `daycount.*` (vectorised) |
| `BloombergDayCountToDCC` | `daycount.day_count_from_description` |
| `CouponFreqNum` | `daycount.coupon_frequency_from_description` |
| `SwapFloatFamily` | `instruments.classify_float_family` |
| `SwapFloatFamilyStatus` | `instruments.float_family_status` |
| `ImpliedRepoBloomberg` and the basis suite | *(not carried over — see below)* |
| `GetBondSelectQuery` etc. | *(gone — the SQL lives with whatever produces the CSVs)* |

### What was deliberately not carried over

**The implied-repo / basis suite** (`ImpliedRepoBloomberg`, `NetBasis`,
`ForwardCleanPrice`, `TheoreticalFuturesPrice`, `OptimalDeliveryDate`,
`ImpliedRepoSchedule`, `ImpliedRepoLinear`, `ImpliedRepoIterated`). It is a
correct, self-contained closed-form solver, but nothing in the *attribution*
consumed it — `Futures!R` was a display column. Porting it would have meant
carrying ~200 lines that no reported figure depends on. Gross basis, net basis
and implied repo as published by the feed still flow through to the Futures sheet.

If the desk wants the calculated implied repo back, the maths is in
`src/modImpRepo.bas` and it is a self-contained port.

**The 18 dead private procedures** the legacy linter reported (including
`GetBondSelectQuery` and `GetFuturesSelectQuery`, dead once bonds moved to an
Excel query). They had no callers then and have none now.

**The Bloomberg ticker resolver** (`BBG_Cand_*`, the `/isin/` → `Corp` → `Govt`
→ `Mtge` → `M-Mkt` cascade, the T0 staging sheet, the batched BDH pulls, the
refresh-and-wait loop). All of it exists to get data out of the add-in. The
Python takes data as an input.

---

## Old tooling

`tools/vba_lint.py`, `tools/check_layout.py`, `tools/vbaparse.py`,
`tools/vba_run.py`, `tools/vba_builtins.py` and `tests/*.bas` are kept, and
still work against `src/`. They are the legacy build.

They have no Python equivalent because they have no Python job: the layout
checker enforced sheet geometry that no longer exists, the linter reproduced
compile errors a Python interpreter reports itself, and the LibreOffice Basic
runner existed to execute pure procedures that are now ordinary functions with
ordinary tests.

The pinning test `tests/test_formula_equivalence.bas` — which held the
`modEconFormulas` builders character-for-character against the strings
`WritePNLRow` used to build inline — has no successor by design. Its purpose was
to prove a refactor changed nothing. This migration **does** change things, and
each change is listed in [MODEL.md](MODEL.md#what-changed-and-what-it-costs) with
what it is worth.

---

## Running the old and the new side by side

The Python is additive: `src/` is untouched and the workbook still imports and
runs. To compare a day:

1. Run the workbook as usual and export `Bonds`, `Futures`, `Swaps` and
   `OIS_Curves` to CSV.
2. Rename to the canonical filenames (or point `paths.*` at them — the loaders
   accept the legacy headers as aliases).
3. Set `run.prior` to the workbook's `Config!B4` and `run.as_of` to `Config!B5`.
   **Not the other way round** — see the warning above.
4. To reproduce the legacy numbers rather than the corrected ones:

   ```yaml
   model:
     attribution_risk_date: current   # legacy risk-timing convention
     coupon_carry: smooth             # legacy carry
     fx_cross_term: false             # legacy FX treatment
   ```

   The remaining differences are the ones in
   [MODEL.md](MODEL.md#what-changed-and-what-it-costs) that have no switch,
   because they were defects rather than conventions: the swap model DV01, the
   gross-basis price, and the month-end schedule drift.
