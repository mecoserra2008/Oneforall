# PNL Explainer

Daily PnL attribution for the fixed-income book: bonds, bond futures and
interest-rate swaps. VBA in Excel, with OPICS for positions and Bloomberg for
market data.

```
docs/      the documentation set — start here
tools/     static analysis, a headless VBA runner, the formula dumper
tests/     the suite
pnlx/      a Python port of the model (separate; see docs/MIGRATION.md)
```

The five VBA modules live in the repository root as `.txt`, exported from the
VBE. Their filenames change with the export date; nothing in the tooling matches
on a filename.

## The documentation set

| | |
|---|---|
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | how the workbook is put together, what depends on what, and the failure modes that do not announce themselves |
| [BUTTONS.md](docs/BUTTONS.md) | the three buttons, what each touches, and why two old ones are gone |
| [COLUMNS.md](docs/COLUMNS.md) | **the column contract, and step by step how to add, move, rename or remove one** |
| [DASHBOARD.md](docs/DASHBOARD.md) | every block on the Dashboard, how the sheet is laid out, how to change it, and the fault map from its audit |
| [FORMULAS.md](docs/FORMULAS.md) | every formula the workbook writes, generated from the source |
| [DECOMPOSITIONS.md](docs/DECOMPOSITIONS.md) | the three ways to decompose the PnL, and what each one tells you |
| [MODEL.md](docs/MODEL.md) | the economics review from the Python port |
| [MIGRATION.md](docs/MIGRATION.md) | VBA → Python: where everything went |
| [RUNBOOK.md](docs/RUNBOOK.md) | **start here to set it up or run it: ten steps from a bare machine to a stored run you can load back** |
| [PIPELINE_SPEC.md](docs/PIPELINE_SPEC.md) | the field-level specification — every source, field, condition, unit and formula, end to end |
| [SETUP_ACCESS_STORE.md](docs/SETUP_ACCESS_STORE.md) | *superseded by RUNBOOK.md*; kept for its bitness diagnosis |
| [ACCESS_STORE.md](docs/ACCESS_STORE.md) | **the Access store, built and wired in: how runs are separated, how a duplicate position is made impossible, and how to load a past run back** |
| [ACCESS_ARCHITECTURE.md](docs/ACCESS_ARCHITECTURE.md) | the wider plan the store is the first phase of — computed facts, curves, Bloomberg points, memory |
| [PIPELINE.md](docs/PIPELINE.md) | **which of the 330 columns are fetched and which are computed, what depends on what, and how the two Excel sheets get served** |

## The daily run

```
1  Button1_LoadBonds                 Bonds, OIS_Curves
2  Button2_LoadHedgesAndAttribute    Futures, Swaps, PNL_Attribution
3  Button3_BuildDashboard            Dashboard
```

`InstallButtons` puts all three on the Dashboard in run order.

There is no "write formulas" button and no "refresh market data" button. A
formula is only ever missing because a row appeared without one, so the writers
live in the button that creates the rows they fill; and nothing is frozen, so
there is no captured snapshot to refresh. [BUTTONS.md](docs/BUTTONS.md) has the
full reasoning.

## Getting the code into Excel

Import, don't paste. In the VBE: **File → Import File…** and pick each of:

| File | Module | Contents |
|---|---|---|
| `modPNL_*.txt` | `modPNL` | buttons, sheet layout, formula writers, PnL model |
| `Dashboard.txt` | `modDashboard` | the Dashboard |
| `Economic_formula_library.txt` | `modEconFormulas` | named builders for the economically meaningful formula strings |
| `ImpRepo.txt` | `modImpRepo` | day counts, accrued interest, coupon schedules, implied repo / basis |
| `modAccess.txt` | `modAccess` | the Access run store: history, position keys, restore |

The VBE takes the module name from the file, so rename each file to the module
name on import, or set `Attribute VB_Name` at the top. The four modules depend on
each other, so importing only some of them will not compile.

Then run `Setup_Workbook_Layout` once, and `InstallButtons`.

## Tests

```
bash tests/run_all.sh
```

There is no Excel here, so the suite works two ways.

**Static analysis.** `tools/vba_lint.py` reproduces the compile errors that
actually stop this workbook: `Option Explicit` violations, undefined procedures,
ambiguous names across modules, unbalanced blocks, continuation limits,
identifiers shadowing VBA keywords, and unreferenced private procedures.
`tools/check_layout.py` enforces the sheet geometry the column letter constants
only imply, and the `PNL_Attribution` ↔ Dashboard column contract.
`tools/check_line_endings.py` enforces CRLF — a `.bas` with bare LFs imports as
joined lines and fails to compile somewhere unrelated to the edit that caused it.

Both linters are self-checked: `tests/test_lint_selftest.py` feeds them one
deliberately broken module per diagnostic, plus a clean one that must stay
silent. A linter that quietly stops detecting things is worse than none.

**Real execution.** `tools/vba_run.py` lifts pure procedures out of the modules
and runs them through headless LibreOffice Basic. That is why the model code is
split into a sheet-facing wrapper and a numeric core: the core touches no
worksheet, so it can be executed and asserted against.

- `tests/test_pull_to_par.bas` — 20 assertions on the pull-to-par maths,
  including the exact forward-return identity on flat, steep and inverted curves.
- `tests/test_formula_equivalence.bas` — pins the `modEconFormulas` builders
  character-for-character against the strings `WritePNLRow` used to build inline.
  The actual side is built from the real `PCOL_`/`BCOL_` constants, so these
  cases also fail when a column moves — which is what makes moving one safe.

Procedures that touch `Range`/`Worksheet` cannot be executed this way; they are
covered by the static checks only.

The tools find the modules by a procedure only each one defines, so they work on
a fresh clone with no staging step, and survive the exports being renamed.

## The column contract, in one paragraph

`modDashboard` used to find `PNL_Attribution`'s columns by searching row 4 for
the header text, which made every string in that row part of the interface
between the two modules — while looking exactly like a row of labels. Retitling
`L4` to `Bond BPVs`, the right name for the desk to read, killed the entire
Dashboard build on error 9901 before it drew a cell. Now every column is declared
once in `PnlLayout` as a **letter**, a **key** and a **label**: the header writer
writes labels, `PublishPnlColumnNames` publishes keys as workbook names against
letters, and the Dashboard uses only the names. Row 4 can say anything, in any
language, and columns can move. See [COLUMNS.md](docs/COLUMNS.md).

## Pull to par

The part of a bond's clean price change that is pure passage of time, with the
curve and the credit spread held at their prior levels. Each remaining cash flow
is discounted at the forward rate the prior curve implies between the T0 horizon
and that cash flow:

```
(1 + f(h,t))^(t-h) = (1 + y(t))^t / (1 + y(h))^h
```

Repricing on the spot curve at a shorter maturity instead would book the
roll-down of an upward-sloping curve as PnL the desk never earned. Coupons paid
inside the horizon drop out — they are already reported by `Carry_Coupon`.

`BondPullToParCore` is the numeric core; `BondPullToParPrice` is the worksheet
UDF that validates a row and loads the curve for it.

## The OPICS bond query

`Bonds!A:K` is the Excel query's output; VBA owns `L:CM`.
`BONDS_QUERY_LAST_COL` states the boundary, `tools/check_layout.py` enforces it,
and `Button1_LoadBonds` refuses to load if the live query table is the wrong
width rather than misaligning every row.

## Known-dead code

`tools/vba_lint.py` reports 18 unreferenced private procedures in `modPNL`, all
pre-dating this work — among them `GetBondSelectQuery` and
`GetFuturesSelectQuery`, which are dead because bonds are now loaded through an
Excel query rather than an ADO recordset. They are left in place: removing them
is a separate decision, not a side effect.
