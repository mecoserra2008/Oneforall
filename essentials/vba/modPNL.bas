Option Explicit


' =============================================================================
' PNL Explainer BiG - Optimized & Corrected Workflow Module (v2)
'
' This revision applies the full diagnostic remediation:
'   * Bond market-data fallback chain: resolved ticker, then <ISIN> ISIN,
'     then /isin/, then the Corp / Govt / Mtge / M-Mkt venue forms.
'   * I-spread / G-spread / DV01 / SpreadDuration are DERIVED, not pulled from
'     fields that are Not Applicable / Invalid for this book.
'   * All config references inside FormulaR1C1 use CfgR1C1() (no literal sheet
'     name concatenation, which broke DaysToDeliv and futures price pulls).
'   * Futures CTD bond resolved from OPICS CTD_ISIN, not FUT_CTD_BOND.
'   * Futures generic ticker built from a CurveMap/FutMap lookup table, not
'     blind  code & "1 Comdty".
'   * T0 missing is preserved as "" (never silently coerced to 0), so spread
'     deltas are blanked instead of producing fake spread PnL.
'   * Status columns require T0 presence; PNL_Attribution is fully rebuilt to
'     clear stale #REF! formulas.
'
' Buttons:
'   0) SetupWorkbookFinalLayout
'   1) LoadOPICS_Bonds        Bonds!A:K  (Excel query against OPICS)
'   2) LoadOPICS_Hedges       Futures, Swaps
'   3) WriteAllModelFormulas  every formula on every sheet, no data fetched
'   4) RefreshMarketData      T-1 and T0 from Bloomberg, then recalculate
'   5) BuildDashboard_Step6   (modDashboard)
'   *) RebuildPNLOnly         fast PNL-only rebuild
'
' Nothing is frozen.  The T-1 columns are BQL point-in-time queries dated from
' Config!B4, so a rerun reproduces them exactly; there is no snapshot to capture and
' none to protect.  See docs/BUTTON_MAP.md.
'
' Bloomberg failure behaviour: formulas are written by button 3 and never removed, so a
' failed fetch leaves #N/A in place of a value rather than blanking the sheet.
' RefreshMarketData waits per section, so a failure names the section that failed.
'
' Core PnL mapping:  y = r + g + q + i
'   r = OIS rate;  g = Gov - OIS;  q = Swap - Gov;  i = I-spread
'
' FORMULA ORGANIZATION (this reorganized copy):
'   Formula BODIES reference columns by their letter-constant, not raw R1C1 numbers:
'     RC(BCOL_GOV_T0)  emits "RC55"   (see "FORMULA REFERENCE HELPERS" section)
'   so a column shift updates both the range target AND every formula that reads it.
'   The letter constants are the ONLY place a column is named.  A write that
'   targets a literal letter is a bug even when it currently points at the right
'   column - tools/check_layout.py fails the build on one (L008), because such a
'   write and the range that waits for it drift apart the moment a column moves.
'     * R1C1 bodies:      raw RCn  -> RC(const)          (Bonds/Futures/Curves/Swaps-BQL)
'     * repeated diffs:   IF(AND(ISNUMBER..))  -> DiffFormula(a, b [,scaled])
'     * A1 bodies:        WritePNLRow & WriteSwapsCalculatedFormulas cross-sheet/self refs
'     * empty-row guards: =IF(RC1="","",body) -> WrapIfPresent(RC(col), body)
'     * economic objects: Swaps N:Y (YearFrac..Model_PnL) now call named builders
'                         (DiscountFactorFml, AnnuityFml, SwapModelPVFml, SwapDV01Fml,
'                          ModelPnLFml, SwapYearFracFml, SwapOisRateFml, FxToBaseFml)
'                         in the "ECONOMIC FORMULA LIBRARY" section - style-agnostic,
'                         take ready-made refs ("$B5" or RC(...)), reusable anywhere.
'   Auditability: WriteSwapsCalculatedFormulas is called in Step 3F after the
'   Bloomberg swap inputs are written/refreshed. Step 5 may re-write the same
'   formulas as part of the final PNL recalculation.
'   Helpers: RC(col), ColIdx(letters), WrapIfPresent(guard, body), DiffFormula(a, b [,scaled]).
'   NOT yet applied (optional): UDF lifts; merging the 4 date-source builder pairs.
' =============================================================================




' =============================================================================
' CONFIG CELLS
' =============================================================================


' -------------------------------------------------------------------------
' DATE CONVENTION WARNING
'
' Legacy internal names:
'   CFG_T0_DATE = Config!B4 = user-facing T-1 / prior snapshot date
'   CFG_T1_DATE = Config!B5 = user-facing T0 / current reporting date
'
' Do not confuse legacy VBA T1/T0 naming with report labels T0/T-1.
' -------------------------------------------------------------------------


Private Const CFG_T0_DATE          As String = "B4"
Private Const CFG_T1_DATE          As String = "B5"
Private Const CFG_SELL_PORTS       As String = "B9"
Private Const CFG_BR               As String = "B10"
Private Const CFG_DSN              As String = "B13"
Private Const CFG_UID              As String = "B14"
Private Const CFG_ISIN_TYPE        As String = "B15"
Private Const CFG_SPREAD_FRAMEWORK As String = "B18"
Private Const CFG_FX_HEDGE_TOKENS   As String = "B19"
Private Const CFG_LAST_LOAD        As String = "B20"
Private Const CFG_BOND_CNT         As String = "B21"
Private Const CFG_SWAP_CNT         As String = "B22"
Private Const CFG_FUT_CNT          As String = "B23"
Private Const CFG_LAST_BBG         As String = "B24"


Private Const CFG_BASE_CCY            As String = "B27"
Private Const CFG_FX_FIX_SOURCE       As String = "B28"
Private Const CFG_FUT_PRICE_FIELD     As String = "B29"
Private Const CFG_BBG_TIMEOUT         As String = "B30"
Private Const CFG_YEARSLEFT_BASIS     As String = "B31"
Private Const CFG_CURVE_SOURCE        As String = "B32"
Private Const CFG_FUNDING_SOURCE      As String = "B33"
Private Const CFG_ACTUAL_PNL_SOURCE   As String = "B34"
Private Const CFG_DIAGNOSTICS_ENABLED As String = "B35"
Private Const CFG_CONVEXITY_BUMP_BP As String = "B36"
Private Const CFG_KEEP_FORMULAS As String = "B39" ' Config!B39 = TRUE -> offline/no-BBG: keep all formulas, skip freezes & BBG waits
Private Const CFG_T0_FROZEN_AT As String = "B49"  ' set by Step 4 when the T0 snapshot is frozen; cleared when formulas are rewritten over it
Private Const CFG_T0_SNAPSHOT_CUTOFF As String = "B50" ' resolved prior-day cut-off datetime (audit only); NEVER write this back into B4 - see GetOrSetT0SnapshotDateTime
Private Const CFG_AUTO_DATE As String = "B6"    ' Config!B6  = "AUTO" -> auto-populate dates
Private Const CFG_SQL_PWD   As String = "B16"   ' Config!B16 = optional SQL password cell






' =============================================================================
' SHEETS / ROWS
' =============================================================================


Private Const SH_CONFIG As String = "Config"
Private Const SH_OIS    As String = "OIS_Curves"
Private Const SH_BONDS  As String = "Bonds"
Private Const SH_SWAPS  As String = "Swaps"
Private Const SH_FUTURES As String = "Futures"
Private Const SH_PNL    As String = "PNL_Attribution"

' Prefix for the workbook names PublishPnlColumnNames creates.  modDashboard
' has the same constant; the two must agree, and nothing else may use it.
Private Const PNL_NAME_PREFIX As String = "Pnl_"
Private Const SH_DASH   As String = "Dashboard"
Private Const SH_DIAG   As String = "Diagnostics"
Private Const SH_FUTMAP As String = "FutMap"      ' OPICS code -> Bloomberg generic
Private Const SH_T0_STG As String = "BBG_T0_Staging" ' hidden support sheet for T0 BDH array pulls


Private Const SH_SWAPMAP As String = "SwapMap"
Private Const SH_SPREADOVR As String = "SpreadOverride"   ' manual per-bond spread framework override (ISIN -> framework)


' Every framework code PNL_Attribution!CF may resolve to.  Both SWITCHes that
' consume CF (PnL_Duration_Total and SpreadPnL_Used) carry an arm for each of
' these plus a "" default, so no code can produce #N/A.  Keep the three in sync.
' Attribution_Status thresholds.  Kept here rather than inline so the desk can
' retune them in one place and so each has a stated reason.
'
'   HEDGE_RATIO_TOL   a hedge ratio of exactly 1.00 is perfect, so "over-hedged"
'                     must not fire on rounding.  2% of the bond DV01.
'   RESIDUAL_PCT_TOL  fraction of official PnL above which the residual is worth
'                     a human look.
'   IDENTITY_TOL_PCT  fraction of PnL_Duration_Total above which the framework
'                     chain is treated as not tying.
'   IDENTITY_TOL_EUR  absolute floor for the above, so tiny positions do not
'                     trip the check on pure rounding noise.
Private Const HEDGE_RATIO_TOL As String = "0.02"
Private Const RESIDUAL_PCT_TOL As String = "0.1"
Private Const IDENTITY_TOL_PCT As String = "0.01"
Private Const IDENTITY_TOL_EUR As String = "1"

' Smallest share of combined hedge BPV a leg must carry before the automatic
' spread framework treats the hedge as genuinely split and classifies MIXED.
' 0.2 = the minor leg has to be at least a fifth of the hedge; below that the
' dominant leg is what the residual risk is really against.
Private Const MIXED_MIN_SHARE As String = "0.2"


Private Const FRAMEWORK_CODES_ARRAY As String = _
    "{""G"",""I"",""ASW"",""Z"",""OAS"",""OIS"",""SOFR"",""MIXED"",""REVIEW""}"
Private Const SH_COV_SUPPORT As String = "Coverage support"
Private Const SH_COV_FUTURES As String = "CoverageFutures"
Private Const SH_COV_TOTAL_SUPPORT As String = "Coverage Total support"


Private Const SWAP_ID_BOOK_PATH As String = "M:\P_Pires\TRADING\Hedge_Risco Tx Juro.xlsm"
Private Const SWAP_ID_BOOK_SHEET As String = "Resumo"


Private Const SWAP_ID_FIRST_ROW As Long = 9
Private Const SWAP_ID_LAST_ROW As Long = 220


Private Const HEDGE_RISCO_COPY_LAST_COL As String = "AI"
Private Const HEDGE_RISCO_FORMULA_A_COL As String = "AJ"
Private Const HEDGE_RISCO_PORTFOLIO_TYPE_COL As String = "L"


' =============================================================================
' SECOND COVERAGE BOOK - HEDGE RISCO TOTAL
'
' The desk's futures hedges are split across TWO coverage files.  Hedge_Risco
' Tx Juro (above) carries the interest-rate hedges; Hedge_Risco Total carries
' the rest, and nothing in the first file knows about the second.  Loading only
' one of them makes every bond hedged out of the other read as UNHEDGED, which
' then picks the wrong spread framework and books the whole futures PnL as
' unexplained - so the two are loaded together and tagged by source.
'
' Resumo layout in this book (it is NOT the same as Tx Juro):
'
'   A       #RC, the coverage relation
'   B:K     the element being COVERED    B ISIN, C name, D maturity,
'           E notional, F bpv, G start px, H px now, I MtM, J p/d, K net
'   L:W     the instrument of COVERAGE   L type, M MD, N counterparty,
'           O years, P notional, Q start date, R ccy, S bpv, T MtM,
'           U accrued, V daily, W net
'
' Rows 8..34 are scanned in full because the block is NOT contiguous - blank
' rows sit between groups, so stopping at the first empty row would silently
' drop everything below it.  The header row is inside that range and is
' rejected by the HtC&S filter like any other non-qualifying row.
' =============================================================================
Private Const HEDGE_TOTAL_BOOK_PATH As String = "M:\P_Pires\TRADING\Hedge_Risco Total.xlsm"
Private Const HEDGE_TOTAL_BOOK_SHEET As String = "Resumo"


Private Const HEDGE_TOTAL_FIRST_ROW As Long = 8
Private Const HEDGE_TOTAL_LAST_ROW As Long = 34


Private Const HEDGE_TOTAL_COPY_LAST_COL As String = "W"
Private Const HEDGE_TOTAL_FORMULA_A_COL As String = "X"


' The column carrying the accounting classification to filter on.
'
' Stated as its own constant because the two books do not agree: Tx Juro holds
' it in L, this one in M.  If a future load returns zero rows, this is the first
' cell to check against the file - it is a one-line change, not a code change.
Private Const HEDGE_TOTAL_PORTFOLIO_TYPE_COL As String = "M"


' Coverage-instrument fields, named so the mapping above is auditable in code.
Private Const HEDGE_TOTAL_TYPE_COL As String = "L"          ' instrument type / label
Private Const HEDGE_TOTAL_COUNTERPARTY_COL As String = "N"
Private Const HEDGE_TOTAL_NOTIONAL_COL As String = "P"
Private Const HEDGE_TOTAL_STARTDATE_COL As String = "Q"
Private Const HEDGE_TOTAL_CCY_COL As String = "R"
Private Const HEDGE_TOTAL_BPV_COL As String = "S"




Private Const DATA_ROW      As Long = 5
' Bonds!1..BOND_COMMENT_ROWS are reserved for the desk's comments and are never
' written by this module.  The header and first data row are DERIVED from that,
' so widening the comment band is a one-line change here - and the accompanying
' EnsureBondsCommentRows moves the OPICS query table to match, which is the part
' that cannot be done by editing a constant alone.
Private Const BOND_COMMENT_ROWS As Long = 2
Private Const BOND_HEADER_ROW As Long = BOND_COMMENT_ROWS + 1
Private Const BOND_DATA_ROW As Long = BOND_HEADER_ROW + 1
Private Const FUT_HEADER_ROW   As Long = 4   ' Futures header row
Private Const SWAP_HEADER_ROW  As Long = 4   ' Swaps header row
Private Const CURVE_HEADER_ROW As Long = 4   ' OIS_Curves header row
Private Const PNL_HEADER_ROW   As Long = 4   ' PNL_Attribution header row
' -----------------------------------------------------------------------------
' POSITION COUNTS ARE NOT CONSTANTS
'
' The book does not hold a fixed number of bonds, futures or swaps - the counts
' change with every load, and OPICS / Hedge Risco decide them, not this module.
' The old N_BOND_ROWS = 600 / N_FUT_ROWS = 220 / N_SWAP_ROWS = 200 constants were
' used for three different jobs at once, and each one broke differently once the
' book outgrew them: rows past the limit were never scanned, never cleared, and
' silently never written.
'
' They are replaced by three explicit questions, answered from the sheet:
'
'   Last*DataRow(ws)          the last row that actually HOLDS a position
'   SheetClearLastRow(...)    how far to clear/scan - reaches past today's data
'                             so rows left by a LARGER previous run are removed
'   MAX_SHEET_ROWS            a runaway guard only, never a capacity limit
' -----------------------------------------------------------------------------

' Pure runaway guard: a corrupt source must not be able to drive a write loop
' for a million rows.  Set far above any plausible book; if a real run ever
' approaches it, raise it - do not start truncating positions.
Private Const MAX_SHEET_ROWS As Long = 50000
Private Const PNL_LAST_COL As String = "CI"
Private Const PNL_CLEAR_LAST_COL As String = "CR"


Private Const CURVE_FIRST_ROW As Long = 7
Private Const CURVE_LAST_ROW  As Long = 19


Private Const COL_BBG_TICKER As Long = 41       ' AO  (= ColNum(BCOL_BBG_TICKER))
Private Const COL_BBG_TICKER_STATUS As Long = 42 ' AP  (= ColNum(BCOL_TICKER_STATUS))


Private Const COL_BBG_CAND_FIRST As Long = 72   ' BT  (= ColNum(BCOL_BBG_CAND_ISIN))
Private Const COL_BBG_CAND_LAST  As Long = 81   ' CC  (= ColNum(BCOL_BBG_CAND_MMKT))

' The deepest a generated bond formula may nest.
'
' Excel stops at 64 levels of nested functions, and a multi-field fallback
' formula costs 6 + 2 * candidates * fields.  BondFallbackColCount solves that
' for the candidate count, against this rather than against 64, so a future
' field added to a list cannot silently push a formula over the real limit.
Private Const BBG_MAX_FORMULA_NESTING As Long = 60
' Rows of bond T0 BDH formulas alive on the staging sheet at once in Step 4.  Smaller =
' fewer concurrent BDH array requests held in the BLP add-in -> lower peak memory (matters
' on 32-bit Excel's 2 GB ceiling).  Reduced 50 -> 25 as part of the Step 4 OOM fix.
Private Const T0_BOND_BATCH_ROWS As Long = 25

' Runaway guard for the pull-to-par cash-flow schedule.  A quarterly 100-year
' bond has 400 flows; anything past this means the maturity or the frequency is
' corrupt and the schedule loop must not be allowed to spin.
Private Const MAX_BOND_CASHFLOWS As Long = 1200






' =============================================================================
' THE POSITION STORE'S VOCABULARY
'
' Which population is being stored or restored.  PUBLIC because modAccess names
' them too - it asks this module for "BOND" and hands back rows for "BOND", and
' a shared spelling that is a constant in both places cannot drift into a typo
' that silently matches nothing.  The bridge itself is at the bottom of the
' file, under THE POSITION STORE BRIDGE.
'
' Declared HERE, with the other module constants, because VBA requires every
' module-level Const to precede the first procedure.  Put one further down and
' the module does not compile.
' =============================================================================
Public Const STORE_KIND_BOND As String = "BOND"
Public Const STORE_KIND_SWAP As String = "SWAP"
Public Const STORE_KIND_FUTURE As String = "FUTURE"

' A field whose value is the row number the position sat on, rather than the
' contents of a column.  Stands where a column letter would in the field map.
Private Const STORE_ROWNUM_TOKEN As String = "#ROW"


' =============================================================================
' COLUMN CONSTANTS  (one letter-string per sheet column; single source of truth)
'
'   Address builders use these directly:  ws.Range(BCOL_BBG_TICKER & r)
'   R1C1 formula interiors are intentionally left literal (see notebook / docs);
'   ColNum(BCOL_x) converts a column letter to its number where a number is needed.
'   To shift a column, edit ONLY its constant here.
' =============================================================================


' --- Bonds columns ---
Private Const BCOL_ISIN As String = "A"   ' col 1 = ISIN
Private Const BCOL_NAME As String = "B"   ' col 2 = Name
Private Const BCOL_CCY As String = "C"   ' col 3 = CCY
Private Const BCOL_COUPON      As String = "D"   ' col 4 = Coupon (query-owned; used in convexity/carry formulas)
Private Const BCOL_COUPON_FREQ As String = "E"   ' col 5 = CouponFreq descr (query-owned; used by CouponFreqNum)
Private Const BCOL_MATURITY As String = "F"   ' col 6 = Maturity
Private Const BCOL_NOTIONAL As String = "G"   ' col 7 = Notional
Private Const BCOL_ACCTGCAT As String = "I"   ' col 9 = AcctgCat
Private Const BCOL_PORTFOLIO As String = "J"   ' col 10 = Portfolio
Private Const BCOL_BOOKVAL As String = "K"   ' col 11 = BookVal

' -----------------------------------------------------------------------------
' WHERE THE OPICS QUERY STOPS AND THE MACRO STARTS
'
' Bonds!A:BONDS_QUERY_LAST_COL is written by the Excel query, not by VBA.  The
' macro owns everything to its right and must begin in the very next column: a
' one-column disagreement writes DaysLeft over a query column, or leaves a blank
' stripe, and every bond row's two halves come apart - silently, because both
' halves still look populated.
'
' PositionSide used to sit in column L and nothing read it.  Removing it moved
' the macro block one column left, so THE QUERY MUST NO LONGER SELECT
' PositionSide.  LoadOPICS_Bonds refuses to run if the query still returns it
' (see BondsLayoutIsSane), rather than loading a whole book of misaligned rows.
Private Const BONDS_QUERY_LAST_COL As String = "K"   ' col 11 = BookVal
Private Const BCOL_DAYSLEFT As String = "L"   ' col 12 = DaysLeft
Private Const BCOL_OIS_T0 As String = "M"   ' col 13 = OIS_T0
Private Const BCOL_OIS_TM1 As String = "N"   ' col 14 = OIS_T-1
Private Const BCOL_DF_T0 As String = "O"   ' col 15 = DF_T0
Private Const BCOL_DF_TM1 As String = "P"   ' col 16 = DF_T-1
Private Const BCOL_FX_T0 As String = "Q"   ' col 17 = FX_T0
Private Const BCOL_FX_TM1 As String = "R"   ' col 18 = FX_T-1
Private Const BCOL_CLEANPX_T0 As String = "S"   ' col 19 = CleanPx_T0
Private Const BCOL_DIRTYPX_T0 As String = "T"   ' col 20 = DirtyPx_T0
Private Const BCOL_ACCRUEDINTEREST_T0 As String = "U"   ' col 21 = AccruedInterest_T0
Private Const BCOL_YTM_T0 As String = "V"   ' col 22 = YTM_T0
Private Const BCOL_MODDUR_T0 As String = "W"   ' col 23 = ModDur_T0
Private Const BCOL_CONVEXITY As String = "X"   ' col 24 = Convexity
Private Const BCOL_ZSPRD_T0 As String = "Y"   ' col 25 = ZSprd_T0
Private Const BCOL_ASW_T0 As String = "Z"   ' col 26 = ASW_T0
Private Const BCOL_CLEANPX_TM1 As String = "AA"   ' col 27 = CleanPx_T-1
Private Const BCOL_DIRTYPX_TM1 As String = "AB"   ' col 28 = DirtyPx_T-1
Private Const BCOL_YTM_TM1 As String = "AC"   ' col 29 = YTM_T-1
Private Const BCOL_ZSPRD_TM1 As String = "AD"   ' col 30 = ZSprd_T-1
Private Const BCOL_ASW_TM1 As String = "AE"   ' col 31 = ASW_T-1
Private Const BCOL_DV01_EUR As String = "AF"   ' col 32 = DV01_EUR
Private Const BCOL_DIRTYMV_T0_EUR As String = "AG"   ' col 33 = DirtyMV_T0_EUR
Private Const BCOL_DIRTYMV_TM1_EUR As String = "AH"   ' col 34 = DirtyMV_T-1_EUR
Private Const BCOL_BOOKVAL_EUR2 As String = "AI"   ' col 35 = BookVal_EUR2
Private Const BCOL_ISPREAD_T0 As String = "AJ"   ' col 36 = ISpread_T0
Private Const BCOL_ISPREAD_TM1 As String = "AK"   ' col 37 = ISpread_T-1
Private Const BCOL_GSPREAD_T0 As String = "AL"   ' col 38 = GSpread_T0
Private Const BCOL_GSPREAD_TM1 As String = "AM"   ' col 39 = GSpread_T-1
Private Const BCOL_ISPREAD_STATUS As String = "AN"   ' col 40 = ISpread_Status
Private Const BCOL_BBG_TICKER As String = "AO"   ' col 41 = BBG_Ticker
Private Const BCOL_TICKER_STATUS As String = "AP"   ' col 42 = Ticker_Status
Private Const BCOL_OAS_T0 As String = "AQ"   ' col 43 = OAS_T0
Private Const BCOL_OAS_TM1 As String = "AR"   ' col 44 = OAS_T-1
Private Const BCOL_DELTAOAS As String = "AS"   ' col 45 = DeltaOAS
Private Const BCOL_DV01_UNIT As String = "AT"   ' col 46 = DV01_Unit
Private Const BCOL_SPREADDURATION_USED As String = "AU"   ' col 47 = SpreadDuration_Used
Private Const BCOL_OAS_MODDURATION_RAW As String = "AV"   ' col 48 = OAS_ModDuration_Raw
Private Const BCOL_OAS_CONVEXITY As String = "AW"   ' col 49 = OAS_Convexity
Private Const BCOL_PRICING_SOURCE As String = "AX"   ' col 50 = Pricing_Source
Private Const BCOL_LAST_PRICING_DATE As String = "AY"   ' col 51 = Last_Pricing_Date
Private Const BCOL_BENCHMARK_BOND As String = "AZ"   ' col 52 = Benchmark_Bond
Private Const BCOL_BENCHMARK_NAME As String = "BA"   ' col 53 = Benchmark_Name
Private Const BCOL_GOV_T0 As String = "BB"   ' col 54 = Gov_T0
Private Const BCOL_GOV_TM1 As String = "BC"   ' col 55 = Gov_T-1
Private Const BCOL_SWAP_T0 As String = "BD"   ' col 56 = Swap_T0
Private Const BCOL_SWAP_TM1 As String = "BE"   ' col 57 = Swap_T-1
Private Const BCOL_G_T0 As String = "BF"   ' col 58 = g_T0
Private Const BCOL_G_TM1 As String = "BG"   ' col 59 = g_T-1
Private Const BCOL_DELTA_G_BP As String = "BH"   ' col 60 = Delta_g_bp
Private Const BCOL_Q_T0 As String = "BI"   ' col 61 = q_T0
Private Const BCOL_Q_TM1 As String = "BJ"   ' col 62 = q_T-1
Private Const BCOL_DELTA_Q_BP As String = "BK"   ' col 63 = Delta_q_bp
Private Const BCOL_DELTA_I_BP As String = "BL"   ' col 64 = Delta_i_bp
Private Const BCOL_DELTA_Y_BP As String = "BM"   ' col 65 = Delta_y_bp
Private Const BCOL_FUNDTKR As String = "BN"   ' col 66 = FundTkr
Private Const BCOL_FUNDRATE_T0 As String = "BO"   ' col 67 = FundRate_T0
Private Const BCOL_FUNDRATE_TM1 As String = "BP"   ' col 68 = FundRate_T-1
Private Const BCOL_BOND_STATUS As String = "BQ"   ' col 69 = Bond_Status
Private Const BCOL_SWAPLINK As String = "BR"   ' col 70 = SwapLink
Private Const BCOL_FUTLINK As String = "BS"   ' col 71 = FutLink
Private Const BCOL_BBG_CAND_ISIN As String = "BT"   ' col 72 = BBG_Cand_ISIN
Private Const BCOL_BBG_CAND_SLASHISIN As String = "BU"   ' col 73 = BBG_Cand_slashISIN
Private Const BCOL_BBG_CAND_CORP As String = "BV"   ' col 74 = BBG_Cand_Corp
Private Const BCOL_BBG_CAND_BVAL_CORP As String = "BW"   ' col 75 = BBG_Cand_BVAL_Corp
Private Const BCOL_BBG_CAND_BGN_CORP As String = "BX"   ' col 76 = BBG_Cand_BGN_Corp
Private Const BCOL_BBG_CAND_GOVT As String = "BY"   ' col 77 = BBG_Cand_Govt
Private Const BCOL_BBG_CAND_BVAL_GOVT As String = "BZ"   ' col 78 = BBG_Cand_BVAL_Govt
Private Const BCOL_BBG_CAND_BGN_GOVT As String = "CA"   ' col 79 = BBG_Cand_BGN_Govt
Private Const BCOL_BBG_CAND_MTGE As String = "CB"   ' col 80 = BBG_Cand_Mtge
Private Const BCOL_BBG_CAND_MMKT As String = "CC"   ' col 81 = BBG_Cand_MMkt
Private Const BCOL_CONVEXITY_BUMP_BP As String = "CD"   ' col 82 = Convexity_Bump_bp
Private Const BCOL_COUPONFREQ_NUM As String = "CE"   ' col 83 = CouponFreq_Num
Private Const BCOL_PRICE_BASE As String = "CF"   ' col 84 = Price_Base
Private Const BCOL_PRICE_UP As String = "CG"   ' col 85 = Price_Up
Private Const BCOL_PRICE_DOWN As String = "CH"   ' col 86 = Price_Down
Private Const BCOL_CONVEXITY_BUMP As String = "CI"   ' col 87 = Convexity_Bump
Private Const BCOL_DAY_CNT_DES As String = "CJ"   ' col 88 = DAY_CNT_DES
Private Const BCOL_BONDDCC_CODE As String = "CK"   ' col 89 = BondDCC_Code
Private Const BCOL_BONDDCC_NAME As String = "CL"   ' col 90 = BondDCC_Name
Private Const BCOL_DV01_OPENING_EUR As String = "CM"   ' col 91 = DV01_Opening_EUR

Private Const PCOL_ISIN As String = "A"   ' col 1 = ISIN
Private Const PCOL_NAME As String = "B"   ' col 2 = Name
Private Const PCOL_CCY As String = "C"   ' col 3 = CCY
Private Const PCOL_PORTFOLIO As String = "D"   ' col 4 = Portfolio
Private Const PCOL_ACCTGCAT As String = "E"   ' col 5 = AcctgCat
Private Const PCOL_NOTIONAL As String = "F"   ' col 6 = Notional

Private Const PCOL_MODDUR As String = "G"   ' col 7 = ModDur
Private Const PCOL_CONVEXITY As String = "H"   ' col 8 = Convexity
Private Const PCOL_SPREADDURATION As String = "I"   ' col 9 = SpreadDuration
Private Const PCOL_DAYS As String = "J"   ' col 10 = Days
Private Const PCOL_YEARFRAC As String = "K"   ' col 11 = YearFrac

Private Const PCOL_BOND_DV01_CURRENT As String = "L"   ' col 12 = Bond_DV01_Current
Private Const PCOL_BOND_DV01_CREDIT_SPREAD As String = "M"   ' col 13 = Bond_DV01_Credit_Spread
Private Const PCOL_HEDGE_DV01 As String = "N"   ' col 14 = Actual_Hedge_DV01
Private Const PCOL_PLAINSWAP_DV01 As String = "O"   ' col 15 = PlainSwap_DV01
Private Const PCOL_FUTURES_RTJ_DV01 As String = "P"   ' col 16 = FuturesRTJ_DV01
Private Const PCOL_FUTURES_RT_DV01 As String = "Q"   ' col 17 = FuturesRT_DV01
Private Const PCOL_HEDGE_DV01_GAP As String = "R"   ' col 18 = Hedge_DV01_Gap
Private Const PCOL_SYNTHETICSWAP_DV01 As String = "S"   ' col 19 = SyntheticSwap_DV01
Private Const PCOL_TARGET_HEDGE_DV01 As String = "T"   ' col 20 = Target_Hedge_DV01

Private Const PCOL_DELTA_DIRTY_MV_EUR As String = "U"   ' col 21 = Delta_Dirty_MV_EUR
Private Const PCOL_DELTA_Y_BP As String = "V"   ' col 22 = Delta_Y_bp
Private Const PCOL_DELTA_R_BP As String = "W"   ' col 23 = Delta_r_bp
Private Const PCOL_DELTA_GOV_BP As String = "X"   ' col 24 = Delta_Gov_bp
Private Const PCOL_DELTA_G_BP As String = "Y"   ' col 25 = Delta_g_bp
Private Const PCOL_DELTA_SWAP_BP As String = "Z"   ' col 26 = Delta_Swap_bp
Private Const PCOL_DELTA_Q_BP As String = "AA"   ' col 27 = Delta_q_bp
Private Const PCOL_DELTA_I_BP As String = "AB"   ' col 28 = Delta_i_bp
Private Const PCOL_DELTA_Z_BP As String = "AC"   ' col 29 = Delta_Z_bp
Private Const PCOL_DELTA_G_BP_T As String = "AD"   ' col 30 = Delta_GSpread_bp
Private Const PCOL_DELTA_ASW_BP As String = "AE"   ' col 31 = Delta_ASW_bp
Private Const PCOL_DELTA_OAS_BP As String = "AF"   ' col 32 = Delta_OAS_bp



Private Const PCOL_PNL_DURATION_TOTAL As String = "AG"   ' col 33 = PnL_Duration_Total
Private Const PCOL_PNL_OIS As String = "AH"   ' col 34 = PnL_OIS
Private Const PCOL_PNL_GOVBASIS As String = "AI"   ' col 35 = PnL_GovBasis
Private Const PCOL_PNL_SWAPGOVBASIS As String = "AJ"   ' col 36 = PnL_SwapGovBasis
Private Const PCOL_PNL_CREDIT_ISPREAD As String = "AK"   ' col 37 = PnL_Credit_Ispread
Private Const PCOL_PNL_CONVEXITY As String = "AL"   ' col 38 = PnL_Convexity

Private Const PCOL_CARRY_COUPON As String = "AM"   ' col 39 = Carry_Coupon
Private Const PCOL_CARRY_ROLLTOPAR As String = "AN"   ' col 40 = Carry_RollToPar
Private Const PCOL_CARRY_FUNDING As String = "AO"   ' col 41 = Funding_Carry_Memo
Private Const PCOL_CARRY_TOTAL As String = "AP"   ' col 42 = Carry_Total
Private Const PCOL_PNL_ZSPREAD As String = "AQ"   ' col 43 = PnL_ZSpread
Private Const PCOL_PNL_GSPREAD As String = "AR"   ' col 44 = PnL_GSpread
Private Const PCOL_PNL_ASW As String = "AS"   ' col 45 = PnL_ASW
Private Const PCOL_PNL_OAS As String = "AT"   ' col 46 = PnL_OAS
Private Const PCOL_SPREADPNL_USED As String = "AU"   ' col 47 = SpreadPnL_Used
Private Const PCOL_PNL_FX As String = "AV"   ' col 48 = PnL_FX

Private Const PCOL_PNL_FUTURES As String = "AW"   ' col 49 = Futures_Gov_Model_PnL
Private Const PCOL_PNL_SWAP As String = "AX"   ' col 50 = Swap_Curve_Model_PnL
Private Const PCOL_TOTAL_HEDGE As String = "AY"   ' col 51 = Hedge_Curve_Model_PnL
Private Const PCOL_BASIS_PNL As String = "AZ"   ' col 52 = Hedge_Model_Residual_PnL
Private Const PCOL_TOTAL_EXPLAINED As String = "BA"   ' col 53 = Total_Model_Explained
Private Const PCOL_OFFICIAL_PNL As String = "BB"   ' col 54 = Official_Total_PnL
Private Const PCOL_RESIDUAL As String = "BC"   ' col 55 = Unexplained_Residual_PnL
Private Const PCOL_RESIDUAL_AX As String = "BD"   ' col 56 = Unexplained_Residual_Pct

Private Const PCOL_RESIDUAL_DV01 As String = "BE"   ' col 57 = Residual_DV01
Private Const PCOL_HEDGE_RATIO As String = "BF"   ' col 58 = Hedge_Ratio
Private Const PCOL_HEDGE_EFFICIENCY As String = "BG"   ' col 59 = Hedge_Efficiency
Private Const PCOL_ATTRIBUTION_STATUS As String = "BH"   ' col 60 = Attribution_Status
Private Const PCOL_THEORETICALHEDGE_DV01 As String = "BI"   ' col 61 = Synthetic_Alternative_DV01

' --- Matching, raw PnL and framework diagnostics ---

Private Const PCOL_FUT_MATCH_COUNT As String = "BJ"   ' col 62 = Futures_Match_Count
Private Const PCOL_FUT_PNL_RAW As String = "BK"   ' col 63 = Actual_Futures_PnL
Private Const PCOL_SWAP_ALL_MATCH_COUNT As String = "BL"   ' col 64 = Swap_All_Match_Count
Private Const PCOL_SWAP_PLAIN_MATCH_COUNT As String = "BM"   ' col 65 = PlainSwap_Match_Count
Private Const PCOL_SWAP_PLAIN_PNL_RAW As String = "BN"   ' col 66 = Actual_PlainSwap_PnL
Private Const PCOL_SWAP_SYNTHETIC_MATCH_COUNT As String = "BO"   ' col 67 = SyntheticSwap_Match_Count
Private Const PCOL_SWAP_SYNTHETIC_PNL_RAW As String = "BP"   ' col 68 = Actual_SyntheticSwap_PnL
Private Const PCOL_SWAP_ALL_PNL_RAW As String = "BQ"   ' col 69 = AllSwapRows_PnL_Diagnostic
Private Const PCOL_ACTUAL_FUTURES_PNL_RAW As String = "BR"   ' col 70 = Actual_Futures_PnL_Check
Private Const PCOL_FUTURES_GOV_MODEL_PNL As String = "BS"   ' col 71 = Futures_Gov_Model_PnL_Check
Private Const PCOL_FUTURES_BASIS_PNL As String = "BT"   ' col 72 = Futures_Model_Residual_PnL
Private Const PCOL_ACTUAL_SWAP_PNL_RAW As String = "BU"   ' col 73 = Actual_PlainSwap_PnL_Check
Private Const PCOL_SWAP_CURVE_MODEL_PNL As String = "BV"   ' col 74 = Swap_Curve_Model_PnL_Check
Private Const PCOL_SWAP_BASIS_PNL As String = "BW"   ' col 75 = Swap_Model_Residual_PnL
Private Const PCOL_ACTUAL_HEDGE_PNL_RAW As String = "BX"   ' col 76 = Actual_Hedge_PnL
Private Const PCOL_HEDGE_BASIS_PNL As String = "BY"   ' col 77 = Hedge_Model_Residual_PnL_Check
Private Const PCOL_SPREAD_FRAMEWORK_AUTO As String = "BZ"   ' col 78 = Spread_Framework_Auto
Private Const PCOL_SPREAD_FRAMEWORK_REASON As String = "CA"   ' col 79 = Spread_Framework_Reason
Private Const PCOL_DURATION_IDENTITY_CHECK As String = "CB"   ' col 80 = Duration_Identity_Check

' --- row quarantine ---
Private Const PCOL_ROW_VALID As String = "CC"   ' col 81 = Row_Valid
Private Const PCOL_ROW_EXCLUSION_REASON As String = "CD"   ' col 82 = Row_Exclusion_Reason

' --- FX hedge reference ---
Private Const PCOL_FX_EXPOSURE_EUR As String = "CE"   ' col 83 = FX_Exposure_EUR

' --- opening-risk anchor for the attribution ---
Private Const PCOL_BOND_DV01_OPENING As String = "CF"   ' col 84 = Bond_DV01_Opening
Private Const PCOL_RISK_TIMING_BIAS As String = "CG"   ' col 85 = Risk_Timing_Bias
Private Const PCOL_COUPON_PAID_EUR As String = "CH"   ' col 86 = Coupon_Paid_EUR

' --- Futures columns ---
Private Const FCOL_CONTRACTCODE As String = "A"   ' col 1 = ContractCode
Private Const FCOL_EXCHANGE As String = "B"   ' col 2 = Exchange
Private Const FCOL_CCY As String = "C"   ' col 3 = CCY
Private Const FCOL_CONTRACTS As String = "D"   ' col 4 = Contracts
Private Const FCOL_FACEVALUE As String = "E"   ' col 5 = FaceValue
Private Const FCOL_DELIVDATE As String = "F"   ' col 6 = DelivDate
Private Const FCOL_CTD_ISIN As String = "G"   ' col 7 = CTD_ISIN
Private Const FCOL_CTD_CF As String = "H"   ' col 8 = CTD_CF
Private Const FCOL_PORTFOLIO As String = "I"   ' col 9 = Portfolio
Private Const FCOL_LINKEDISIN As String = "J"   ' col 10 = LinkedISIN
Private Const FCOL_AVGENTRYPX As String = "K"   ' col 11 = AvgEntryPx
Private Const FCOL_DAYSTODELIV As String = "L"   ' col 12 = DaysToDeliv
Private Const FCOL_FX_T0 As String = "M"   ' col 13 = FX_T0
Private Const FCOL_FUTPX_T0 As String = "N"   ' col 14 = FutPx_T0
Private Const FCOL_CTD_DIRTYPX_T0 As String = "O"   ' col 15 = CTD_DirtyPx_T0
Private Const FCOL_CTD_TICKER_TM1 As String = "P"   ' col 16 = CTD_Ticker_T-1
Private Const FCOL_FUTPX_TM1 As String = "Q"   ' col 17 = FutPx_T-1
Private Const FCOL_IMPLIEDREPO_CALC As String = "R"   ' col 18 = ImpliedRepo_Calc
Private Const FCOL_GROSSBASIS_CALC As String = "S"   ' col 19 = GrossBasis_Calc
Private Const FCOL_NOTIONALVALUE_EUR As String = "T"   ' col 20 = NotionalValue_EUR
Private Const FCOL_FUTURESPNL_EUR As String = "U"   ' col 21 = FuturesPnL_EUR
Private Const FCOL_BBG_TICKER As String = "V"   ' col 22 = BBG_Ticker
Private Const FCOL_STATUS_T0 As String = "W"   ' col 23 = Status_T0
Private Const FCOL_STATUS_TM1 As String = "X"   ' col 24 = Status_T-1
Private Const FCOL_FUT_VAL_PT As String = "Y"   ' col 25 = FUT_VAL_PT
Private Const FCOL_CTD_TICKER As String = "Z"   ' col 26 = CTD_Ticker
Private Const FCOL_CONVFACTOR As String = "AA"   ' col 27 = ConvFactor
Private Const FCOL_HEDGEUNITDV01 As String = "AB"   ' col 28 = HedgeUnitDV01
Private Const FCOL_FUTURES_DV01_EUR As String = "AC"   ' col 29 = Futures_DV01_EUR
Private Const FCOL_IMPLIEDREPO_BBG As String = "AD"   ' col 30 = ImpliedRepo_BBG
Private Const FCOL_NETBASIS_BBG As String = "AE"   ' col 31 = NetBasis_BBG
Private Const FCOL_GROSSBASIS_BBG As String = "AF"   ' col 32 = GrossBasis_BBG
Private Const FCOL_STATUS As String = "AG"   ' col 33 = Status
Private Const FCOL_COVERAGE_SOURCEROW As String = "AH"   ' col 34 = Coverage_SourceRow
Private Const FCOL_COVERAGERELATION As String = "AI"   ' col 35 = CoverageRelation
Private Const FCOL_FUTURELABEL As String = "AJ"   ' col 36 = FutureLabel
Private Const FCOL_HEDGETYPE As String = "AK"   ' col 37 = HedgeType
Private Const FCOL_COVERAGE_STARTDATE As String = "AL"   ' col 38 = Coverage_StartDate
Private Const FCOL_COVERAGEINFO_D As String = "AM"   ' col 39 = CoverageInfo_D
Private Const FCOL_HEDGE_SOURCE As String = "AN"   ' col 40 = Hedge_Source
Private Const FCOL_COVERAGE_BPV As String = "AO"   ' col 41 = Coverage_BPV
Private Const FCOL_HEDGE_CLASS As String = "AP"   ' col 42 = Hedge_Class

' Which coverage file a futures row came from.  The desk keeps its futures
' hedges in two books and PNL_Attribution reports the risk from each
' separately, so every row has to say which one it is.  These two strings are
' the join key behind FuturesRTJ_DV01 / FuturesRT_DV01 - change one and the
' SUMIFS behind that column silently returns zero, so they are named once.
Private Const HEDGE_SOURCE_RTJ As String = "RTJ"   ' Hedge_Risco Tx Juro.xlsm
Private Const HEDGE_SOURCE_RT As String = "RT"     ' Hedge_Risco Total.xlsm

' What a futures/forward row on the Futures sheet actually hedges.  RATES is
' the bond-future case the DV01 machinery is built for; FX is a currency
' hedge, whose DV01 against a yield curve is meaningless and must not be
' added to a bond's rates hedge.
Private Const HEDGE_CLASS_RATES As String = "RATES"
Private Const HEDGE_CLASS_FX As String = "FX"


' --- Swaps columns ---
Private Const WCOL_DEALID As String = "A"   ' col 1 = DealID
Private Const WCOL_CCY As String = "B"   ' col 2 = Ccy
Private Const WCOL_NOTIONAL As String = "C"   ' col 3 = Notional
Private Const WCOL_FIXEDRATE As String = "D"   ' col 4 = FixedRate
Private Const WCOL_FLOATINDEX As String = "E"   ' col 5 = FloatIndex
Private Const WCOL_FLOATSPREAD As String = "F"   ' col 6 = FloatSpread
Private Const WCOL_STARTDATE As String = "G"   ' col 7 = StartDate
Private Const WCOL_ENDDATE As String = "H"   ' col 8 = EndDate
Private Const WCOL_PAYFIXED As String = "I"   ' col 9 = PayFixed
Private Const WCOL_PORTFOLIO As String = "J"   ' col 10 = Portfolio
Private Const WCOL_FLOATCURVE_TYPE As String = "K"   ' col 11 = FloatCurve_Type
Private Const WCOL_LINKEDISIN As String = "L"   ' col 12 = LinkedISIN
Private Const WCOL_CPTY As String = "M"   ' col 13 = Cpty
Private Const WCOL_YEARFRAC As String = "N"   ' col 14 = YearFrac
Private Const WCOL_OIS_T0 As String = "O"   ' col 15 = OIS_T0
Private Const WCOL_OIS_TM1 As String = "P"   ' col 16 = OIS_T-1
Private Const WCOL_DF_T0 As String = "Q"   ' col 17 = DF_T0
Private Const WCOL_DF_TM1 As String = "R"   ' col 18 = DF_T-1
Private Const WCOL_ANNUITY_T0 As String = "S"   ' col 19 = Annuity_T0
Private Const WCOL_ANNUITY_TM1 As String = "T"   ' col 20 = Annuity_T-1
Private Const WCOL_FX As String = "U"   ' col 21 = FX
Private Const WCOL_PV_T0_MODEL As String = "V"   ' col 22 = PV_T0_Model
Private Const WCOL_PV_TM1_MODEL As String = "W"   ' col 23 = PV_T-1_Model
Private Const WCOL_SWAP_DV01_EUR As String = "X"   ' col 24 = Swap_DV01_EUR
Private Const WCOL_MODEL_PNL As String = "Y"   ' col 25 = Model_PnL
Private Const WCOL_STATUS As String = "Z"   ' col 26 = Status
Private Const WCOL_PNL_SOURCE As String = "AA"   ' col 27 = PnL_Source
Private Const WCOL_PNL As String = "AB"   ' col 28 = PnL
Private Const WCOL_AF_STATUS As String = "AC"   ' col 29 = AF_Status
Private Const WCOL_BBG_SWAP_DIRECT_ID As String = "AD"   ' col 30 = BBG_Swap_Direct_ID
Private Const WCOL_BBG_FIXED_LEG_ID As String = "AE"   ' col 31 = BBG_Fixed_Leg_ID
Private Const WCOL_BBG_FLOAT_LEG_ID As String = "AF"   ' col 32 = BBG_Float_Leg_ID
Private Const WCOL_SWAP_ID_SOURCE As String = "AG"   ' col 33 = Swap_ID_Source
Private Const WCOL_COVERAGERELATION As String = "AH"   ' col 34 = CoverageRelation
Private Const WCOL_SWAPMAP_SOURCEROW As String = "AI"   ' col 35 = SwapMap_SourceRow
Private Const WCOL_SWAPMAP_CLASS As String = "AJ"   ' col 36 = SwapMap_Class
Private Const WCOL_SWAPMAP_STATUS As String = "AK"   ' col 37 = SwapMap_Status
Private Const WCOL_MAP_NOTIONAL As String = "AL"   ' col 38 = Map_Notional
Private Const WCOL_MAP_CCY As String = "AM"   ' col 39 = Map_CCY
Private Const WCOL_MAP_COUNTERPARTY As String = "AN"   ' col 40 = Map_Counterparty
Private Const WCOL_NOTIONAL_FINAL As String = "AO"   ' col 41 = Notional_Final
Private Const WCOL_NOTIONAL_SOURCE As String = "AP"   ' col 42 = Notional_Source
Private Const WCOL_BQL_NPV_DIRECT_T0 As String = "AQ"   ' col 43 = BQL_NPV_Direct_T0
Private Const WCOL_BQL_NPV_FIXED_T0 As String = "AR"   ' col 44 = BQL_NPV_Fixed_T0
Private Const WCOL_BQL_NPV_FLOAT_T0 As String = "AS"   ' col 45 = BQL_NPV_Float_T0
Private Const WCOL_BQL_NPV_TOTAL_T0 As String = "AT"   ' col 46 = BQL_NPV_Total_T0
Private Const WCOL_BQL_NPV_DIRECT_TM1 As String = "AU"   ' col 47 = BQL_NPV_Direct_T-1
Private Const WCOL_BQL_NPV_FIXED_TM1 As String = "AV"   ' col 48 = BQL_NPV_Fixed_T-1
Private Const WCOL_BQL_NPV_FLOAT_TM1 As String = "AW"   ' col 49 = BQL_NPV_Float_T-1
Private Const WCOL_BQL_NPV_TOTAL_TM1 As String = "AX"   ' col 50 = BQL_NPV_Total_T-1
Private Const WCOL_BQL_SWAP_PNL As String = "AY"   ' col 51 = BQL_Swap_PnL
Private Const WCOL_BQL_SWAP_STATUS As String = "AZ"
' col 52 = BQL_Swap_Status
Private Const WCOL_PAY_FLT_RATE_IDX As String = "BA"   ' col 53 = PAY_FLT_RATE_IDX
Private Const WCOL_FLOATINDEX_FAMILY As String = "BB"   ' col 54 = FloatIndex_Family
Private Const WCOL_FLOATINDEX_TENOR As String = "BC"   ' col 55 = FloatIndex_Tenor
Private Const WCOL_FLOATCURVE_T0 As String = "BD"   ' col 56 = FloatCurve_T0
Private Const WCOL_FLOATCURVE_TM1 As String = "BE"   ' col 57 = FloatCurve_T-1
Private Const WCOL_MODELSPREAD_T0 As String = "BF"   ' col 58 = ModelSpread_T0
Private Const WCOL_MODELSPREAD_TM1 As String = "BG"   ' col 59 = ModelSpread_T-1
Private Const WCOL_DELTA_FLOATCURVE_BP As String = "BH"   ' col 60 = Delta_FloatCurve_bp
Private Const WCOL_DELTA_MODELSPREAD_BP As String = "BI"   ' col 61 = Delta_ModelSpread_bp
Private Const WCOL_RESERVED_1 As String = "BJ"
Private Const WCOL_RESERVED_2 As String = "BK"
Private Const WCOL_RESERVED_3 As String = "BL"
Private Const WCOL_RESERVED_4 As String = "BM"
Private Const WCOL_DV01_BBG As String = "BN"
Private Const WCOL_FLOATFAMILY_STATUS As String = "BO"








' --- OIS_Curves columns ---
Private Const CVCOL_TENOR As String = "A"   ' col 1 = Tenor
Private Const CVCOL_YEARS As String = "B"   ' col 2 = Years
Private Const CVCOL_EUR_ESTR_OIS_TICKER As String = "C"   ' col 3 = EUR_ESTR_OIS_Ticker
Private Const CVCOL_EUR_ESTR_OIS_TM1 As String = "D"   ' col 4 = EUR_ESTR_OIS_T-1
Private Const CVCOL_EUR_ESTR_OIS_T0 As String = "E"   ' col 5 = EUR_ESTR_OIS_T0
Private Const CVCOL_EUR_GOV_TICKER As String = "F"   ' col 6 = EUR_Gov_Ticker
Private Const CVCOL_EUR_GOV_TM1 As String = "G"   ' col 7 = EUR_Gov_T-1
Private Const CVCOL_EUR_GOV_T0 As String = "H"   ' col 8 = EUR_Gov_T0
Private Const CVCOL_EUR_EURIBOR_SWAP_TICKER As String = "I"   ' col 9 = EUR_EURIBOR_Swap_Ticker
Private Const CVCOL_EUR_EURIBOR_SWAP_TM1 As String = "J"   ' col 10 = EUR_EURIBOR_Swap_T-1
Private Const CVCOL_EUR_EURIBOR_SWAP_T0 As String = "K"   ' col 11 = EUR_EURIBOR_Swap_T0
Private Const CVCOL_EUR_G_TM1 As String = "L"   ' col 12 = EUR_g_TM1
Private Const CVCOL_EUR_G_T0 As String = "M"   ' col 13 = EUR_g_T0
Private Const CVCOL_EUR_Q_TM1 As String = "N"   ' col 14 = EUR_q_TM1
Private Const CVCOL_EUR_Q_T0 As String = "O"   ' col 15 = EUR_q_T0
Private Const CVCOL_EUR_STATUS As String = "P"   ' col 16 = EUR_Status
Private Const CVCOL_YEARS_S As String = "S"   ' col 19 = Years
Private Const CVCOL_USD_OIS_TICKER As String = "T"   ' col 20 = USD_OIS_Ticker
Private Const CVCOL_USD_OIS_TM1 As String = "U"   ' col 21 = USD_OIS_T-1
Private Const CVCOL_USD_OIS_T0 As String = "V"   ' col 22 = USD_OIS_T0
Private Const CVCOL_USD_GOV_TICKER As String = "W"   ' col 23 = USD_Gov_Ticker
Private Const CVCOL_USD_GOV_TM1 As String = "X"   ' col 24 = USD_Gov_T-1
Private Const CVCOL_USD_GOV_T0 As String = "Y"   ' col 25 = USD_Gov_T0
Private Const CVCOL_USD_SWAP_TICKER As String = "Z"   ' col 26 = USD_Swap_Ticker
Private Const CVCOL_USD_SWAP_TM1 As String = "AA"   ' col 27 = USD_Swap_T-1
Private Const CVCOL_USD_SWAP_T0 As String = "AB"   ' col 28 = USD_Swap_T0
Private Const CVCOL_USD_G_TM1 As String = "AC"   ' col 29 = USD_g_TM1
Private Const CVCOL_USD_G_T0 As String = "AD"   ' col 30 = USD_g_T0
Private Const CVCOL_USD_Q_TM1 As String = "AE"   ' col 31 = USD_q_TM1
Private Const CVCOL_USD_Q_T0 As String = "AF"   ' col 32 = USD_q_T0
Private Const CVCOL_USD_STATUS As String = "AG"   ' col 33 = USD_Status
Private Const CVCOL_YEARS_AJ As String = "AJ"   ' col 36 = Years
Private Const CVCOL_GBP_OIS_TICKER As String = "AK"   ' col 37 = GBP_OIS_Ticker
Private Const CVCOL_GBP_OIS_TM1 As String = "AL"   ' col 38 = GBP_OIS_T-1
Private Const CVCOL_GBP_OIS_T0 As String = "AM"   ' col 39 = GBP_OIS_T0
Private Const CVCOL_GBP_GOV_TICKER As String = "AN"   ' col 40 = GBP_Gov_Ticker
Private Const CVCOL_GBP_GOV_TM1 As String = "AO"   ' col 41 = GBP_Gov_T-1
Private Const CVCOL_GBP_GOV_T0 As String = "AP"   ' col 42 = GBP_Gov_T0
Private Const CVCOL_GBP_SWAP_TICKER As String = "AQ"   ' col 43 = GBP_Swap_Ticker
Private Const CVCOL_GBP_SWAP_TM1 As String = "AR"   ' col 44 = GBP_Swap_T-1
Private Const CVCOL_GBP_SWAP_T0 As String = "AS"   ' col 45 = GBP_Swap_T0
Private Const CVCOL_GBP_G_TM1 As String = "AT"   ' col 46 = GBP_g_TM1
Private Const CVCOL_GBP_G_T0 As String = "AU"   ' col 47 = GBP_g_T0
Private Const CVCOL_GBP_Q_TM1 As String = "AV"   ' col 48 = GBP_q_TM1
Private Const CVCOL_GBP_Q_T0 As String = "AW"   ' col 49 = GBP_q_T0
Private Const CVCOL_GBP_STATUS As String = "AX"   ' col 50 = GBP_Status


' --- SwapMap (support) columns ---
Private Const SMCOL_SOURCEROW As String = "A"   ' col 1 = SourceRow
Private Const SMCOL_COVERAGERELATION As String = "B"   ' col 2 = CoverageRelation
Private Const SMCOL_LINKEDISIN As String = "C"   ' col 3 = LinkedISIN
Private Const SMCOL_SWAP_ID_SOURCE As String = "D"   ' col 4 = Swap_ID_Source
Private Const SMCOL_SWAP_DIRECT_ID As String = "E"   ' col 5 = Swap_Direct_ID
Private Const SMCOL_FIXED_LEG_ID As String = "F"   ' col 6 = Fixed_Leg_ID
Private Const SMCOL_FLOAT_LEG_ID As String = "G"   ' col 7 = Float_Leg_ID
Private Const SMCOL_IMPORTSTATUS As String = "H"   ' col 8 = ImportStatus
Private Const SMCOL_COUNTERPARTY As String = "I"   ' col 9 = Counterparty
Private Const SMCOL_NOTIONAL As String = "J"   ' col 10 = Notional
Private Const SMCOL_STARTDATE As String = "K"   ' col 11 = StartDate
Private Const SMCOL_ENDDATE As String = "L"   ' col 12 = EndDate
Private Const SMCOL_CCY As String = "M"   ' col 13 = CCY
Private Const SMCOL_NOTIONAL_SOURCE As String = "N"   ' col 14 = Notional_Source


' --- CoverageFutures (support) columns ---
Private Const CFCOL_SOURCEROW As String = "A"   ' col 1 = SourceRow
Private Const CFCOL_COVERAGERELATION As String = "B"   ' col 2 = CoverageRelation
Private Const CFCOL_LINKEDISIN As String = "C"   ' col 3 = LinkedISIN
Private Const CFCOL_COVERAGEINFO_D As String = "D"   ' col 4 = CoverageInfo_D
Private Const CFCOL_FUTURELABEL As String = "E"   ' col 5 = FutureLabel
Private Const CFCOL_FUTURECODE As String = "F"   ' col 6 = FutureCode
Private Const CFCOL_HEDGETYPE As String = "G"   ' col 7 = HedgeType
Private Const CFCOL_CONTRACTS As String = "H"   ' col 8 = Contracts
Private Const CFCOL_STARTDATE As String = "I"   ' col 9 = StartDate
Private Const CFCOL_CCY As String = "J"   ' col 10 = CCY
Private Const CFCOL_IMPORTSTATUS As String = "K"   ' col 11 = ImportStatus
Private Const CFCOL_HEDGE_SOURCE As String = "L"   ' col 12 = Hedge_Source
Private Const CFCOL_BPV As String = "M"   ' col 13 = Coverage_BPV


' --- FutMap (support) columns ---
Private Const FMCOL_OPICS_CODE As String = "A"   ' col 1 = OPICS_Code
Private Const FMCOL_BLOOMBERG_GENERIC As String = "B"   ' col 2 = Bloomberg_Generic
Private Const FMCOL_DESCRIPTION As String = "C"   ' col 3 = Description
Private Const FMCOL_STATUS As String = "D"   ' col 4 = Status


' =============================================================================
' FIXED RANGE / LABEL / CELL CONSTANTS  (whole ranges & label cells that are not
' a single schema column: titles, curve scaffolding, multi-area & support ranges)
' =============================================================================


' The bond table's own header cell.  Was the literal "A1"; with rows 1-2 now
' reserved for comments, A1 is a comment cell and must not be touched.
Private Const CELL_BONDS_HEADER As String = BCOL_ISIN
Private Const CELL_CFG_A1 As String = "A1"
Private Const CELL_COVF_A1 As String = "A1"
Private Const CELL_COV_A1 As String = "A1"
Private Const CELL_COV_A2 As String = "A2"
Private Const CELL_COV_A8 As String = "A8"
Private Const CELL_CV_A1 As String = "A1"
Private Const CELL_CV_A4 As String = "A4"
Private Const CELL_CV_A44 As String = "A44"
Private Const CELL_CV_A46 As String = "A46"
Private Const CELL_CV_A47 As String = "A47"
Private Const CELL_CV_A48 As String = "A48"
Private Const CELL_CV_AJ4 As String = "AJ4"
Private Const CELL_CV_B4 As String = "B4"
Private Const CELL_CV_B44 As String = "B44"
Private Const CELL_CV_B46 As String = "B46"
Private Const CELL_CV_B47 As String = "B47"
Private Const CELL_CV_B48 As String = "B48"
Private Const CELL_CV_S4 As String = "S4"
Private Const CELL_DIAG_A1 As String = "A1"
' FutMap has its own geometry.  It used to borrow BOND_HEADER_ROW as a stand-in
' for "row 1", which was harmless only while the Bonds header happened to be on
' row 1 - moving the Bonds header for the comment band would have put FutMap's
' header on row 3 while its seed rows still went to row 2.
Private Const FUTMAP_HEADER_ROW As Long = 1
Private Const FUTMAP_DATA_ROW As Long = FUTMAP_HEADER_ROW + 1
Private Const FUTMAP_NAMED_LAST_ROW As Long = 200
Private Const CELL_FUT_A1 As String = "A1"
Private Const CELL_PNL_A1 As String = "A1"
Private Const CELL_SWAPMAP_A1 As String = "A1"
Private Const CELL_SWAP_A1 As String = "A1"
Private Const RNG_COV_A1_A2 As String = "A1:A2"
Private Const RNG_CV_A51_A112__D51_D112__G51_G112__R51_R112__U51_U112__Y51_Y112__AI51_AI112__AL51_AL112__AO51_AO112 As String = "A51:A112,D51:D112,G51:G112,R51:R112,U51:U112,Y51:Y112,AI51:AI112,AL51:AL112,AO51:AO112"
Private Const RNG_DIAG_A1_E1 As String = "A1:E1"
Private Const RNG_DIAG_A1_F1 As String = "A1:F1"


' =============================================================================
' BLOOMBERG FIELD CONSTANTS
' =============================================================================


Private Const BBG_PX_LAST          As String = "PX_LAST"
Private Const BBG_PX_DIRTY_MID     As String = "PX_DIRTY_MID"
Private Const BBG_YLD_YTM_MID      As String = "YLD_YTM_MID"
Private Const BBG_DUR_ADJ_MID      As String = "DUR_ADJ_MID"
Private Const BBG_CONVEXITY        As String = "CONVEXITY"
Private Const BBG_CONVEXITY_OAS    As String = "CONVEXITY_OAS"
Private Const BBG_ZSPREAD          As String = "Z_SPRD_MID"
Private Const BBG_ASW              As String = "ASSET_SWAP_SPD_MID"
Private Const BBG_OAS              As String = "OAS_SPREAD_MID"
Private Const BBG_DUR_ADJ_OAS_MID  As String = "DUR_ADJ_OAS_MID"
Private Const BBG_PRICING_SOURCE   As String = "PRICING_SOURCE"


Private Const BBG_FUT_VAL_PT       As String = "FUT_VAL_PT"
Private Const BBG_FUT_CONT_SIZE    As String = "FUT_CONT_SIZE"
Private Const BBG_PAY_FLT_RATE_IDX As String = "PAY_FLT_RATE_IDX"


' =============================================================================
' BQL FIELD / EXPRESSION CONSTANTS FOR T0 ONLY
' =============================================================================


Private Const BQL_PX_LAST      As String = "PX_LAST"
Private Const BQL_CLEAN_PX     As String = "PX_LAST"
Private Const BQL_DIRTY_PX     As String = "price(price_type='dirty')"
Private Const BQL_YTM          As String = "YIELD(YIELD_TYPE='YTM')"
Private Const BQL_ZSPREAD      As String = "SPREAD(SPREAD_TYPE='Z')"
Private Const BQL_ASW          As String = "SPREAD(SPREAD_TYPE='ASW')"
Private Const BQL_OAS          As String = "SPREAD(SPREAD_TYPE='OAS')"
Private Const BQL_SWAP_MV_T1   As String = "SW_MARKET_VAL"
Private Const BQL_SWAP_MV_T0   As String = "SW_MARKET_VAL_PRIOR"
Private Const BQL_SWAP_MV      As String = "SW_MARKET_VAL"
Private Const BQL_SWAP_MV_CHG  As String = "SW_MARKET_VAL_CHG"
Private Const BQL_SWAP_NOTIONAL As String = "notional"
Private Const BQL_SWAP_DV01    As String = "dv01"




' =============================================================================
' INTERNAL BOND DAY-COUNT CODES
' =============================================================================
Private Const DCC_ACT_ACT_ICMA As Long = 0
Private Const DCC_ACT_ACT_ISDA As Long = 1
Private Const DCC_ACT_365F As Long = 2
Private Const DCC_ACT_360 As Long = 3
Private Const DCC_30_360_BOND As Long = 4
Private Const DCC_30E_360 As Long = 5
Private Const DCC_30E_360_ISDA As Long = 6


' =============================================================================
' DIAGNOSTIC STATE
'   mLastFormulaTarget records the most recent range a formula was being written
'   to, so step error handlers can report exactly WHERE a .FormulaR1C1 assignment
'   failed (e.g. an over-nested formula -> Excel run-time error 1004) instead of a
'   hardcoded "most likely location" guess.
' =============================================================================


Private mLastFormulaTarget As String
Private mLastRangeTarget As String

' Hedge-sheet extents for the SUMIFS/COUNTIFS ranges WritePNLRow emits.
' Resolved ONCE per PNL rebuild by WritePNLAttributionRows rather than per row:
' every bond row emits the same hedge lookup ranges, and re-measuring the
' Futures and Swaps sheets for each of several hundred bonds is pure waste.
Private mHedgeLookupFutLast As Long
Private mHedgeLookupSwLast As Long


Private Function CaptureErrInfo() As String
    Dim s As String


    s = "Error " & Err.Number & ": " & Err.Description


    If Len(mLastFormulaTarget) > 0 Then
        s = s & vbCrLf & "Last formula target: " & mLastFormulaTarget
    End If


    If Len(mLastRangeTarget) > 0 Then
        s = s & vbCrLf & "Last range target: " & mLastRangeTarget
    End If


    CaptureErrInfo = s
End Function


' =============================================================================
' BLOOMBERG DAY_CNT_DES -> INTERNAL DCC CODE
'
' This mapper is based on the actual values observed from:
'   =BDP(A5 & " ISIN", "DAY_CNT_DES")
'
' Observed values include:
'   ACT/ACT
'   ACT/360
'   ISMA-30/360
'   ISDA ACT/ACT
'   ACT/ACT NON-EOM
'   ACT/360(102)
'   30/360(104)
'   30/360
'   ISMA-30/360 NONEOM
'   ISDA SWAPS:30/360
'
' Output:
'   DCC_ACT_ACT_ICMA = 0
'   DCC_ACT_ACT_ISDA = 1
'   DCC_ACT_365F     = 2
'   DCC_ACT_360      = 3
'   DCC_30_360_BOND  = 4
'   DCC_30E_360      = 5
'   DCC_30E_360_ISDA = 6
'
' Design:
'   - Handles line breaks and hidden characters.
'   - Handles Bloomberg numeric suffixes like ACT/360(102).
'   - Handles NON-EOM / NONEOM schedule flags.
'   - Returns #N/A for unknown values instead of guessing.
' =============================================================================


Public Function BloombergDayCountToDCC(ByVal dayCntDesc As Variant) As Variant
    If IsError(dayCntDesc) Then
        BloombergDayCountToDCC = CVErr(xlErrNA)
        Exit Function
    End If


    Dim s As String
    Dim compact As String


    s = CStr(dayCntDesc)


    ' Remove hidden / line-break characters that may come from Bloomberg.
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")


    ' Normalize case.
    s = UCase$(Trim$(s))


    If Len(s) = 0 Then
        BloombergDayCountToDCC = CVErr(xlErrNA)
        Exit Function
    End If


    ' Normalize separators.
    s = Replace(s, "-", "/")
    s = Replace(s, "_", "/")
    s = Replace(s, "\", "/")
    s = Replace(s, ":", " ")


    ' Keep parentheses content as text, but separate it.
    ' Example: ACT/360(102) -> ACT/360 102
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")


    ' Normalize common words.
    s = Replace(s, "ACTUAL", "ACT")
    s = Replace(s, "NON EOM", "NONEOM")
    s = Replace(s, "NON/EOM", "NONEOM")
    s = Replace(s, "NON-EOM", "NONEOM")
    s = Replace(s, "BOND BASIS", "BOND")
    s = Replace(s, "US SIA", "US")
    s = Replace(s, "NASD", "US")


    ' Collapse repeated spaces.
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop


    s = Trim$(s)


    ' Compact version for easier pattern matching.
    compact = Replace(s, " ", "")


    ' -------------------------------------------------------------------------
    ' Explicit synonym table for the real DAY_CNT_DES strings Bloomberg serves.
    '
    ' These are normalized (UPPER/TRIM, separators unified, spaces stripped) and
    ' map straight onto the 0..6 codes so common conventions never fall through
    ' to #N/A. The pattern blocks below still catch formatting variants.
    ' -------------------------------------------------------------------------
    Select Case compact
        Case "ACT/ACT", "ACT/ACTICMA"
            BloombergDayCountToDCC = DCC_ACT_ACT_ICMA
            Exit Function
        Case "ACT/ACTISDA"
            BloombergDayCountToDCC = DCC_ACT_ACT_ISDA
            Exit Function
        Case "ACT/365", "ACT/365F", "ACT/365FIXED", "ACT/365L", "NL/365"
            BloombergDayCountToDCC = DCC_ACT_365F
            Exit Function
        Case "ACT/360"
            BloombergDayCountToDCC = DCC_ACT_360
            Exit Function
        Case "30/360", "30/360BOND"
            BloombergDayCountToDCC = DCC_30_360_BOND
            Exit Function
        Case "30E/360", "ISMA/30/360"
            BloombergDayCountToDCC = DCC_30E_360
            Exit Function
        Case "30E/360ISDA"
            BloombergDayCountToDCC = DCC_30E_360_ISDA
            Exit Function
    End Select


    ' -------------------------------------------------------------------------
    ' ISDA ACT/ACT
    '
    ' Observed:
    '   ISDA ACT/ACT
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ISDA", vbTextCompare) > 0 And _
       (InStr(1, compact, "ACT/ACT", vbTextCompare) > 0 Or _
        InStr(1, compact, "ACTACT", vbTextCompare) > 0) Then


        BloombergDayCountToDCC = DCC_ACT_ACT_ISDA
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' ACT/ACT, including ACT/ACT NON-EOM
    '
    ' Observed:
    '   ACT/ACT
    '   ACT/ACT NON-EOM
    '
    ' NON-EOM / NONEOM is a schedule flag, not a separate day-count convention.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ACT/ACT", vbTextCompare) > 0 Or _
       InStr(1, compact, "ACTACT", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_ACT_ACT_ICMA
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' ACT/365 variants
    '
    ' Not observed in your sample, but included because it is common.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ACT/365", vbTextCompare) > 0 Or _
       InStr(1, compact, "ACT365", vbTextCompare) > 0 Or _
       InStr(1, compact, "A/365", vbTextCompare) > 0 Or _
       InStr(1, compact, "365F", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_ACT_365F
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' ACT/360, including ACT/360(102)
    '
    ' Observed:
    '   ACT/360
    '   ACT/360(102)
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ACT/360", vbTextCompare) > 0 Or _
       InStr(1, compact, "ACT360", vbTextCompare) > 0 Or _
       InStr(1, compact, "A/360", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_ACT_360
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' ISMA-30/360, including ISMA-30/360 NONEOM
    '
    ' Observed:
    '   ISMA-30/360
    '   ISMA-30/360 NONEOM
    '
    ' Mapped to Eurobond 30E/360.
    ' NONEOM is a schedule flag, not a separate day-count convention.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ISMA", vbTextCompare) > 0 And _
       InStr(1, compact, "30/360", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_30E_360
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' 30E/360 ISDA
    '
    ' Not observed in your sample, but included for completeness.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "30E/360", vbTextCompare) > 0 And _
       InStr(1, compact, "ISDA", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_30E_360_ISDA
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' 30E/360
    '
    ' Not observed directly in your sample, but included for completeness.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "30E/360", vbTextCompare) > 0 Then
        BloombergDayCountToDCC = DCC_30E_360
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' ISDA SWAPS:30/360
    '
    ' Observed:
    '   ISDA SWAPS:30/360
    '
    ' Your current UDF library does not have a separate DCC for "ISDA swaps 30/360".
    ' The closest supported convention is DCC_30_360_BOND.
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "ISDA", vbTextCompare) > 0 And _
       InStr(1, compact, "SWAPS", vbTextCompare) > 0 And _
       InStr(1, compact, "30/360", vbTextCompare) > 0 Then


        BloombergDayCountToDCC = DCC_30_360_BOND
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' Plain 30/360, including 30/360(104)
    '
    ' Observed:
    '   30/360
    '   30/360(104)
    ' -------------------------------------------------------------------------
    If InStr(1, compact, "30/360", vbTextCompare) > 0 Then
        BloombergDayCountToDCC = DCC_30_360_BOND
        Exit Function
    End If


    ' -------------------------------------------------------------------------
    ' Unknown convention.
    '
    ' Do not silently default. Dirty price should be blank/reviewed if the
    ' convention is unknown.
    ' -------------------------------------------------------------------------
    BloombergDayCountToDCC = CVErr(xlErrNA)
End Function


Public Function ExcelPriceBasisFromBondDCC(ByVal dccCode As Variant) As Variant


    If IsError(dccCode) Then
        ExcelPriceBasisFromBondDCC = 1
        Exit Function
    End If


    If CleanText(dccCode) = "" Then
        ExcelPriceBasisFromBondDCC = 1
        Exit Function
    End If


    If Not IsNumeric(dccCode) Then
        ExcelPriceBasisFromBondDCC = 1
        Exit Function
    End If


    Select Case CLng(dccCode)


        Case 0, 1
            ' Internal:
            '   0 = ACT/ACT ICMA
            '   1 = ACT/ACT ISDA
            '
            ' Excel PRICE basis:
            '   1 = Actual/Actual
            ExcelPriceBasisFromBondDCC = 1


        Case 2
            ' Internal:
            '   2 = ACT/365F
            '
            ' Excel PRICE basis:
            '   3 = Actual/365
            ExcelPriceBasisFromBondDCC = 3


        Case 3
            ' Internal:
            '   3 = ACT/360
            '
            ' Excel PRICE basis:
            '   2 = Actual/360
            ExcelPriceBasisFromBondDCC = 2


        Case 4
            ' Internal:
            '   4 = 30/360 BOND
            '
            ' Excel PRICE basis:
            '   0 = US/NASD 30/360
            ExcelPriceBasisFromBondDCC = 0


        Case 5, 6
            ' Internal:
            '   5 = 30E/360
            '   6 = 30E/360 ISDA
            '
            ' Excel PRICE basis:
            '   4 = European 30/360
            ExcelPriceBasisFromBondDCC = 4


        Case Else
            ' Conservative fallback:
            ' use Actual/Actual so the bump convexity formula remains alive.
            ExcelPriceBasisFromBondDCC = 1


    End Select


End Function






' =============================================================================
' BUTTON 0 - SETUP FINAL LAYOUT
' =============================================================================


Public Sub SetupWorkbookFinalLayout()


    Dim wb As Workbook
    Set wb = ThisWorkbook


    On Error GoTo CleanFail


    Application.ScreenUpdating = False
    Application.EnableEvents = False


    SetupExtendedConfig PnlGetOrCreateSheet(SH_CONFIG)
    SetupNamedRanges


    EnsureFutMapSheet
    EnsureCoverageSupportSheet
    EnsureSwapMapSheet
    EnsureCoverageFuturesSheet
    EnsureSpreadOverrideSheet


    SetupCurvesLayout PnlGetOrCreateSheet(SH_OIS)


    ' Bonds!A:L are query-owned.
    ' SetupBondsFinalHeaders only writes M:CN.
    EnsureBondsCommentRows PnlGetOrCreateSheet(SH_BONDS)
    SetupBondsFinalHeaders PnlGetOrCreateSheet(SH_BONDS)


    SetupFuturesFinalHeaders PnlGetOrCreateSheet(SH_FUTURES)
    SetupSwapsFinalHeaders PnlGetOrCreateSheet(SH_SWAPS)
    SetupPNLAttributionHeaders PnlGetOrCreateSheet(SH_PNL)
    SetupDiagnosticsSheet


CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    If Err.Number = 0 Then
        MsgBox "Final workbook layout created/updated.", vbInformation, "Setup Complete"
    End If


    Exit Sub


CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox "SetupWorkbookFinalLayout failed:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical


End Sub




' =============================================================================
' CONFIG R1C1 HELPER  (the single fix that repairs DaysToDeliv & futures pulls)
'   Returns an absolute cross-sheet R1C1 reference to a Config cell, e.g.
'   CfgR1C1(CFG_T1_DATE) -> "'Config'!R5C2"
' =============================================================================


Private Function CfgR1C1(ByVal a1Cell As String) As String
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets(SH_CONFIG).Range(a1Cell)
    CfgR1C1 = "'" & SH_CONFIG & "'!R" & rng.Row & "C" & rng.Column
End Function


' =============================================================================
' FORMULA REFERENCE HELPERS
'   Bridge the letter column-constants (BCOL_*/PCOL_*/...) into formula BODIES so
'   formula strings reference columns by name, stay readable, and auto-follow a
'   column shift (single source of truth).  RC(BCOL_GOV_T0) emits the R1C1 token
'   "RC55"; the generated Excel formulas are unchanged.
' =============================================================================
Private Function ColIdx(ByVal letters As String) As Long
    Dim i As Long, ch As Long, n As Long
    letters = UCase$(Trim$(letters))
    For i = 1 To Len(letters)
        ch = Asc(Mid$(letters, i, 1)) - 64
        If ch < 1 Or ch > 26 Then ColIdx = 0: Exit Function
        n = n * 26 + ch
    Next i
    ColIdx = n
End Function


' Absolute same-row R1C1 column token from a column-letter constant.
'   RC(BCOL_GOV_T0) -> "RC55"
Private Function RC(ByVal colConst As String) As String
    RC = "RC" & ColIdx(colConst)
End Function


' Total futures DV01 for one PNL_Attribution row.
'
' The desk's futures hedges arrive in TWO coverage files, so the risk is
' reported in two columns.  Everything that needs "the futures hedge" as a whole
' - the hedge total, the framework's DV01 mix, the model futures PnL - has to
' add both.  A bond hedged entirely out of one file would otherwise read as
' unhedged, pick the wrong spread framework, and book its whole futures PnL as
' unexplained.
'
' Both columns are IFERROR(SUMIFS(...),0), so this is always numeric and can be
' wrapped in ISNUMBER() exactly like the single column it replaces.
Private Function FutDv01(ByVal p As String) As String
    FutDv01 = "(" & PCOL_FUTURES_RTJ_DV01 & p & "+" & PCOL_FUTURES_RT_DV01 & p & ")"
End Function


' The extra SUMIFS/COUNTIFS criterion that keeps a futures aggregate to one
' hedge class.  Every futures sum on PNL_Attribution is a RATES sum: the EUR/USD
' contracts on the same sheet hedge currency, not the curve, and adding their
' notional-derived DV01 to a bond's rates hedge overstates the hedge and moves
' the spread framework onto the wrong curve.  The FX rows are reported
' separately, in the FX block on the Dashboard.
Private Function FutClassCrit( _
    ByVal fu As String, _
    ByVal futLast As String, _
    ByVal hedgeClass As String) As String

    FutClassCrit = "," & fu & "$" & FCOL_HEDGE_CLASS & "$" & DATA_ROW & _
        ":$" & FCOL_HEDGE_CLASS & "$" & futLast & ",""" & hedgeClass & """"

End Function


' Empty-row guard wrapper (the ubiquitous =IF(A5="","",...) pattern):
'   WrapIfPresent(RC(BCOL_ISIN), body) -> "=IF(RC1="""",""""," & body & ")"
Private Function WrapIfPresent(ByVal guardRef As String, ByVal bodyExpr As String) As String
    WrapIfPresent = "=IF(" & guardRef & "="""",""""," & bodyExpr & ")"
End Function


' Tokens whose presence in a futures row's label marks it an FX hedge rather
' than a rates hedge.  Deliberately about the CURRENCY PAIR and the instrument
' word, not about exchange tickers: a ticker list goes stale the first time the
' desk rolls into a contract nobody wrote down, and a wrong entry there would
' silently reclassify a real bond future.
Private Function FxHedgeLabelTokens() As Variant
    FxHedgeLabelTokens = Array( _
        "EURUSD", "EUR/USD", "EUR USD", "USDEUR", "USD/EUR", _
        "FX", "FWD", "FORWARD", "NDF")
End Function


' R1C1 expression for Futures!Hedge_Class.  Searches the contract code, the
' coverage label and the hedge type together, so a pair named in any one of
' them is caught.
Private Function FxHedgeClassExpr() As String

    Dim tokens As Variant
    Dim i As Long
    Dim s As String

    tokens = FxHedgeLabelTokens()

    s = "LET(_lbl,UPPER(TRIM(" & RC(FCOL_CONTRACTCODE) & "&"" ""&" & _
        RC(FCOL_FUTURELABEL) & "&"" ""&" & RC(FCOL_HEDGETYPE) & "))," & _
        "_extra,UPPER(TRIM(" & CfgR1C1(CFG_FX_HEDGE_TOKENS) & "))," & _
        "_hit,"

    For i = LBound(tokens) To UBound(tokens)
        If i > LBound(tokens) Then s = s & "+"
        s = s & "--ISNUMBER(SEARCH(""" & CStr(tokens(i)) & """,_lbl))"
    Next i

    ' The Config list is a semicolon-separated set, so it has to be split
    ' before searching; TEXTSPLIT of an empty string errors, hence the guard.
    s = s & "+IF(_extra="""",0,SUMPRODUCT(--ISNUMBER(SEARCH(TEXTSPLIT(_extra,"";""),_lbl))))," & _
        "IF(_hit>0,""" & HEDGE_CLASS_FX & """,""" & HEDGE_CLASS_RATES & """))"

    FxHedgeClassExpr = s

End Function




' =============================================================================
' SPREAD FRAMEWORK LIBRARY
'
' A "framework chain" is the ordered set of PNL_Attribution columns whose sum
' reproduces -DV01 * Delta_y for the selected framework.  Each chain is an exact
' telescoping identity in the underlying rates, which is why the split is safe:
'
'   G     Delta_r + Delta(Gov-OIS) + Delta(YTM-Gov)                    == Delta_y
'   I     Delta_r + Delta(Gov-OIS) + Delta(Swap-Gov) + Delta(YTM-Swap) == Delta_y
'   ASW   Delta_r + Delta(Gov-OIS) + Delta(Swap-Gov) + Delta(ASW)      ~= Delta_y
'   Z     Delta_r + Delta(Gov-OIS) + Delta(Swap-Gov) + Delta(Z)        ~= Delta_y
'   OAS   Delta_r + Delta(Gov-OIS) + Delta(Swap-Gov) + Delta(OAS)      ~= Delta_y
'
' The framework therefore changes only HOW the same total is split, never the
' total itself.  Duration_Identity_Check measures the residual of that identity,
' so a non-zero value means an input is stale, not that the framework is wrong.
'
' ASW / Z / OAS are quoted on their own conventions rather than being derived
' from YTM, so their chains tie approximately rather than exactly; the identity
' check is the place that difference shows up.
' =============================================================================


' Guarded sum of a framework chain:
'   ChainSumFml("5", Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_GSPREAD))
'   -> IF(AND(ISNUMBER(AB5),ISNUMBER(AC5),ISNUMBER(AL5)),AB5+AC5+AL5,"")
Private Function ChainSumFml(ByVal p As String, ByVal cols As Variant) As String
    Dim i As Long
    Dim guard As String
    Dim body As String


    For i = LBound(cols) To UBound(cols)
        If i > LBound(cols) Then
            guard = guard & ","
            body = body & "+"
        End If
        guard = guard & "ISNUMBER(" & cols(i) & p & ")"
        body = body & cols(i) & p
    Next i


    ChainSumFml = "IF(AND(" & guard & ")," & body & ","""")"
End Function


' Whole yield move, no decomposition:  -DV01 * Delta_y_bp.  Used by the OIS and
' SOFR frameworks, and it is also the benchmark the identity check compares to.
Private Function YieldOnlyPnLFml(ByVal p As String) As String
    YieldOnlyPnLFml = _
        "IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Y_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_Y_BP & p & ","""")"
End Function


' Futures/swap DV01-weighted blend of two expressions, for the MIXED framework.
' Weights are ABS DV01 shares, so a bond hedged half by futures and half by
' swaps is attributed half on the govie chain and half on the swap chain.
Private Function DV01BlendFml(ByVal p As String, ByVal futExpr As String, ByVal swpExpr As String) As String
    DV01BlendFml = _
        "LET(_f,ABS" & FutDv01(p) & "," & _
        "_s,ABS(" & PCOL_PLAINSWAP_DV01 & p & ")," & _
        "_t,_f+_s," & _
        "IF(_t=0,""""," & _
        "IFERROR(_f/_t*(" & futExpr & ")+_s/_t*(" & swpExpr & "),"""")))"
End Function


' Availability chain: the first framework code in `order` whose PnL column is
' numeric on this row, "" if none of them is.  Returning "" instead of a hard
' default is what lets the automatic rule abstain rather than assert a framework
' it has no evidence for.
Private Function FrameworkPrefFml(ByVal order As Variant) As String

    Dim i As Long
    Dim s As String
    Dim closers As String

    For i = LBound(order) To UBound(order)
        s = s & "IF(_has" & CStr(order(i)) & ",""" & CStr(order(i)) & ""","
        closers = closers & ")"
    Next i

    FrameworkPrefFml = s & """""" & closers

End Function


' Spread_Framework_Auto.  See the block comment at the CA/CB writer in
' WritePNLRow for the economics; this builds the string.
'
' Assembled by concatenation rather than one continued literal because the
' finished formula is well past VBA's 25-continuation limit for a single
' logical line.
Private Function SpreadFrameworkAutoFml(ByVal p As String, ByVal c As String) As String

    Dim s As String
    Dim prefFut As String
    Dim prefSwp As String
    Dim prefNone As String

    ' Futures hedge government risk, so a govie-relative spread is preferred.
    prefFut = FrameworkPrefFml(Array("G", "I", "ASW", "Z", "OAS"))
    ' Swaps hedge swap risk, so a swap-relative spread is preferred.
    prefSwp = FrameworkPrefFml(Array("I", "ASW", "G", "Z", "OAS"))
    ' Unhedged: no hedge to follow, fall back on what the data supports.
    prefNone = FrameworkPrefFml(Array("I", "G", "ASW", "Z", "OAS"))

    s = "=IF(" & PCOL_ISIN & p & "="""","""","
    s = s & "LET("
    s = s & "_ovr,IFERROR(UPPER(TRIM(VLOOKUP(" & PCOL_ISIN & p & ",SpreadOverrideTable,2,FALSE))),""""),"
    s = s & "_gbl,IFERROR(UPPER(TRIM(" & c & CFG_SPREAD_FRAMEWORK & ")),""""),"
    s = s & "_fut,ABS" & FutDv01(p) & ","
    s = s & "_swp,ABS(" & PCOL_PLAINSWAP_DV01 & p & "),"
    s = s & "_tot,_fut+_swp,"
    s = s & "_hasI,ISNUMBER(" & PCOL_PNL_CREDIT_ISPREAD & p & "),"
    s = s & "_hasG,ISNUMBER(" & PCOL_PNL_GSPREAD & p & "),"
    s = s & "_hasASW,ISNUMBER(" & PCOL_PNL_ASW & p & "),"
    s = s & "_hasZ,ISNUMBER(" & PCOL_PNL_ZSPREAD & p & "),"
    s = s & "_hasOAS,ISNUMBER(" & PCOL_PNL_OAS & p & "),"

    ' Neither leg small enough to ignore, and both chains priceable -> MIXED.
    s = s & "_both,AND(_fut>0,_swp>0,_hasG,_hasI,"
    s = s & "MIN(_fut,_swp)>=" & MIXED_MIN_SHARE & "*_tot),"

    s = s & "_auto,IF(_both,""MIXED"","
    s = s & "IF(AND(_fut>0,_fut>=_swp)," & prefFut & ","
    s = s & "IF(_swp>0," & prefSwp & ","
    s = s & prefNone & "))),"

    ' Per-bond override first, then the hedge-driven rule, and only then the
    ' global cell - and REVIEW if even that is blank.
    s = s & "_pick,IF(_ovr<>"""",_ovr,IF(_auto<>"""",_auto,IF(_gbl<>"""",_gbl,""REVIEW""))),"
    s = s & "IF(ISNA(MATCH(_pick," & FRAMEWORK_CODES_ARRAY & ",0)),""REVIEW"",_pick)))"

    SpreadFrameworkAutoFml = s

End Function


' Spread_Framework_Reason.  Says both WHAT was picked and WHICH of the three
' precedence steps picked it, so a surprising framework on the Dashboard can be
' traced without opening the formula.
Private Function SpreadFrameworkReasonFml(ByVal p As String, ByVal c As String) As String

    Dim s As String

    s = "=IF(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & "="""","""","
    s = s & "LET("
    s = s & "_ovr,IFERROR(UPPER(TRIM(VLOOKUP(" & PCOL_ISIN & p & ",SpreadOverrideTable,2,FALSE))),""""),"
    s = s & "_gbl,IFERROR(UPPER(TRIM(" & c & CFG_SPREAD_FRAMEWORK & ")),""""),"
    s = s & "_fut,ABS" & FutDv01(p) & ","
    s = s & "_swp,ABS(" & PCOL_PLAINSWAP_DV01 & p & "),"

    ' The automatic rule abstains exactly when no spread leg is numeric, so
    ' _anyLeg reproduces "did step 2 resolve?" without restating the whole
    ' _auto expression here.  MIXED needs both G and I, which implies _anyLeg.
    s = s & "_anyLeg,OR(ISNUMBER(" & PCOL_PNL_CREDIT_ISPREAD & p & "),ISNUMBER(" & PCOL_PNL_GSPREAD & p & "),"
    s = s & "ISNUMBER(" & PCOL_PNL_ASW & p & "),ISNUMBER(" & PCOL_PNL_ZSPREAD & p & "),ISNUMBER(" & PCOL_PNL_OAS & p & ")),"

    s = s & "_src,IF(_ovr<>"""",""per-bond override (SpreadOverride sheet)"","
    s = s & "IF(_anyLeg,IF(OR(_fut>0,_swp>0),""automatic from hedge DV01 mix"","
    s = s & """automatic from available spread legs (no hedge attached)""),"
    s = s & "IF(_gbl<>"""",""Config!" & CFG_SPREAD_FRAMEWORK & " fallback - automatic rule had no spread leg to pick"","
    s = s & """no framework resolvable - no spread leg and no fallback""))),"
    s = s & "_why,SWITCH(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & ","
    s = s & """G"",""measured vs governments - futures/govie hedge or G-spread strongest available"","
    s = s & """I"",""measured vs swaps - swap hedge or default credit framework"","
    s = s & """ASW"",""asset-swap spread"","
    s = s & """Z"",""Z-spread fallback"","
    s = s & """OAS"",""option-adjusted spread"","
    s = s & """OIS"",""no decomposition - whole yield move taken as -DV01 x Delta_y"","
    s = s & """SOFR"",""no decomposition - whole yield move taken as -DV01 x Delta_y"","
    s = s & """MIXED"",""hedge split across futures and swaps - DV01-weighted blend of the govie and swap chains"","
    s = s & """REVIEW"",""attribution suppressed - unrecognised override code, or no usable spread leg on this row"","
    s = s & """framework fallback""),"
    s = s & "_why&"" ["" &_src& ""]""))"

    SpreadFrameworkReasonFml = s

End Function


' The columns that must be numeric for a PNL_Attribution row to be trusted,
' paired with the name the Dashboard shows when one of them is not.
'
' Deliberately the LEGS OF THE BRIDGE plus the two ends, not every column on
' the sheet: a blank Delta_ASW_bp on a G-framework bond is not a fault, it is
' a leg that framework does not use.  Requiring literally every column would
' quarantine the whole book on the first framework that ignores a spread.
Private Function RowValidRequiredCols() As Variant

    Dim spec As Variant

    spec = Array( _
        Array(PCOL_BOND_DV01_CURRENT, "bond DV01"), _
        Array(PCOL_BOND_DV01_OPENING, "opening bond DV01"), _
        Array(PCOL_HEDGE_DV01, "hedge DV01"), _
        Array(PCOL_DELTA_DIRTY_MV_EUR, "dirty MV change"), _
        Array(PCOL_DELTA_Y_BP, "yield change"), _
        Array(PCOL_PNL_DURATION_TOTAL, "duration PnL"), _
        Array(PCOL_PNL_CONVEXITY, "convexity PnL"), _
        Array(PCOL_CARRY_TOTAL, "carry"), _
        Array(PCOL_SPREADPNL_USED, "spread PnL"), _
        Array(PCOL_PNL_FX, "FX PnL"), _
        Array(PCOL_TOTAL_EXPLAINED, "explained total"), _
        Array(PCOL_OFFICIAL_PNL, "official PnL"), _
        Array(PCOL_RESIDUAL, "residual"))

    RowValidRequiredCols = spec

End Function


' Row_Exclusion_Reason: "" when the row is sound, otherwise the FIRST reason it
' is not, ordered cause before symptom so the reader is pointed at the input
' that failed rather than at everything downstream of it.
Private Function RowExclusionReasonFml( _
    ByVal p As String, _
    ByVal B As String, _
    ByVal b2 As String) As String

    Dim s As String
    Dim closers As String
    Dim spec As Variant
    Dim i As Long

    s = "=IF(" & PCOL_ISIN & p & "="""","""","
    closers = ")"

    ' 1. The bond''s own market data, which everything else is built on.
    s = s & "IF(" & B & BCOL_BOND_STATUS & b2 & "<>""OK"",""Bond data: ""&" & B & BCOL_BOND_STATUS & b2 & ","
    closers = closers & ")"

    ' 2. No framework means no credit leg and no chain to sum.
    s = s & "IF(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & "=""REVIEW"",""Spread framework unresolved"","
    closers = closers & ")"

    ' 3-4. A hedge that matched but whose actual PnL never arrived.  The bond
    '      would otherwise be explained against a hedge PnL of nothing.
    s = s & "IF(AND(" & PCOL_FUT_MATCH_COUNT & p & ">0,NOT(ISNUMBER(" & PCOL_FUT_PNL_RAW & p & ")))," & _
        """Futures matched but actual futures PnL missing"","
    closers = closers & ")"

    s = s & "IF(AND(" & PCOL_SWAP_PLAIN_MATCH_COUNT & p & ">0,NOT(ISNUMBER(" & PCOL_SWAP_PLAIN_PNL_RAW & p & ")))," & _
        """Swaps matched but actual swap PnL missing"","
    closers = closers & ")"

    ' 5. Every leg of the bridge, and both of its ends.
    spec = RowValidRequiredCols()
    For i = LBound(spec) To UBound(spec)
        s = s & "IF(NOT(ISNUMBER(" & CStr(spec(i)(0)) & p & ")),""Missing " & CStr(spec(i)(1)) & ""","
        closers = closers & ")"
    Next i

    ' 6. The chain does not reproduce -DV01 x Delta_y, so the split is not
    '    trustworthy even though every leg is present.
    s = s & "IF(AND(ISNUMBER(" & PCOL_DURATION_IDENTITY_CHECK & p & "),ABS(" & PCOL_DURATION_IDENTITY_CHECK & p & ")>" & _
        "MAX(" & IDENTITY_TOL_EUR & "," & IDENTITY_TOL_PCT & "*ABS(" & PCOL_PNL_DURATION_TOTAL & p & ")))," & _
        """Duration chain does not tie"","
    closers = closers & ")"

    RowExclusionReasonFml = s & """""" & closers

End Function


' =============================================================================
' ECONOMIC FORMULA LIBRARY
'   Named, reusable builders for the economically-meaningful derived formulas
'   (in the spirit of the InterpOIS curve object).  Each takes already-formatted
'   cell REFERENCES - A1 like "$B5" or R1C1 like RC(BCOL_OIS_T0) - and returns the
'   Excel formula string.  The economic definition lives in ONE place; call it and
'   pass the parameters wherever that quantity is needed (any sheet/ref style).
' =============================================================================

Public Function BondSpreadTMinus1( _
    ByVal framework As Variant, _
    ByVal gSpreadTM1 As Variant, _
    ByVal iSpreadTM1 As Variant, _
    ByVal aswTM1 As Variant, _
    ByVal zSpreadTM1 As Variant, _
    ByVal oasTM1 As Variant) As Variant

    Dim fw As String

    fw = UCase$(Trim$(CStr(framework)))

    Select Case fw

        Case "G"
            BondSpreadTMinus1 = NumericOrNA(gSpreadTM1)

        ' MIXED is a hedge-attribution split, not a discounting choice.  The
        ' bond has ONE price, so GOV+G and SWAP+I reprice it identically at
        ' T-1; the forward price differs only through the shape difference
        ' between the two curves over the one-day horizon, which is far below
        ' the precision of the spread inputs.  Pricing MIXED on SWAP+I keeps
        ' it to one UDF call per row instead of two.  Without this arm MIXED
        ' returned #N/A, which blanked Carry_RollToPar, then Carry_Total, then
        ' the whole bond - the reason the automatic rule could not be allowed
        ' to emit MIXED before now.
        Case "I", "MIXED"
            BondSpreadTMinus1 = NumericOrNA(iSpreadTM1)

        Case "ASW"
            BondSpreadTMinus1 = NumericOrNA(aswTM1)

        Case "Z"
            BondSpreadTMinus1 = NumericOrNA(zSpreadTM1)

        Case "OAS"
            BondSpreadTMinus1 = NumericOrNA(oasTM1)

        Case Else
            BondSpreadTMinus1 = CVErr(xlErrNA)

    End Select

End Function

Private Function NumericOrNA(ByVal v As Variant) As Variant

    If IsError(v) Then
        NumericOrNA = CVErr(xlErrNA)
    ElseIf IsNumeric(v) Then
        NumericOrNA = CDbl(v)
    Else
        NumericOrNA = CVErr(xlErrNA)
    End If

End Function

' =============================================================================
' PULL TO PAR  (a.k.a. RollToPar)
'
' WHAT IT MEASURES
'   The part of a bond's CLEAN price change between the prior date (T-1) and the
'   current date (T0) that is pure passage of time: the curve and the credit
'   spread are both held at their T-1 levels, and the bond is simply repriced
'   one period further along.
'
' METHOD - arbitrage-free forward pricing (NOT "same curve, shorter maturity")
'   Repricing off the *spot* curve at a shorter maturity would book the entire
'   roll-down of an upward-sloping curve as free PnL.  The no-arbitrage answer
'   discounts each remaining cash flow at the FORWARD rate implied by the T-1
'   curve between the T0 horizon and that cash flow:
'
'       (1 + f(h,t))^(t-h)  =  (1 + y(t))^t / (1 + y(h))^h
'
'   so the forward value at the horizon of a cash flow C_t is
'
'       C_t / (1 + f(h,t))^(t-h)  =  C_t * (1 + y(h))^h / (1 + y(t))^t
'
'   Summing over the cash flows that are still outstanding at the horizon gives
'
'       DirtyPx_fwd = (1 + y(h))^h * SUM_{t>h} C_t / (1 + y(t))^t
'
'   i.e. the classic "spot price less the PV of interim coupons, compounded to
'   the horizon".  Coupons paid on or before the horizon drop out here on
'   purpose: cash received is already reported in Carry_Coupon, and counting it
'   twice is the single easiest way to double the carry leg.
'
'   y(t) = zero rate from the T-1 curve chosen by the spread framework, plus the
'   T-1 spread.  Applying the spread flat across tenors is the same calibration
'   the rest of the sheet uses (see BondSpreadTMinus1).
'
' RESULT
'   Clean price change per 100 nominal:
'       (DirtyPx_fwd - AI(T0)) - (DirtyPx_spot - AI(T-1))
'   Accrued interest is stripped at BOTH ends so coupon accrual never leaks in.
'   The caller multiplies by Notional * FX_T-1 / 100.
'
' MODEL-vs-MARKET NOTE
'   DirtyPx_spot is a MODEL price, so it will not tie exactly to Bonds!AC.  This
'   function returns a DIFFERENCE of two model prices, so the calibration error
'   is common to both terms and cancels to first order.  Do not use either leg
'   on its own as a price.
'
' DEPENDENCIES
'   modImpRepo: PrevCouponDate, NextCouponDate, AccruedInterest
'   modPNL:     BondPullCurveNodes / BondPullInterp (curve access)
' =============================================================================

Public Function BondPullToParPrice( _
    ByVal priorDate As Variant, _
    ByVal currentDate As Variant, _
    ByVal maturity As Variant, _
    ByVal couponPct As Variant, _
    ByVal frequency As Variant, _
    ByVal dccCode As Variant, _
    ByVal ccy As Variant, _
    ByVal framework As Variant, _
    ByVal spreadTM1 As Variant) As Variant

    ' Sheet-facing wrapper: validate the row's inputs, fetch the T-1 curve once,
    ' then hand everything to the pure numeric core.  The split is deliberate -
    ' BondPullToParCore touches no worksheet and is exercised directly by
    ' tests/test_pull_to_par.bas.

    On Error GoTo Fail

    Dim priorSettlement As Date
    Dim currentSettlement As Date
    Dim maturityDate As Date

    Dim couponRate As Double
    Dim couponFrequency As Long
    Dim bondDcc As Long
    Dim spreadDecimal As Double

    Dim curveYears() As Double
    Dim curveRates() As Double
    Dim nodeCount As Long

    If Not IsDate(priorDate) Then GoTo Fail
    If Not IsDate(currentDate) Then GoTo Fail
    If Not IsDate(maturity) Then GoTo Fail

    If Not IsNumeric(couponPct) Then GoTo Fail
    If Not IsNumeric(frequency) Then GoTo Fail
    If Not IsNumeric(dccCode) Then GoTo Fail
    If Not IsNumeric(spreadTM1) Then GoTo Fail

    priorSettlement = DateValue(CDate(priorDate))
    currentSettlement = DateValue(CDate(currentDate))
    maturityDate = DateValue(CDate(maturity))

    couponRate = CDbl(couponPct)
    If Abs(couponRate) > 1 Then couponRate = couponRate / 100

    couponFrequency = CLng(frequency)
    bondDcc = CLng(dccCode)

    ' Every spread on this sheet is quoted in basis points.
    spreadDecimal = CDbl(spreadTM1) / 10000#

    nodeCount = BondPullCurveNodes( _
        CStr(ccy), CStr(framework), False, curveYears, curveRates)

    If nodeCount < 2 Then GoTo Fail

    BondPullToParPrice = BondPullToParCore( _
        priorSettlement, currentSettlement, maturityDate, _
        couponRate, couponFrequency, bondDcc, spreadDecimal, _
        curveYears, curveRates, nodeCount)

    Exit Function

Fail:
    BondPullToParPrice = CVErr(xlErrNA)

End Function


' -----------------------------------------------------------------------------
' Pure numeric core of the pull-to-par calculation.
'
' No worksheet access: the curve arrives as sorted (tenor, rate-in-percent)
' arrays and everything else as plain scalars.  Returns the CLEAN price change
' per 100 nominal, or Error 2042 (#N/A) when the inputs cannot support a price.
' -----------------------------------------------------------------------------
Public Function BondPullToParCore( _
    ByVal priorSettlement As Date, _
    ByVal currentSettlement As Date, _
    ByVal maturityDate As Date, _
    ByVal couponRate As Double, _
    ByVal couponFrequency As Long, _
    ByVal bondDcc As Long, _
    ByVal spreadDecimal As Double, _
    ByRef curveYears() As Double, _
    ByRef curveRates() As Double, _
    ByVal nodeCount As Long) As Variant

    On Error GoTo Fail

    Dim horizonYears As Double
    Dim spotToHorizon As Double
    Dim horizonGrowth As Double

    Dim spotDirty As Double
    Dim forwardDirty As Double
    Dim remainingPv As Double

    Dim accruedPrior As Double
    Dim accruedCurrent As Double

    If currentSettlement <= priorSettlement Then GoTo Fail
    If priorSettlement >= maturityDate Then GoTo Fail
    If currentSettlement >= maturityDate Then GoTo Fail

    If couponFrequency <> 1 And _
       couponFrequency <> 2 And _
       couponFrequency <> 4 Then
        GoTo Fail
    End If

    If bondDcc < 0 Or bondDcc > 6 Then GoTo Fail
    If nodeCount < 2 Then GoTo Fail

    horizonYears = BondPullYears(priorSettlement, currentSettlement)
    If horizonYears <= 0 Then GoTo Fail

    spotToHorizon = _
        BondPullInterp(curveYears, curveRates, nodeCount, horizonYears) / 100# _
        + spreadDecimal

    If 1 + spotToHorizon <= 0 Then GoTo Fail
    horizonGrowth = (1 + spotToHorizon) ^ horizonYears

    ' --- present value on the T-1 curve, twice ------------------------------
    ' spotDirty   = every remaining cash flow
    ' remainingPv = only the cash flows still outstanding at the T0 horizon;
    '               anything paid on or before it is cash in the bank, reported
    '               by Carry_Coupon, and must not be compounded into the price.
    spotDirty = BondPullPvOnCurve( _
        priorSettlement, maturityDate, couponRate, couponFrequency, _
        spreadDecimal, curveYears, curveRates, nodeCount, 0#)

    remainingPv = BondPullPvOnCurve( _
        priorSettlement, maturityDate, couponRate, couponFrequency, _
        spreadDecimal, curveYears, curveRates, nodeCount, horizonYears)

    If spotDirty <= 0 Then GoTo Fail

    forwardDirty = remainingPv * horizonGrowth

    ' --- strip accrued at both ends -> clean price change -------------------
    accruedPrior = AccruedInterest( _
        priorSettlement, maturityDate, couponRate, _
        CInt(couponFrequency), CInt(bondDcc))

    accruedCurrent = AccruedInterest( _
        currentSettlement, maturityDate, couponRate, _
        CInt(couponFrequency), CInt(bondDcc))

    BondPullToParCore = _
        (forwardDirty - accruedCurrent) - (spotDirty - accruedPrior)

    Exit Function

Fail:
    BondPullToParCore = CVErr(xlErrNA)

End Function


' -----------------------------------------------------------------------------
' Present value of a bond's remaining cash flows on a zero curve.
'
' `minYears` filters the schedule: 0 gives the full dirty price, a horizon in
' years gives only what is still outstanding at that horizon.  One routine, two
' callers, so the discounting convention cannot drift between the spot leg and
' the forward leg of the pull-to-par calculation.
'
' Rates arrive in PERCENT (as they sit on OIS_Curves); `spreadDecimal` is added
' in decimal.  Returns a dirty price per 100 nominal, or 0 when the schedule
' cannot be built.
' -----------------------------------------------------------------------------
Private Function BondPullPvOnCurve( _
    ByVal settlementDate As Date, _
    ByVal maturityDate As Date, _
    ByVal couponRate As Double, _
    ByVal couponFrequency As Long, _
    ByVal spreadDecimal As Double, _
    ByRef curveYears() As Double, _
    ByRef curveRates() As Double, _
    ByVal nodeCount As Long, _
    ByVal minYears As Double) As Double

    Dim cfDates() As Date
    Dim cfCount As Long
    Dim cashFlow As Double
    Dim cashFlowYears As Double
    Dim zeroRate As Double
    Dim pv As Double
    Dim i As Long

    On Error GoTo Fail

    cfCount = BondPullSchedule( _
        settlementDate, maturityDate, couponFrequency, cfDates)

    If cfCount < 1 Then GoTo Fail

    pv = 0#

    For i = 1 To cfCount

        cashFlow = 100# * couponRate / couponFrequency
        If i = cfCount Then cashFlow = cashFlow + 100#

        cashFlowYears = BondPullYears(settlementDate, cfDates(i))

        If cashFlowYears > 0 And cashFlowYears > minYears Then

            zeroRate = BondPullInterp( _
                curveYears, curveRates, nodeCount, cashFlowYears) / 100# _
                + spreadDecimal

            If 1 + zeroRate <= 0 Then GoTo Fail

            pv = pv + cashFlow / (1 + zeroRate) ^ cashFlowYears

        End If

    Next i

    BondPullPvOnCurve = pv
    Exit Function

Fail:
    BondPullPvOnCurve = 0#

End Function


' -----------------------------------------------------------------------------
' Model dirty price of a bond on a zero curve, per 100 nominal.
'
' Exposed because it is the natural way to check the pull-to-par result: the
' forward dirty price must equal this price grown at the horizon zero rate
' whenever no coupon falls inside the horizon.  Also handy as a sheet-side
' sanity check against the Bloomberg dirty price.
' -----------------------------------------------------------------------------
Public Function BondPullModelDirtyPrice( _
    ByVal settlementDate As Date, _
    ByVal maturityDate As Date, _
    ByVal couponRate As Double, _
    ByVal couponFrequency As Long, _
    ByVal spreadDecimal As Double, _
    ByRef curveYears() As Double, _
    ByRef curveRates() As Double, _
    ByVal nodeCount As Long) As Variant

    Dim pv As Double

    If settlementDate >= maturityDate Then
        BondPullModelDirtyPrice = CVErr(xlErrNA)
        Exit Function
    End If

    If nodeCount < 2 Then
        BondPullModelDirtyPrice = CVErr(xlErrNA)
        Exit Function
    End If

    pv = BondPullPvOnCurve( _
        settlementDate, maturityDate, couponRate, couponFrequency, _
        spreadDecimal, curveYears, curveRates, nodeCount, 0#)

    If pv <= 0 Then
        BondPullModelDirtyPrice = CVErr(xlErrNA)
    Else
        BondPullModelDirtyPrice = pv
    End If

End Function


' -----------------------------------------------------------------------------
' Time axis for the pull-to-par discounting.
'
' ACT/365F on purpose, NOT the bond's own day count: the curve nodes on
' OIS_Curves are quoted in decimal years against annually compounded zero rates,
' so the horizon and every cash flow must be measured on the SAME simple axis for
' the forward-rate identity to hold.  The bond day count is still used where it
' belongs - accrued interest.
' -----------------------------------------------------------------------------
Private Function BondPullYears(ByVal fromDate As Date, ByVal toDate As Date) As Double
    BondPullYears = CDbl(CLng(toDate) - CLng(fromDate)) / 365#
End Function


' -----------------------------------------------------------------------------
' Remaining coupon dates strictly after settlement, through maturity inclusive.
' The final entry is forced onto the maturity date so month-end roll drift can
' never leave the redemption a day or two adrift.
' -----------------------------------------------------------------------------
Private Function BondPullSchedule( _
    ByVal settlementDate As Date, _
    ByVal maturityDate As Date, _
    ByVal couponFrequency As Long, _
    ByRef cfDates() As Date) As Long

    Dim monthsPerPeriod As Long
    Dim d As Date
    Dim n As Long
    Dim i As Long

    On Error GoTo Fail

    If couponFrequency < 1 Then GoTo Fail
    monthsPerPeriod = 12 \ couponFrequency
    If monthsPerPeriod < 1 Then GoTo Fail

    d = NextCouponDate(settlementDate, maturityDate, CInt(couponFrequency))

    n = 0
    Do While d <= maturityDate
        n = n + 1
        If n > MAX_BOND_CASHFLOWS Then GoTo Fail
        d = DateAdd("m", monthsPerPeriod, d)
    Loop

    If n = 0 Then
        ' Settlement sits inside the final coupon period: redemption only.
        ReDim cfDates(1 To 1)
        cfDates(1) = maturityDate
        BondPullSchedule = 1
        Exit Function
    End If

    ReDim cfDates(1 To n)
    d = NextCouponDate(settlementDate, maturityDate, CInt(couponFrequency))
    For i = 1 To n
        cfDates(i) = d
        d = DateAdd("m", monthsPerPeriod, d)
    Next i

    cfDates(n) = maturityDate

    BondPullSchedule = n
    Exit Function

Fail:
    BondPullSchedule = 0

End Function


' -----------------------------------------------------------------------------
' Load one curve into plain arrays ONCE per bond.
'
' BondPullToParPrice touches the curve for every cash flow (up to ~120 on a long
' bond).  Going through InterpCurveValue each time would re-read the same 13-row
' worksheet range 120 times per bond and 70k+ times per recalculation.  Reading
' it once into arrays and interpolating in memory is the whole difference
' between a recalc that finishes and one that does not.
'
' Returns the node count, or 0 when the currency/framework has no curve.
' -----------------------------------------------------------------------------
Private Function BondPullCurveNodes( _
    ByVal ccy As String, _
    ByVal framework As String, _
    ByVal useCurrentCurve As Boolean, _
    ByRef curveYears() As Double, _
    ByRef curveRates() As Double) As Long

    Dim yRange As Range
    Dim rRange As Range
    Dim ys As Variant
    Dim rs As Variant
    Dim curveType As String
    Dim n As Long
    Dim i As Long

    On Error GoTo Fail

    curveType = BondPullCurveType(framework)
    If Len(curveType) = 0 Then GoTo Fail

    If Not CurveRanges(ccy, curveType, useCurrentCurve, yRange, rRange) Then
        GoTo Fail
    End If

    ys = yRange.value
    rs = rRange.value

    ReDim curveYears(1 To UBound(ys, 1))
    ReDim curveRates(1 To UBound(rs, 1))

    n = 0
    For i = 1 To UBound(ys, 1)
        If Not IsError(ys(i, 1)) And Not IsError(rs(i, 1)) Then
            If IsNumeric(ys(i, 1)) And IsNumeric(rs(i, 1)) Then
                If Len(CleanText(ys(i, 1))) > 0 And _
                   Len(CleanText(rs(i, 1))) > 0 Then
                    n = n + 1
                    curveYears(n) = CDbl(ys(i, 1))
                    curveRates(n) = CDbl(rs(i, 1))
                End If
            End If
        End If
    Next i

    If n < 2 Then GoTo Fail

    ' insertion sort by tenor (a handful of nodes)
    Dim j As Long, tx As Double, ty As Double
    For i = 2 To n
        tx = curveYears(i): ty = curveRates(i): j = i - 1
        Do While j >= 1
            If curveYears(j) <= tx Then Exit Do
            curveYears(j + 1) = curveYears(j)
            curveRates(j + 1) = curveRates(j)
            j = j - 1
        Loop
        curveYears(j + 1) = tx
        curveRates(j + 1) = ty
    Next i

    BondPullCurveNodes = n
    Exit Function

Fail:
    BondPullCurveNodes = 0

End Function


' In-memory twin of LinInterp: linear in tenor, flat extrapolation at both ends.
Private Function BondPullInterp( _
    ByRef curveYears() As Double, _
    ByRef curveRates() As Double, _
    ByVal nodeCount As Long, _
    ByVal x As Double) As Double

    Dim i As Long

    If nodeCount < 1 Then
        BondPullInterp = 0#
        Exit Function
    End If

    If x <= curveYears(1) Then
        BondPullInterp = curveRates(1)
        Exit Function
    End If

    If x >= curveYears(nodeCount) Then
        BondPullInterp = curveRates(nodeCount)
        Exit Function
    End If

    For i = 1 To nodeCount - 1
        If x >= curveYears(i) And x <= curveYears(i + 1) Then
            If curveYears(i + 1) = curveYears(i) Then
                BondPullInterp = curveRates(i)
            Else
                BondPullInterp = curveRates(i) + _
                    (curveRates(i + 1) - curveRates(i)) * _
                    (x - curveYears(i)) / (curveYears(i + 1) - curveYears(i))
            End If
            Exit Function
        End If
    Next i

    BondPullInterp = curveRates(nodeCount)

End Function


' -----------------------------------------------------------------------------
' Spread framework -> which curve the bond's base rate comes from.
'
' Single source of truth: BondPullSpotRate and BondPullCurveNodes both route
' through this, so a new framework code is added in one place.
' -----------------------------------------------------------------------------
Private Function BondPullCurveType(ByVal framework As String) As String

    Select Case UCase$(Trim$(framework))

        Case "G"
            BondPullCurveType = "GOV"

        Case "I", "ASW", "Z", "OAS", "MIXED"
            BondPullCurveType = "SWAP"

        Case "OIS", "SOFR"
            BondPullCurveType = "OIS"

        Case Else
            BondPullCurveType = ""

    End Select

End Function


' Discount factor from a zero rate (in %) over a year fraction:  EXP(-(rate/100)*years)
Private Function DiscountFactorFml(ByVal rateRef As String, ByVal yearsRef As String) As String
    DiscountFactorFml = "=IFERROR(EXP(-(" & rateRef & "/100)*" & yearsRef & "),"""")"
End Function


' Annuity / PV01 level from a discount factor and rate:  (1-DF)/max(eps,rate)
' Annuity / PV01 level from a discount factor and a zero rate:  (1 - DF) / z
'
' UNIT BUG, FIXED: the rate arrives in PERCENTAGE POINTS (2.65 means 2.65%) and
' this divided by it directly, so the annuity came out a HUNDRED TIMES too
' small.  The annuity is a number of years - a 10y swap should give roughly 8 -
' and it was producing 0.08.
'
' It went unnoticed because the two consumers hid it in opposite ways:
'
'   Swaps!V,W  PV = Notional * ModelSpread(pct) * Annuity * FX.  The spread is
'              ALSO in percentage points, so spread_pct / rate_pct equals
'              spread_dec / rate_dec and the error cancelled exactly.  The PV,
'              and therefore Model_PnL, was always right.
'
'   Swaps!X    DV01 = |Notional| * 0.0001 * Annuity * FX.  That 0.0001 is a
'              DECIMAL basis point, so nothing cancelled and the model swap
'              DV01 was 100x too small.
'
' PNL_Attribution never read Swaps!X - it uses the Bloomberg BPV in Swaps!BN -
' so the model DV01 was wrong in a column nothing consumed.  The Dashboard's
' framework validation table DID read it, which is why that table showed every
' bond as ~99% unhedged while the attribution beside it had correctly picked a
' swap framework.
'
' Dividing by z/100 here makes the annuity a true year count.  SwapModelPVFml
' below now divides its spread by 100 to match, so the PV is unchanged.
Private Function AnnuityFml(ByVal dfRef As String, ByVal rateRef As String) As String
    AnnuityFml = "=IFERROR((1-" & dfRef & ")/MAX(0.000001," & rateRef & "/100),"""")"
End Function


' Model PnL = PV_T0 - PV_T-1  (blank unless both are numeric)
Private Function ModelPnLFml(ByVal pv0Ref As String, ByVal pv1Ref As String) As String
    ModelPnLFml = "=IF(OR(NOT(ISNUMBER(" & pv0Ref & ")),NOT(ISNUMBER(" & pv1Ref & "))),""""," & pv0Ref & "-" & pv1Ref & ")"
End Function


' Swap accrual year fraction from as-of (Config current/T0 date) to the swap end date.
Private Function SwapYearFracFml(ByVal guardRef As String, ByVal endRef As String) As String
    SwapYearFracFml = "=IF(" & guardRef & "="""","""",MAX(0,YEARFRAC('" & SH_CONFIG & "'!" & CFG_T1_DATE & "," & endRef & ",1)))"
End Function


' Interpolated OIS rate at the swap tenor (curve object InterpOIS), current (T1) or prior.
Private Function SwapOisRateFml(ByVal ccyRef As String, ByVal yearsRef As String, ByVal useT1 As Boolean) As String
    SwapOisRateFml = "=IF(" & ccyRef & "="""","""",InterpOIS(" & ccyRef & "," & yearsRef & "," & IIf(useT1, "TRUE", "FALSE") & "))"
End Function


' FX cross of a deal currency into EUR base via Bloomberg BDP (guarded, EUR = 1).
Private Function FxToBaseFml(ByVal guardRef As String, ByVal ccyRef As String) As String
    FxToBaseFml = "=IF(" & guardRef & "="""","""",IF(" & ccyRef & "="""",""""," & _
        "LET(_ccy,UPPER(TRIM(" & ccyRef & "))," & _
        "IF(_ccy=""EUR"",1," & _
        "IFERROR(1/BDP(""EUR""&_ccy&"" Curncy"",""PX_LAST"")," & _
        "IFERROR(BDP(_ccy&""EUR Curncy"",""PX_LAST""),""""))))))"
End Function


' Model PV of a swap:  sign(payFixed) * Notional * ModelSpread * Annuity * FX.
Private Function SwapModelPVFml(ByVal ccyRef As String, ByVal notRef As String, ByVal curveRef As String, _
    ByVal payFixedRef As String, ByVal spreadRef As String, ByVal annRef As String, ByVal fxRef As String) As String
    ' spreadRef is in percentage points and the annuity is now in years, so the
    ' spread is converted to decimal here.  Before the annuity fix both sides
    ' were in percentage points and the units cancelled by accident; they now
    ' cancel on purpose, and the PV is numerically unchanged.
    SwapModelPVFml = "=IF(OR(" & ccyRef & "=""""," & notRef & "=""""," & curveRef & "=""""),""""," & _
        "IFERROR(IF(UPPER(TRIM(" & payFixedRef & "))=""Y"",-1,1)*" & notRef & "*(" & spreadRef & "/100)*" & annRef & "*" & fxRef & ",""""))"
End Function


' Model DV01 (EUR) of a swap:  sign(payFixed) * |Notional| * 1bp * Annuity * FX.
Private Function SwapDV01Fml(ByVal ccyRef As String, ByVal notRef As String, ByVal payFixedRef As String, _
    ByVal annRef As String, ByVal fxRef As String) As String
    SwapDV01Fml = "=IF(OR(" & ccyRef & "=""""," & notRef & "=""""),""""," & _
        "IFERROR(IF(UPPER(TRIM(" & payFixedRef & "))=""Y"",-1,1)*ABS(" & notRef & ")*0.0001*" & annRef & "*" & fxRef & ",""""))"
End Function




' =============================================================================
' OFFLINE / KEEP-FORMULAS MODE
'   When Config!B39 (CFG_KEEP_FORMULAS) is TRUE, the workbook is in offline /
'   inspection mode: every derived cell keeps its live Excel formula (nothing is
'   frozen to a static value) and the Bloomberg refresh/wait calls are skipped so
'   the Step buttons run to completion without a Bloomberg terminal.  Only genuine
'   inputs (SQL/OPICS and Hedge_Risco) stay as static values.
'
'   STRICTLY a debug aid - defaults FALSE.  A workbook saved with this ON does NOT
'   hold a locked prior/T-1 snapshot, so it must not be used for an official run.
' =============================================================================
Private Function KeepFormulasMode() As Boolean
    Dim v As String
    v = UCase$(CleanText(ThisWorkbook.Worksheets(SH_CONFIG).Range(CFG_KEEP_FORMULAS).value))
    KeepFormulasMode = (v = "TRUE" Or v = "YES" Or v = "1" Or v = "ON")
End Function




' =============================================================================
' NAMED RANGES  (corrected - the original Names.Add was a no-op)
' =============================================================================


Private Sub SetupNamedRanges()
    ' Define any workbook-level names the formulas rely on here.
    AddOrReplaceName "Cfg_T0", "'" & SH_CONFIG & "'!" & CFG_T0_DATE
    AddOrReplaceName "Cfg_T1", "'" & SH_CONFIG & "'!" & CFG_T1_DATE
    AddOrReplaceName "Cfg_FxFix", "'" & SH_CONFIG & "'!" & CFG_FX_FIX_SOURCE
    AddOrReplaceName "Cfg_FutPxField", "'" & SH_CONFIG & "'!" & CFG_FUT_PRICE_FIELD
    AddOrReplaceName "Cfg_SpreadFwk", "'" & SH_CONFIG & "'!" & CFG_SPREAD_FRAMEWORK
End Sub


Private Sub AddOrReplaceName(ByVal nameText As String, ByVal refersToText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0


    If Left$(refersToText, 1) <> "=" Then refersToText = "=" & refersToText
    ThisWorkbook.Names.Add Name:=nameText, refersTo:=refersToText
End Sub




' =============================================================================
' FUTURES TICKER MAP SHEET  (OPICS contract code -> Bloomberg generic)
'   Replaces blind  code & "1 Comdty".  Verify the "verify" rows on the desk.
' =============================================================================


Private Sub EnsureFutMapSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_FUTMAP)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_FUTMAP
    End If


    ws.Range(FMCOL_OPICS_CODE & FUTMAP_HEADER_ROW & ":" & FMCOL_STATUS & FUTMAP_HEADER_ROW).value = Array("OPICS_Code", "Bloomberg_Generic", "Description", "Status")
    ws.Range(FMCOL_OPICS_CODE & FUTMAP_HEADER_ROW & ":" & FMCOL_STATUS & FUTMAP_HEADER_ROW).Font.Bold = True


    ' Only (re)write seed rows if the table is empty, so manual edits survive.
    If CleanText(ws.Range(FMCOL_OPICS_CODE & FUTMAP_DATA_ROW).value) = "" Then
        Dim m As Variant, i As Long
        m = Array( _
            Array("DU", "DU1 Comdty", "Euro Schatz", "OK"), _
            Array("OE", "OE1 Comdty", "Euro Bobl", "OK"), _
            Array("RX", "RX1 Comdty", "Euro Bund", "OK"), _
            Array("UB", "UB1 Comdty", "Euro Buxl", "OK"), _
            Array("OAT", "OAT1 Comdty", "OAT future", "OK"), _
            Array("IK", "IK1 Comdty", "BTP future", "verify"), _
            Array("FR", "", "Mapping unverified", "verify"), _
            Array("RP", "", "Mapping unverified", "verify"), _
            Array("G", "", "Mapping unverified", "verify"), _
            Array("EC", "", "Mapping unverified", "verify"), _
            Array("UXY", "", "Mapping unverified", "verify"))
        For i = LBound(m) To UBound(m)
            ws.Cells(FUTMAP_DATA_ROW + i - LBound(m), _
                     colNum(FMCOL_OPICS_CODE)).Resize(1, 4).value = m(i)
        Next i
    End If


    AddOrReplaceName "FutMapTable", _
        "'" & SH_FUTMAP & "'!$" & FMCOL_OPICS_CODE & "$" & _
        CStr(FUTMAP_DATA_ROW) & ":$" & FMCOL_BLOOMBERG_GENERIC & "$" & _
        CStr(FUTMAP_NAMED_LAST_ROW)
End Sub


' =============================================================================
' PER-BOND SPREAD FRAMEWORK OVERRIDE SHEET
'
'   ISIN -> framework code.  Consulted by PNL_Attribution!CF ahead of the
'   Config!B18 global override and ahead of the automatic DV01-weighted choice.
'
'   Valid codes:
'     G      bond measured versus the government curve  (G-spread)
'     I      bond measured versus the swap curve        (I-spread)
'     ASW    asset-swap spread
'     Z      Z-spread
'     OAS    option-adjusted spread (use for callables)
'     OIS    no decomposition; whole move taken as -DV01 * Delta_y
'     SOFR   same as OIS, kept for USD books that label it that way
'     MIXED  DV01-weighted blend of the G and I chains
'     REVIEW suppress attribution for this bond and flag it for a human
'
'   This is a MANUAL table.  It lives on its own sheet, not on PNL_Attribution,
'   precisely so that it survives the ClearContents that every PNL rebuild does
'   over A5:CU604.  Rows added here persist until someone deletes them.
' =============================================================================
Private Sub EnsureSpreadOverrideSheet()
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SPREADOVR)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SPREADOVR
    End If


    ws.Range("A1").value = "Per-bond spread framework override (manual - survives PNL rebuilds)"
    ws.Range("A2").value = "Valid: G, I, ASW, Z, OAS, OIS, SOFR, MIXED, REVIEW.  Blank or unlisted ISIN = use Config!B18, else automatic."


    ws.Range("A4:C4").value = Array("ISIN", "Framework", "Reason / owner")
    ws.Range("A4:C4").Font.Bold = True


    ' The lookup is VLOOKUP(ISIN, SpreadOverrideTable, 2, FALSE) so column A must
    ' stay the ISIN key and column B the framework code.
    AddOrReplaceName "SpreadOverrideTable", "'" & SH_SPREADOVR & "'!$A$5:$B$600"
End Sub


Private Function EnsureSwapMapSheet() As Worksheet
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SWAPMAP)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SWAPMAP
    End If


    ws.Range(CELL_SWAPMAP_A1).value = "SwapMap imported from Coverage support"


    ws.Range(SMCOL_SOURCEROW & 4 & ":" & SMCOL_NOTIONAL_SOURCE & 4).value = Array( _
        "SourceRow", _
        "CoverageRelation", _
        "LinkedISIN", _
        "Swap_ID_Source", _
        "Swap_Direct_ID", _
        "Fixed_Leg_ID", _
        "Float_Leg_ID", _
        "ImportStatus", _
        "Counterparty", _
        "Notional", _
        "StartDate", _
        "EndDate", _
        "CCY", _
        "Notional_Source")


    ws.Range(SMCOL_SOURCEROW & 4 & ":" & SMCOL_NOTIONAL_SOURCE & 4).Font.Bold = True
    ws.Range(SMCOL_SOURCEROW & 4 & ":" & SMCOL_NOTIONAL_SOURCE & 4).HorizontalAlignment = xlCenter


    Set EnsureSwapMapSheet = ws
End Function


Private Function EnsureCoverageFuturesSheet() As Worksheet
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_COV_FUTURES)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_COV_FUTURES
    End If


    ws.Range(CELL_COVF_A1).value = "Coverage futures imported from Coverage support"


    ws.Range(CFCOL_SOURCEROW & 4 & ":" & CFCOL_IMPORTSTATUS & 4).value = Array( _
        "SourceRow", _
        "CoverageRelation", _
        "LinkedISIN", _
        "CoverageInfo_D", _
        "FutureLabel", _
        "FutureCode", _
        "HedgeType", _
        "Contracts", _
        "StartDate", _
        "CCY", _
        "ImportStatus")


    ws.Range(CFCOL_SOURCEROW & 4 & ":" & CFCOL_IMPORTSTATUS & 4).Font.Bold = True
    ws.Range(CFCOL_SOURCEROW & 4 & ":" & CFCOL_IMPORTSTATUS & 4).HorizontalAlignment = xlCenter


    Set EnsureCoverageFuturesSheet = ws
End Function


Private Function EnsureCoverageSupportSheet() As Worksheet
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_COV_SUPPORT)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_COV_SUPPORT
    End If


    ws.Range(CELL_COV_A1).value = "Temporary copy of Hedge_Risco Resumo A9:AI220"
    ws.Range(CELL_COV_A2).value = "This sheet is cleared after SwapMap and CoverageFutures are built."
    ws.Range(CELL_COV_A8).value = "Resumo columns A:AI copied below. Column AJ stores formula"
    ws.Range(RNG_COV_A1_A2).Font.Bold = True


    Set EnsureCoverageSupportSheet = ws
End Function


Private Sub ClearCoverageSupportSheet()
    Dim ws As Worksheet


    Set ws = EnsureCoverageSupportSheet()


    ws.Range("A" & SWAP_ID_FIRST_ROW & ":" & HEDGE_RISCO_FORMULA_A_COL & SWAP_ID_LAST_ROW).ClearContents
    ws.Range("A" & SWAP_ID_FIRST_ROW & ":" & HEDGE_RISCO_FORMULA_A_COL & SWAP_ID_LAST_ROW).ClearFormats
End Sub


Private Function IsResumoSwapRow(ByVal formulaText As String) As Boolean
    Dim s As String
    s = UCase$(formulaText)


    IsResumoSwapRow = (InStr(1, s, "MONTH(EDATE(", vbTextCompare) > 0)
End Function


Private Function IsResumoFutureRow(ByVal formulaText As String) As Boolean
    Dim s As String
    s = UCase$(formulaText)


    IsResumoFutureRow = (InStr(1, s, "HYPERLINK(", vbTextCompare) > 0)
End Function


Private Function IsHtCSCoveragePortfolio(ByVal v As Variant) As Boolean
    Dim s As String


    s = UCase$(CleanText(v))
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, "&AMP;", "&")
    s = Replace(s, "AND", "&")


    IsHtCSCoveragePortfolio = False


    If s = "HTC&S" Then IsHtCSCoveragePortfolio = True
    If s = "HTCS" Then IsHtCSCoveragePortfolio = True
    If InStr(1, s, "HTC&S", vbTextCompare) > 0 Then IsHtCSCoveragePortfolio = True
    If InStr(1, s, "HTCS", vbTextCompare) > 0 Then IsHtCSCoveragePortfolio = True
End Function


Private Function IsAlphanumericLabel(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String


    s = CleanText(s)


    If Len(s) = 0 Then
        IsAlphanumericLabel = False
        Exit Function
    End If


    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            IsAlphanumericLabel = True
            Exit Function
        End If
    Next i


    IsAlphanumericLabel = False
End Function


Private Function IsCoverageFutureRow(ByVal valueA As String, ByVal formulaA As String, ByVal typeH As String) As Boolean
    Dim h As String


    valueA = CleanText(valueA)
    formulaA = CleanText(formulaA)
    h = UCase$(CleanText(typeH))


    If Len(valueA) = 0 Then
        IsCoverageFutureRow = False
        Exit Function
    End If


    If IsResumoSwapRow(formulaA) Then
        IsCoverageFutureRow = False
        Exit Function
    End If


    If h = "PLAIN VANILLA SWAP" Then
        IsCoverageFutureRow = False
        Exit Function
    End If


    IsCoverageFutureRow = IsAlphanumericLabel(valueA)
End Function


Private Function FutureCodeFromCoverageLabel(ByVal futureLabel As String) As String
    Dim s As String
    Dim firstToken As String


    s = CleanText(futureLabel)


    If Len(s) = 0 Then
        FutureCodeFromCoverageLabel = ""
        Exit Function
    End If


    firstToken = Split(s, " ")(0)


    ' Example:
    '   OE1 #163 -> OE1 -> OE
    '   RX1 #164 -> RX1 -> RX
    '   OAT1 #165 -> OAT1 -> OAT
    Do While Len(firstToken) > 0 And Right$(firstToken, 1) Like "[0-9]"
        firstToken = Left$(firstToken, Len(firstToken) - 1)
    Loop


    FutureCodeFromCoverageLabel = firstToken
End Function


Private Function CleanSwapId(ByVal v As Variant) As String
    CleanSwapId = CleanText(v)
End Function


Private Function SwapIdToBBGSecurity(ByVal rawId As Variant) As String
    Dim s As String


    s = CleanText(rawId)


    If Len(s) = 0 Then
        SwapIdToBBGSecurity = ""
    ElseIf InStr(1, UCase$(s), " CORP", vbTextCompare) > 0 Then
        SwapIdToBBGSecurity = s
    Else
        SwapIdToBBGSecurity = s & " Corp"
    End If
End Function


Private Function CoverageNotionalValue(ByVal v As Variant) As Variant
    Dim s As String
    Dim n As Double


    If IsError(v) Then
        CoverageNotionalValue = ""
        Exit Function
    End If


    If IsNumeric(v) Then
        CoverageNotionalValue = CDbl(v)
        Exit Function
    End If


    s = UCase$(CleanText(v))


    If Len(s) = 0 Then
        CoverageNotionalValue = ""
        Exit Function
    End If


    s = Replace(s, "EUR", "")
    s = Replace(s, "USD", "")
    s = Replace(s, "GBP", "")
    s = Replace(s, "ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬", "")
    s = Replace(s, " ", "")


    If InStr(1, s, "MM", vbTextCompare) > 0 Then
        s = Replace(s, "MM", "")
        n = val(Replace(s, ",", "."))
        CoverageNotionalValue = n * 1000000#
        Exit Function
    End If


    If Right$(s, 1) = "M" Then
        s = Left$(s, Len(s) - 1)
        n = val(Replace(s, ",", "."))
        CoverageNotionalValue = n * 1000000#
        Exit Function
    End If


    CoverageNotionalValue = CDblSafe(v)
End Function


Private Function HasAnySwapId(ByVal directId As String, ByVal fixedId As String, ByVal floatId As String) As Boolean
    HasAnySwapId = (Len(directId) > 0 Or Len(fixedId) > 0 Or Len(floatId) > 0)
End Function


Private Sub AppendSwapMapRow( _
    ByVal wsMap As Worksheet, _
    ByRef outRow As Long, _
    ByVal sourceRow As Long, _
    ByVal coverageRelation As String, _
    ByVal linkedISIN As String, _
    ByVal idSource As String, _
    ByVal directId As String, _
    ByVal fixedId As String, _
    ByVal floatId As String, _
    ByVal statusText As String, _
    ByVal counterparty As String, _
    ByVal notional As Variant, _
    ByVal startDate As Variant, _
    ByVal endDate As Variant, _
    ByVal ccy As String, _
    ByVal notionalSource As String)


    wsMap.Cells(outRow, colNum(SMCOL_SOURCEROW)).value = sourceRow
    wsMap.Cells(outRow, colNum(SMCOL_COVERAGERELATION)).value = coverageRelation
    wsMap.Cells(outRow, colNum(SMCOL_LINKEDISIN)).value = linkedISIN
    wsMap.Cells(outRow, colNum(SMCOL_SWAP_ID_SOURCE)).value = idSource
    wsMap.Cells(outRow, colNum(SMCOL_SWAP_DIRECT_ID)).value = directId
    wsMap.Cells(outRow, colNum(SMCOL_FIXED_LEG_ID)).value = fixedId
    wsMap.Cells(outRow, colNum(SMCOL_FLOAT_LEG_ID)).value = floatId
    wsMap.Cells(outRow, colNum(SMCOL_IMPORTSTATUS)).value = statusText
    wsMap.Cells(outRow, colNum(SMCOL_COUNTERPARTY)).value = counterparty
    wsMap.Cells(outRow, colNum(SMCOL_NOTIONAL)).value = notional
    wsMap.Cells(outRow, colNum(SMCOL_STARTDATE)).value = startDate
    wsMap.Cells(outRow, colNum(SMCOL_ENDDATE)).value = endDate
    wsMap.Cells(outRow, colNum(SMCOL_CCY)).value = ccy
    wsMap.Cells(outRow, colNum(SMCOL_NOTIONAL_SOURCE)).value = notionalSource


    outRow = outRow + 1
End Sub


Private Function ImportSwapMapFromHedgeRisco(Optional ByVal showMessage As Boolean = False) As Long
    ' -------------------------------------------------------------------------
    ' Legacy compatibility wrapper.
    '
    ' The old direct-import function used the old AppendSwapMapRow signature.
    ' That no longer works because AppendSwapMapRow now also receives:
    '   Counterparty, Notional, StartDate, EndDate, CCY, Notional_Source.
    '
    ' Current correct workflow:
    '   1) RefreshCoverageSupportFromHedgeRisco
    '   2) BuildSwapMapFromCoverageSupport
    '   3) ClearCoverageSupportSheet
    ' -------------------------------------------------------------------------


    Dim importedCount As Long


    ClearCoverageSupportSheet


    RefreshCoverageSupportFromHedgeRisco


    importedCount = BuildSwapMapFromCoverageSupport()


    ClearCoverageSupportSheet


    ImportSwapMapFromHedgeRisco = importedCount


    If showMessage Then
        MsgBox CStr(importedCount) & " swap mapping rows imported into SwapMap.", vbInformation
    End If
End Function
Private Function AppendSwapMapRowsToSwaps(ByVal wsSw As Worksheet, ByVal currentSwapCount As Long) As Long


    Dim wsMap As Worksheet
    Dim wsBnd As Worksheet


    Dim lastMapRow As Long
    Dim mapRow As Long
    Dim outRow As Long
    Dim maxRow As Long
    Dim scanLastRow As Long
    Dim appendedCount As Long


    Dim directId As String
    Dim fixedId As String
    Dim floatId As String
    Dim linkedISIN As String
    Dim coverageRelation As String
    Dim idSource As String
    Dim baseSwapId As String
    Dim mappingRowId As String
    Dim displaySwapId As String


    Dim counterparty As String
    Dim ccy As String
    Dim notional As Variant
    Dim startDate As Variant
    Dim endDate As Variant
    Dim notionalSource As String


    Dim existing As Object
    Dim r As Long
    Dim key As String
    Dim sourceRow As String


    Set wsMap = EnsureSwapMapSheet()
    Set wsBnd = ThisWorkbook.Worksheets(SH_BONDS)
    Set existing = CreateObject("Scripting.Dictionary")


    ' Write ceiling is the runaway guard, NOT a swap-count assumption: however
    ' many swaps Hedge Risco hands over, all of them get a row.
    maxRow = DATA_ROW + MAX_SHEET_ROWS - 1

    ' Only scan rows that exist; the ceiling above is far too high to loop over.
    scanLastRow = LastSwapDataRow(wsSw)


    ' -------------------------------------------------------------------------
    ' Existing mapping rows are keyed by Swaps!A.
    '
    ' Swaps!A must be unique per mapping row, not only per underlying swap.
    ' One swap can hedge multiple bonds, so the same direct Bloomberg ID is
    ' allowed across several LinkedISIN values.
    ' -------------------------------------------------------------------------
    For r = DATA_ROW To scanLastRow


        key = IdText(wsSw.Cells(r, WCOL_DEALID).value)


        If Len(key) > 0 Then
            existing(key) = True
        End If


    Next r


    outRow = DATA_ROW + currentSwapCount


    If outRow < DATA_ROW Then outRow = DATA_ROW


    If outRow > maxRow Then
        AppendSwapMapRowsToSwaps = currentSwapCount
        Exit Function
    End If


    lastMapRow = wsMap.Cells(wsMap.Rows.Count, SMCOL_SOURCEROW).End(xlUp).Row


    If lastMapRow < 5 Then
        AppendSwapMapRowsToSwaps = currentSwapCount
        Exit Function
    End If


    ' Force identifier and mapping columns to text.
    wsSw.Range(WCOL_DEALID & DATA_ROW & ":" & WCOL_DEALID & maxRow).numberFormat = "@"
    wsSw.Range(WCOL_BBG_SWAP_DIRECT_ID & DATA_ROW & ":" & WCOL_BBG_FLOAT_LEG_ID & maxRow).numberFormat = "@"
    wsSw.Range(WCOL_LINKEDISIN & DATA_ROW & ":" & WCOL_LINKEDISIN & maxRow).numberFormat = "@"


    For mapRow = 5 To lastMapRow


        If outRow > maxRow Then Exit For


        sourceRow = IdText(wsMap.Cells(mapRow, SMCOL_SOURCEROW).value)


        coverageRelation = CleanText(wsMap.Cells(mapRow, SMCOL_COVERAGERELATION).value)
        linkedISIN = UCase$(CleanText(wsMap.Cells(mapRow, SMCOL_LINKEDISIN).value))
        idSource = UCase$(CleanText(wsMap.Cells(mapRow, SMCOL_SWAP_ID_SOURCE).value))


        directId = IdText(wsMap.Cells(mapRow, SMCOL_SWAP_DIRECT_ID).value)
        fixedId = IdText(wsMap.Cells(mapRow, SMCOL_FIXED_LEG_ID).value)
        floatId = IdText(wsMap.Cells(mapRow, SMCOL_FLOAT_LEG_ID).value)


        counterparty = CleanText(wsMap.Cells(mapRow, SMCOL_COUNTERPARTY).value)


        notional = wsMap.Cells(mapRow, SMCOL_NOTIONAL).value
        startDate = wsMap.Cells(mapRow, SMCOL_STARTDATE).value
        endDate = wsMap.Cells(mapRow, SMCOL_ENDDATE).value
        ccy = CleanText(wsMap.Cells(mapRow, SMCOL_CCY).value)
        notionalSource = CleanText(wsMap.Cells(mapRow, SMCOL_NOTIONAL_SOURCE).value)


        ' ---------------------------------------------------------------------
        ' Base swap identifier.
        '
        ' Prefer the direct Bloomberg swap ID.
        ' If there is no direct ID, use the fixed/float leg combination.
        ' If no Bloomberg IDs exist, construct a traceable source-row ID.
        ' ---------------------------------------------------------------------
        If Len(directId) > 0 Then


            baseSwapId = directId


        ElseIf Len(fixedId) > 0 Or Len(floatId) > 0 Then


            baseSwapId = "LEGS_" & fixedId & "_" & floatId


        Else


            baseSwapId = idSource & "_ROW_" & sourceRow


        End If


        ' ---------------------------------------------------------------------
        ' Unique mapping-row identifier.
        '
        ' The same swap can hedge multiple bonds, and one coverage-support row
        ' can contain both a plain and a synthetic swap alternative.
        '
        ' Include idSource so PLAIN and SYNTHETIC rows can never collide even if
        ' the source data unexpectedly contains the same external identifier.
        ' ---------------------------------------------------------------------
        mappingRowId = _
            idSource & "|" & _
            baseSwapId & "|" & _
            linkedISIN & "|" & _
            sourceRow


        ' Human-readable contract reference.
        displaySwapId = baseSwapId


        If Len(mappingRowId) > 0 And Len(linkedISIN) > 0 Then


            If Not existing.Exists(mappingRowId) Then


                With wsSw


                    ' ---------------------------------------------------------
                    ' A = unique internal mapping-row ID.
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_DEALID).numberFormat = "@"
                    .Cells(outRow, WCOL_DEALID).value = mappingRowId


                    ' ---------------------------------------------------------
                    ' Core swap economics.
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_CCY).value = ccy
                    .Cells(outRow, WCOL_NOTIONAL).value = notional
                    .Cells(outRow, WCOL_STARTDATE).value = startDate
                    .Cells(outRow, WCOL_ENDDATE).value = endDate
                    .Cells(outRow, WCOL_PORTFOLIO).value = coverageRelation


                    ' ---------------------------------------------------------
                    ' Mapping to the covered bond.
                    '
                    ' L = LinkedISIN
                    ' M = Counterparty
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_LINKEDISIN).numberFormat = "@"
                    .Cells(outRow, WCOL_LINKEDISIN).value = linkedISIN
                    .Cells(outRow, WCOL_CPTY).value = counterparty


                    ' ---------------------------------------------------------
                    ' Initial status before Bloomberg processing.
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_STATUS).value = "BQL mapping imported"
                    .Cells(outRow, WCOL_PNL_SOURCE).value = "BQL_PENDING"
                    .Cells(outRow, WCOL_AF_STATUS).value = "BQL mapping imported"


                    ' ---------------------------------------------------------
                    ' AD:AF = actual Bloomberg swap identifiers.
                    '
                    ' These are used for current/T0 and prior/T-1 Bloomberg
                    ' market-value retrieval.
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_BBG_SWAP_DIRECT_ID).numberFormat = "@"
                    .Cells(outRow, WCOL_BBG_FIXED_LEG_ID).numberFormat = "@"
                    .Cells(outRow, WCOL_BBG_FLOAT_LEG_ID).numberFormat = "@"


                    .Cells(outRow, WCOL_BBG_SWAP_DIRECT_ID).value = directId
                    .Cells(outRow, WCOL_BBG_FIXED_LEG_ID).value = fixedId
                    .Cells(outRow, WCOL_BBG_FLOAT_LEG_ID).value = floatId


                    ' ---------------------------------------------------------
                    ' AG:AK = mapping classification and traceability.
                    '
                    ' AG = PLAIN or SYNTHETIC
                    ' AH = CoverageRelation
                    ' AI = source row
                    ' AJ = instrument class
                    ' AK = import status
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_SWAP_ID_SOURCE).value = idSource
                    .Cells(outRow, WCOL_COVERAGERELATION).value = coverageRelation
                    .Cells(outRow, WCOL_SWAPMAP_SOURCEROW).value = sourceRow
                    .Cells(outRow, WCOL_SWAPMAP_CLASS).value = "SWAP"
                    .Cells(outRow, WCOL_SWAPMAP_STATUS).value = wsMap.Cells(mapRow, SMCOL_IMPORTSTATUS).value


                    ' ---------------------------------------------------------
                    ' AL:AP = mapped economics.
                    ' ---------------------------------------------------------
                    .Cells(outRow, WCOL_MAP_NOTIONAL).value = notional
                    .Cells(outRow, WCOL_MAP_CCY).value = ccy
                    .Cells(outRow, WCOL_MAP_COUNTERPARTY).value = counterparty
                    .Cells(outRow, WCOL_NOTIONAL_FINAL).value = Abs(CDblSafe(notional))
                    .Cells(outRow, WCOL_NOTIONAL_SOURCE).value = notionalSource


                End With


                ' -------------------------------------------------------------
                ' Bonds!BT = plain-vanilla SwapLink only.
                '
                ' Do not append synthetic IDs to Bonds!BT.
                '
                ' Synthetic swaps remain in:
                '   SwapMap
                '   Swaps
                '   Bloomberg valuation
                '   PNL_Attribution
                '
                ' Synthetic mapping/results continue to work through:
                '   Swaps!L  = LinkedISIN
                '   Swaps!AG = SYNTHETIC
                ' -------------------------------------------------------------
                If idSource = "PLAIN" Then
                    LinkSwapToBond wsBnd, linkedISIN, displaySwapId
                End If


                existing(mappingRowId) = True
                appendedCount = appendedCount + 1
                outRow = outRow + 1


            End If


        End If


    Next mapRow


    AppendSwapMapRowsToSwaps = currentSwapCount + appendedCount


End Function
Private Function AppendCoverageFuturesRowsToFutures(ByVal wsFut As Worksheet, ByVal wsBnd As Worksheet) As Long


    Dim wsMap As Worksheet
    Dim lastMapRow As Long
    Dim mapRow As Long
    Dim outRow As Long
    Dim maxRow As Long
    Dim appendedCount As Long


    Dim sourceRow As Variant
    Dim coverageRelation As String
    Dim linkedISIN As String
    Dim coverageInfoD As String
    Dim futureLabel As String
    Dim futureCode As String
    Dim hedgeType As String
    Dim contracts As Variant
    Dim startDate As Variant
    Dim ccy As String
    Dim statusText As String
    Dim hedgeSource As String
    Dim coverageBpv As Variant


    Set wsMap = EnsureCoverageFuturesSheet()


    ' Runaway guard, not a contract-count assumption.
    maxRow = DATA_ROW + MAX_SHEET_ROWS - 1
    outRow = DATA_ROW


    ' -------------------------------------------------------------------------
    ' New Futures physical layout after deleting old Z and old AB:
    '
    '   AG = Status
    '   AH = Coverage_SourceRow
    '   AI = CoverageRelation
    '   AJ = FutureLabel
    '   AK = HedgeType
    '   AL = Coverage_StartDate
    '   AM = CoverageInfo_D
    '
    ' Therefore final active Futures column is now AM, not AO.
    ' -------------------------------------------------------------------------
    ' Clear as far as the sheet has ever been written, so contracts dropped since
    ' the last run do not survive as stale rows.
    wsFut.Range(FCOL_CONTRACTCODE & DATA_ROW & ":" & FCOL_HEDGE_CLASS & _
        CStr(SheetClearLastRow(wsFut, DATA_ROW, LastFutureDataRow(wsFut)))).ClearContents


    lastMapRow = wsMap.Cells(wsMap.Rows.Count, CFCOL_SOURCEROW).End(xlUp).Row


    If lastMapRow < 5 Then
        AppendCoverageFuturesRowsToFutures = 0
        Exit Function
    End If


    For mapRow = 5 To lastMapRow


        If outRow > maxRow Then Exit For


        sourceRow = wsMap.Cells(mapRow, CFCOL_SOURCEROW).value
        coverageRelation = CleanText(wsMap.Cells(mapRow, CFCOL_COVERAGERELATION).value)
        linkedISIN = UCase$(CleanText(wsMap.Cells(mapRow, CFCOL_LINKEDISIN).value))
        coverageInfoD = CleanText(wsMap.Cells(mapRow, CFCOL_COVERAGEINFO_D).value)
        futureLabel = CleanText(wsMap.Cells(mapRow, CFCOL_FUTURELABEL).value)
        futureCode = CleanText(wsMap.Cells(mapRow, CFCOL_FUTURECODE).value)
        hedgeType = CleanText(wsMap.Cells(mapRow, CFCOL_HEDGETYPE).value)
        contracts = wsMap.Cells(mapRow, CFCOL_CONTRACTS).value
        startDate = wsMap.Cells(mapRow, CFCOL_STARTDATE).value
        ccy = CleanText(wsMap.Cells(mapRow, CFCOL_CCY).value)
        statusText = CleanText(wsMap.Cells(mapRow, CFCOL_IMPORTSTATUS).value)
        hedgeSource = CleanText(wsMap.Cells(mapRow, CFCOL_HEDGE_SOURCE).value)
        coverageBpv = wsMap.Cells(mapRow, CFCOL_BPV).value


        ' A row with no source tag predates the two-book split.  Default it to
        ' Tx Juro rather than leaving it blank: blank matches NEITHER SUMIFS on
        ' PNL_Attribution, so the row would carry risk that no bond can see.
        If Len(hedgeSource) = 0 Then hedgeSource = HEDGE_SOURCE_RTJ


        If Len(futureCode) > 0 Then


            With wsFut


                ' A:K = core futures source fields.
                .Cells(outRow, FCOL_CONTRACTCODE).value = futureCode
                .Cells(outRow, FCOL_CCY).value = ccy
                .Cells(outRow, FCOL_CONTRACTS).value = contracts
                .Cells(outRow, FCOL_PORTFOLIO).value = coverageRelation
                .Cells(outRow, FCOL_LINKEDISIN).value = linkedISIN


                ' -----------------------------------------------------------------
                ' New shifted coverage/status area.
                '
                ' Old layout:
                '   AI = status
                '   AJ:AO = coverage fields
                '
                ' New layout:
                '   AG = status
                '   AH:AM = coverage fields
                ' -----------------------------------------------------------------
                .Cells(outRow, FCOL_STATUS).value = statusText
                .Cells(outRow, FCOL_COVERAGE_SOURCEROW).value = sourceRow
                .Cells(outRow, FCOL_COVERAGERELATION).value = coverageRelation
                .Cells(outRow, FCOL_FUTURELABEL).value = futureLabel
                .Cells(outRow, FCOL_HEDGETYPE).value = hedgeType
                .Cells(outRow, FCOL_COVERAGE_STARTDATE).value = startDate
                .Cells(outRow, FCOL_COVERAGEINFO_D).value = coverageInfoD
                .Cells(outRow, FCOL_HEDGE_SOURCE).value = hedgeSource
                .Cells(outRow, FCOL_COVERAGE_BPV).value = coverageBpv


            End With


            LinkFutureToBond wsBnd, linkedISIN, futureLabel


            appendedCount = appendedCount + 1
            outRow = outRow + 1


        End If


    Next mapRow


    AppendCoverageFuturesRowsToFutures = appendedCount


End Function
Private Sub EnrichCoverageFuturesFromOPICS( _
    ByVal conn As Object, _
    ByVal wsFut As Worksheet, _
    ByVal wsCfg As Worksheet, _
    ByVal isinType As String, _
    ByVal asOf As String)


    Dim lastRow As Long
    Dim codeList As String
    Dim branchCode As String
    Dim sql As String
    Dim rs As Object
    Dim d As Object
    Dim r As Long
    Dim k As String
    Dim arr As Variant


    lastRow = LastFutureDataRow(wsFut)
    If lastRow < DATA_ROW Then Exit Sub


    codeList = BuildLoadedFutureCodeInClause(wsFut, lastRow)
    If Len(codeList) = 0 Then Exit Sub


    branchCode = CleanText(wsCfg.Range(CFG_BR).value)


    sql = ""
    sql = sql & "WITH CTD AS ("
    sql = sql & " SELECT BR, CONTCODE, DELVDATE, CONVFACTOR_8, UNDSECID,"
    sql = sql & " ROW_NUMBER() OVER(PARTITION BY BR, CONTCODE ORDER BY DELVDATE) AS RN"
    sql = sql & " FROM dbo.FDEL"
    sql = sql & " WHERE DELVDATE >= '" & asOf & "'"
    If Len(branchCode) > 0 Then
        sql = sql & " AND BR = '" & EscapeSQL(branchCode) & "'"
    End If
    sql = sql & "),"
    sql = sql & " AE AS ("
    sql = sql & " SELECT BR, CONTCODE,"
    sql = sql & " SUM(NUMCONT * CONTPRICE_8) / NULLIF(SUM(ABS(NUMCONT)),0) AS AvgPx"
    sql = sql & " FROM dbo.FFDH"
    sql = sql & " WHERE VERIND='Y' AND (REVREASON IS NULL OR REVREASON='')"
    If Len(branchCode) > 0 Then
        sql = sql & " AND BR = '" & EscapeSQL(branchCode) & "'"
    End If
    sql = sql & " GROUP BY BR, CONTCODE"
    sql = sql & ")"
    sql = sql & " SELECT fc.CONTCODE AS ContractCode,"
    sql = sql & " fc.EXCHANGE AS Exchange,"
    sql = sql & " fc.FACEVALUE AS FaceValue,"
    sql = sql & " ctd.DELVDATE AS DelivDate,"
    sql = sql & " isin.SECALTID AS CTD_ISIN,"
    sql = sql & " ctd.CONVFACTOR_8 AS CTD_CF,"
    sql = sql & " ae.AvgPx AS AvgEntryPx"
    sql = sql & " FROM dbo.FCON fc"
    sql = sql & " LEFT JOIN CTD ctd ON ctd.BR = fc.BR AND ctd.CONTCODE = fc.CONTCODE AND ctd.RN = 1"
    sql = sql & " LEFT JOIN dbo.ASID isin ON isin.SECID = ctd.UNDSECID"
    sql = sql & " AND isin.SECIDTYPE = '" & EscapeSQL(isinType) & "'"
    sql = sql & " LEFT JOIN AE ae ON ae.BR = fc.BR AND ae.CONTCODE = fc.CONTCODE"
    sql = sql & " WHERE fc.CONTCODE IN (" & codeList & ")"
    If Len(branchCode) > 0 Then
        sql = sql & " AND fc.BR = '" & EscapeSQL(branchCode) & "'"
    End If


    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1


    Set d = CreateObject("Scripting.Dictionary")


    Do While Not rs.EOF
        k = CleanText(Fld(rs, "ContractCode"))
        If Len(k) > 0 Then
            d(k) = Array( _
                NullToStr(Fld(rs, "Exchange")), _
                NullToNumBlank(Fld(rs, "FaceValue")), _
                NullToDate(Fld(rs, "DelivDate")), _
                NullToStr(Fld(rs, "CTD_ISIN")), _
                NullToNumBlank(Fld(rs, "CTD_CF")), _
                NullToNumBlank(Fld(rs, "AvgEntryPx")) _
            )
        End If
        rs.MoveNext
    Loop


    rs.Close


    For r = DATA_ROW To lastRow
        k = CleanText(wsFut.Cells(r, colNum(FCOL_CONTRACTCODE)).value)
        If d.Exists(k) Then


            arr = d(k)


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_EXCHANGE)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_EXCHANGE) _
                ).value = arr(0)


            End If


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_FACEVALUE)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_FACEVALUE) _
                ).value = arr(1)


            End If


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_DELIVDATE)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_DELIVDATE) _
                ).value = arr(2)


            End If


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_CTD_ISIN)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_CTD_ISIN) _
                ).value = arr(3)


            End If


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_CTD_CF)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_CTD_CF) _
                ).value = arr(4)


            End If


            If CleanText( _
                wsFut.Cells(r, colNum(FCOL_AVGENTRYPX)).value _
            ) = "" Then


                wsFut.Cells( _
                    r, _
                    colNum(FCOL_AVGENTRYPX) _
                ).value = arr(5)


            End If


        End If
    Next r
End Sub


Private Function BuildLoadedFutureCodeInClause(ByVal wsFut As Worksheet, ByVal lastRow As Long) As String
    Dim d As Object
    Dim r As Long
    Dim code As String
    Dim s As String
    Dim key As Variant


    Set d = CreateObject("Scripting.Dictionary")


    For r = DATA_ROW To lastRow
        code = CleanText(wsFut.Cells(r, FCOL_CONTRACTCODE).value)
        If Len(code) > 0 Then d(code) = True
    Next r


    For Each key In d.Keys
        If Len(s) > 0 Then s = s & ","
        s = s & "'" & EscapeSQL(CStr(key)) & "'"
    Next key


    BuildLoadedFutureCodeInClause = s
End Function


' =============================================================================
' DIAGNOSTIC - TEST OPICS CONNECTION
' =============================================================================


Public Sub Test_OPICS_Connection()
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)


    Dim serverName As String, uid As String
    serverName = Replace(CleanText(wsCfg.Range(CFG_DSN).value), "/", "\")
    uid = CleanText(wsCfg.Range(CFG_UID).value)


    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")


    If OpenOPICS(conn, wsCfg) Then
        MsgBox "OPICS connection SUCCESS." & vbCrLf & vbCrLf & _
               "Server: " & serverName & vbCrLf & _
               "Excel: " & GetExcelBitness() & vbCrLf & _
               "Auth: " & IIf(uid = "", "Windows", "SQL user/pass"), _
               vbInformation, "OPICS Connection Test"
        On Error Resume Next
        conn.Close
        On Error GoTo 0
    End If
End Sub




' =============================================================================
' THE THREE BUTTONS
'
' A daily run is three presses, left to right:
'
'   1  Button1_LoadBonds                Bonds     - positions, then their formulas
'   2  Button2_LoadHedgesAndAttribute   Futures, Swaps, PNL_Attribution
'   3  Button3_BuildDashboard           Dashboard
'
' and one maintenance macro that is NOT a button:
'
'      Setup_Workbook_Layout            once, or after a sheet is added
'
' WHERE THE RUN GOES AFTERWARDS
'
'   Buttons 1 and 2 each end by calling modAccess.Access_AutoSaveAfterLoad,
'   which stores the positions in Access under a new RunID.  It is silent and
'   cannot fail the button that called it: a database that is unreachable must
'   not lose a load that has already succeeded.  Saving twice with nothing
'   changed does not produce two runs - the second fingerprints the same and is
'   refused - so Button 1's save and Button 2's save minutes later are one run's
'   worth of history, not two.  See docs/ACCESS_STORE.md.
'
' WHY THERE IS NO "WRITE FORMULAS" BUTTON
'
'   There used to be one, and pressing it was the only thing standing between a
'   freshly loaded book and a sheet full of blanks.  But a formula is only ever
'   missing for one reason: a row appeared that did not have one yet.  The load
'   is the moment that happens and the only moment that knows how many rows
'   there now are, so the load is where the formulas belong.  Splitting them
'   apart created a state - rows loaded, formulas not written - that is always
'   wrong and that the workbook could not tell you it was in.
'
'   Every writer still exists and none of them changed.  They are called from
'   the button that creates the rows they fill.
'
' WHY THERE IS NO "REFRESH MARKET DATA" BUTTON
'
'   Nothing is frozen.  Both snapshots are live formulas - BDP for T0, BQL dated
'   from Config!B4 for T-1 - so the prices are whatever Bloomberg last answered,
'   and re-running reproduces them exactly.  There is no captured value to
'   refresh.
'
'   What DID have to survive is the full rebuild.  Several model quantities come
'   from UDFs that read OIS_Curves directly (InterpOIS, InterpGov, InterpSwap,
'   BondPullToParPrice).  Excel cannot see those reads in its dependency graph,
'   so an ordinary recalculation leaves them holding the PREVIOUS run's numbers -
'   silently, and looking entirely plausible.  Application.CalculateFullRebuild
'   is the only thing that fixes it, and it now runs inside Buttons 2 and 3
'   rather than behind a button someone has to remember to press.
'
' WHAT THIS MEANS FOR A COLUMN CHANGE
'
'   Adding a column to PNL_Attribution is one edit to PnlLayout plus one writer
'   line in WritePNLRow.  Button 2 picks both up.  See docs/COLUMNS.md.
' =============================================================================


' -----------------------------------------------------------------------------
' BUTTON 1  -  Load the bond book, then give every bond its formulas
'
' The OPICS query owns Bonds!A:K and its conditions are unchanged; this refreshes
' it in place.  Everything from BCOL_DAYSLEFT rightwards is macro-owned, cleared
' by the load and rewritten here for exactly the rows that came back.
'
' Curves are written here too.  They are not bond data, but every bond spread on
' the sheet is measured against them, and a bond formula written against an
' empty curve sheet returns blank rather than wrong - which reads as "no PnL"
' instead of "no curve".
' -----------------------------------------------------------------------------
Public Sub Button1_LoadBonds()

    Dim wb As Workbook
    Dim wsOIS As Worksheet
    Dim wsBnd As Worksheet
    Dim lastBondRow As Long
    Dim problems As String
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    Set wb = ThisWorkbook
    Set wsOIS = wb.Worksheets(SH_OIS)
    Set wsBnd = wb.Worksheets(SH_BONDS)

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo Button1Fail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Application.StatusBar = "Button 1: refreshing the bond query..."
    LoadOPICS_Bonds silent:=True

    ' The query refresh is the moment the data was retrieved, so it is the moment
    ' the run is named after.  forceNew because Button 1 STARTS a run: pressing it
    ' again is a new pull and belongs in a new database, whatever the previous
    ' stamp said.  Button 2 then joins this run rather than opening its own.
    AccBeginRun forceNew:=True

    lastBondRow = LastBondDataRow(wsBnd)

    If lastBondRow < BOND_DATA_ROW Then
        Application.StatusBar = False
        MsgBox "No bonds came back from the query. Nothing else to do.", _
               vbExclamation, "Button 1"
        GoTo Button1Exit
    End If

    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Button 1: writing curve formulas..."
    problems = problems & WriteCurveFormulasSafe(wsOIS)

    Application.StatusBar = "Button 1: writing bond formulas (" & _
                            CStr(BondRowCount(lastBondRow)) & " rows)..."
    problems = problems & WriteBondFormulasSafe(wsBnd, lastBondRow)

    Application.Calculation = oldCalc

    ' Ask Bloomberg, wait for it, then rebuild.  The wait is why this is a call
    ' and not just CalculateFullRebuild - see RefreshMarketDataCore.
    Application.StatusBar = "Button 1: refreshing market data..."
    problems = problems & RefreshMarketDataCore()

    ' Store the run BEFORE the message box, so the positions are safe even if
    ' the next thing anyone does is close the book without reading it.  Silent
    ' and unable to fail the button - see Access_AutoSaveAfterLoad.
    Application.StatusBar = "Button 1: saving the run..."
    Access_AutoSaveAfterLoad "Button 1"

    Application.StatusBar = False

    If Len(problems) = 0 Then
        MsgBox CStr(BondRowCount(lastBondRow)) & _
               " bonds loaded and given their formulas." & vbCrLf & vbCrLf & _
               "Next: Button 2 (load hedges and attribute).", _
               vbInformation, "Button 1 complete"
    Else
        MsgBox CStr(BondRowCount(lastBondRow)) & " bonds loaded, but:" & _
               vbCrLf & vbCrLf & problems, _
               vbExclamation, "Button 1 finished with problems"
    End If

Button1Exit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    Exit Sub

Button1Fail:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    MsgBox "Button 1 failed:" & vbCrLf & vbCrLf & CaptureErrInfo(), _
           vbCritical, "Button 1"

End Sub


' -----------------------------------------------------------------------------
' BUTTON 2  -  Load the hedges, give them their formulas, attribute them
'
' Three things, in an order that matters:
'
'   1. Load the hedge ROWS.  Swaps and futures both come from Hedge Risco Tx
'      Juro; futures also come from Hedge Risco Total, and the two books are
'      counted separately so a book that stopped arriving is visible instead of
'      being netted away inside one total.
'
'   2. Write their FORMULAS - futures and swaps alike.  Futures are the half
'      that is easy to forget, because the swap leg is the one people look at:
'      the futures sheet carries the CTD chain, the conversion factor, the
'      contract multiplier and the unit DV01, and PNL_Attribution's
'      FuturesRTJ_DV01 / FuturesRT_DV01 / Actual_Futures_PnL are all SUMIFS over
'      those columns.  Skip them and the futures hedge silently reads as zero -
'      a bond that is fully hedged looks completely unhedged, and its whole
'      duration move lands in the residual.
'
'   3. Attribute: rebuild PNL_Attribution, one row per bond, and publish the
'      column names the Dashboard reads.  This is where the hedges meet the
'      bonds - each row's hedge columns are SUMIFS over the futures and swaps
'      whose LinkedISIN is that bond.
'
' Then a full rebuild, so the UDF-driven cells hold today's curve rather than
' the last run's.
' -----------------------------------------------------------------------------
Public Sub Button2_LoadHedgesAndAttribute()

    Dim wb As Workbook
    Dim wsBnd As Worksheet
    Dim wsFut As Worksheet
    Dim wsSw As Worksheet
    Dim wsPnl As Worksheet

    Dim lastBondRow As Long
    Dim lastFutRow As Long
    Dim lastSwapRow As Long
    Dim problems As String

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    Set wb = ThisWorkbook
    Set wsBnd = wb.Worksheets(SH_BONDS)
    Set wsFut = wb.Worksheets(SH_FUTURES)
    Set wsSw = wb.Worksheets(SH_SWAPS)
    Set wsPnl = wb.Worksheets(SH_PNL)

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo Button2Fail

    lastBondRow = LastBondDataRow(wsBnd)

    If lastBondRow < BOND_DATA_ROW Then
        MsgBox "No bonds on the Bonds sheet. Run Button 1 first.", _
               vbExclamation, "Button 2"
        GoTo Button2Exit
    End If

    ' --- 1. the hedge rows ---------------------------------------------------
    Application.StatusBar = "Button 2: loading hedges..."
    LoadHedges_Step2 silent:=True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastFutRow = LastFutureDataRow(wsFut)
    lastSwapRow = LastSwapDataRow(wsSw)

    ' --- 2. their formulas ---------------------------------------------------
    Application.StatusBar = "Button 2: writing futures formulas (" & _
                            CStr(HedgeRowCount(lastFutRow)) & " rows)..."
    problems = problems & WriteFuturesFormulasSafe(wsFut, lastFutRow)

    Application.StatusBar = "Button 2: writing swap formulas (" & _
                            CStr(HedgeRowCount(lastSwapRow)) & " rows)..."
    problems = problems & WriteSwapFormulasSafe(wsSw, lastSwapRow)

    ' --- 3. the attribution --------------------------------------------------
    Application.StatusBar = "Button 2: attributing " & _
                            CStr(BondRowCount(lastBondRow)) & " bonds..."
    problems = problems & WritePNLSectionSafe(wsPnl, wsBnd, lastBondRow)

    PublishPnlColumnNames DATA_ROW + PnlRowCount(lastBondRow) - 1

    ' --- and the rebuild -----------------------------------------------------
    Application.Calculation = oldCalc

    Application.StatusBar = "Button 2: refreshing market data..."
    problems = problems & RefreshMarketDataCore()

    ' The full population - bonds, swaps and futures - now exists, so this is
    ' the save that matters.  The Button 1 save minutes ago held bonds only;
    ' this one supersedes it rather than duplicating it, because the
    ' fingerprint covers all three sheets.
    Application.StatusBar = "Button 2: saving the run..."
    Access_AutoSaveAfterLoad "Button 2"

    Application.StatusBar = False

    MsgBox _
        CStr(BondRowCount(lastBondRow)) & " bonds attributed against " & _
        CStr(HedgeRowCount(lastFutRow)) & " futures rows and " & _
        CStr(HedgeRowCount(lastSwapRow)) & " swap rows." & vbCrLf & vbCrLf & _
        IIf(Len(problems) = 0, _
            "Next: Button 3 (build the Dashboard).", _
            "Problems:" & vbCrLf & problems), _
        IIf(Len(problems) = 0, vbInformation, vbExclamation), _
        "Button 2 complete"

Button2Exit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    Exit Sub

Button2Fail:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    MsgBox "Button 2 failed:" & vbCrLf & vbCrLf & CaptureErrInfo(), _
           vbCritical, "Button 2"

End Sub


' -----------------------------------------------------------------------------
' BUTTON 3  -  Build the Dashboard
'
' A full rebuild first, for the same UDF reason as Button 2: the Dashboard reads
' PNL_Attribution, and PNL_Attribution's curve-dependent cells are only current
' after a full rebuild.  Building on top of a stale sheet produces a Dashboard
' that is internally consistent and quietly a day old.
' -----------------------------------------------------------------------------
Public Sub Button3_BuildDashboard()

    Application.StatusBar = "Button 3: recalculating before the build..."
    Application.CalculateFullRebuild
    DoEvents
    Application.StatusBar = False

    BuildDashboard_Step6

End Sub


' -----------------------------------------------------------------------------
' Maintenance, not a button.  Creates or repairs every sheet and its headers.
' Run once on a new workbook, or after changing PnlLayout.
' -----------------------------------------------------------------------------
Public Sub Setup_Workbook_Layout()
    SetupWorkbookFinalLayout
End Sub


' -----------------------------------------------------------------------------
' Section writers.
'
' Each owns its own error handler and returns a description of what went wrong
' rather than raising, so one failing section names itself and the rest of the
' button still runs.  A blanket "On Error Resume Next" in the CALLER cannot do
' this: it resumes in the caller, so an error partway through a writer abandons
' that writer silently and leaves the sheet half written.
' -----------------------------------------------------------------------------
Private Function WriteCurveFormulasSafe(ByVal wsOIS As Worksheet) As String

    On Error GoTo Fail

    ValidateCurrentCurveTickers wsOIS
    WriteCurves_T1_BDP_Efficient wsOIS
    WriteCurveT0BQLFormulas wsOIS
    WriteCurveDerivedFormulas_Efficient wsOIS

    Exit Function

Fail:
    WriteCurveFormulasSafe = "  - Curves: " & Err.Description & vbCrLf

End Function


Private Function WriteBondFormulasSafe( _
    ByVal wsBnd As Worksheet, _
    ByVal lastBondRow As Long) As String

    On Error GoTo Fail

    If lastBondRow < BOND_DATA_ROW Then Exit Function

    ' Order matters: the ticker has to exist before any BDP that keys off it,
    ' and Calc / ConvexityBump write last because they overwrite intermediate
    ' columns the earlier writers fill.
    WriteBondSecurityCandidates wsBnd, lastBondRow
    WriteBondTickerFormulas wsBnd, lastBondRow
    WriteBonds_T1_BDP_Efficient wsBnd, lastBondRow
    WriteBondT0BQLFormulas wsBnd, lastBondRow
    WriteBonds_T0_DerivedSpreads wsBnd, lastBondRow
    WriteBondsCalculatedFormulas_Efficient wsBnd, lastBondRow
    WriteBondConvexityBumpFormulas wsBnd, lastBondRow

    Exit Function

Fail:
    WriteBondFormulasSafe = "  - Bonds: " & Err.Description & vbCrLf

End Function


Private Function WriteFuturesFormulasSafe( _
    ByVal wsFut As Worksheet, _
    ByVal lastFutRow As Long) As String

    On Error GoTo Fail

    If lastFutRow < DATA_ROW Then Exit Function

    WriteFutures_T1_BDP_Efficient wsFut, lastFutRow
    WriteFutureT0BQLFormulas wsFut, lastFutRow
    WriteFuturesCalculatedFormulas_Efficient wsFut, lastFutRow
    WriteFutures_T0_Status wsFut, lastFutRow

    Exit Function

Fail:
    WriteFuturesFormulasSafe = "  - Futures: " & Err.Description & vbCrLf

End Function


Private Function WriteSwapFormulasSafe( _
    ByVal wsSw As Worksheet, _
    ByVal lastSwapRow As Long) As String

    On Error GoTo Fail

    If lastSwapRow < DATA_ROW Then Exit Function

    WriteSwaps_T1_BQL_Formulas wsSw, lastSwapRow
    WriteSwapsT0BQLFormulas wsSw, lastSwapRow
    WriteSwapsCalculatedFormulas wsSw

    Exit Function

Fail:
    WriteSwapFormulasSafe = "  - Swaps: " & Err.Description & vbCrLf

End Function


' =============================================================================
' LOAD OPICS BONDS  (called by Button1_LoadBonds)
' =============================================================================

Public Sub LoadOPICS_Bonds(Optional ByVal silent As Boolean = False)

    Dim wsBnd As Worksheet
    Dim lo As ListObject
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAutoExpand As Boolean

    Set wsBnd = ThisWorkbook.Worksheets(SH_BONDS)

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAutoExpand = Application.AutoCorrect.AutoExpandListRange

    On Error GoTo LoadFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AutoCorrect.AutoExpandListRange = False

    If wsBnd.ListObjects.Count = 0 Then
        Err.Raise vbObjectError + 1000, , _
            "No Excel query table was found on the Bonds sheet."
    End If

    Set lo = wsBnd.ListObjects(1)

    If lo.HeaderRowRange.Row <> 3 Then
        Err.Raise vbObjectError + 1001, , _
            "The Bonds table header must be on row 3."
    End If

    If lo.Range.Column <> 1 Then
        Err.Raise vbObjectError + 1002, , _
            "The Bonds table must start in column A."
    End If

    If lo.ListColumns.Count <> 11 Then
        Err.Raise vbObjectError + 1003, , _
            "The Bonds table must contain exactly 11 columns, from A to K."
    End If

    If Not lo.QueryTable Is Nothing Then
        lo.QueryTable.Refresh BackgroundQuery:=False
    Else
        Err.Raise vbObjectError + 1004, , _
            "The Bonds table is not connected to a refreshable query."
    End If

    If Not silent Then
        MsgBox CStr(lo.ListRows.Count) & _
            " bond positions refreshed.", _
            vbInformation, _
            "Bonds refreshed"
    End If

CleanExit:
    Application.AutoCorrect.AutoExpandListRange = oldAutoExpand
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub

LoadFail:
    MsgBox "Error refreshing the Bonds table:" & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "LoadOPICS_Bonds"

    Resume CleanExit

End Sub

' =============================================================================
' BUTTON 2 - LOAD HEDGES
'
' Primary source:
'   Hedge Risco Tx Juro
'
' OPICS use:
'   Approved missing-field enrichment only
' =============================================================================


Public Sub LoadOPICS_Hedges(Optional ByVal silent As Boolean = False)
    LoadHedges_Step2 silent
End Sub


Public Sub LoadHedges_Step2(Optional ByVal silent As Boolean = False)


    Dim wb As Workbook
    Dim wsCfg As Worksheet
    Dim wsBnd As Worksheet
    Dim wsSw As Worksheet
    Dim wsFut As Worksheet


    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldAskLinks As Boolean


    Dim lastBondRowForLinks As Long
    Dim swapCount As Long
    Dim futCount As Long
    Dim mappedSwapCount As Long
    Dim mappedFutureCount As Long
    Dim totalFutureCount As Long
    Dim totalBookError As String


    Dim isinType As String
    Dim asOf As String
    Dim allPorts As String
    Dim isinList As String


    Dim conn As Object
    Dim rs As Object
    Dim opicsOpened As Boolean


    Dim payFixedExact As Object
    Dim payFixedDates As Object
    Dim payFixedEndDate As Object
    Dim payFixedByCoverage As Object
    
    Dim coverageKey As String
    Dim opicsExactKey As String
    Dim opicsDatesKey As String
    Dim opicsEndDateKey As String


    Dim swapExactKey As String
    Dim swapDatesKey As String
    Dim swapEndDateKey As String


    Dim netPayRaw As String
    Dim payFixedYN As String
    Dim payFixedResult As String
    Dim payFixedStatus As String


    Dim pfRow As Long


    Set wb = ThisWorkbook
    Set wsCfg = wb.Worksheets(SH_CONFIG)
    Set wsBnd = wb.Worksheets(SH_BONDS)
    Set wsSw = wb.Worksheets(SH_SWAPS)
    Set wsFut = wb.Worksheets(SH_FUTURES)


    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldAskLinks = Application.AskToUpdateLinks


    On Error GoTo HedgeErr


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False


    Application.StatusBar = "Step 2: preparing hedge load..."


    Update_Config_Dates_5pmUTC wsCfg


    SetupFuturesFinalHeaders wsFut
    SetupSwapsFinalHeaders wsSw


    If CleanText( _
        wsBnd.Cells( _
            BOND_DATA_ROW, _
            colNum(BCOL_ISIN) _
        ).value _
    ) = "" Then


        MsgBox "No bonds loaded. Run Button 1 first.", vbExclamation
        GoTo CleanExit
    End If


    isinType = CleanText(wsCfg.Range(CFG_ISIN_TYPE).value)


    If isinType = "" Then
        MsgBox _
            "Missing ISIN type in Config " & CFG_ISIN_TYPE & ".", _
            vbExclamation


        GoTo CleanExit
    End If


    asOf = GetAsOfDateString(wsCfg)
    allPorts = "'PORT'"
    isinList = BuildLoadedBondISINList(wsBnd)


    ' -------------------------------------------------------------------------
    ' Clear previous bond-to-hedge links.
    ' -------------------------------------------------------------------------
    lastBondRowForLinks = LastBondDataRow(wsBnd)


    If lastBondRowForLinks >= BOND_DATA_ROW Then


        With wsBnd.Range( _
            BCOL_SWAPLINK & BOND_DATA_ROW & ":" & _
            BCOL_FUTLINK & lastBondRowForLinks _
        )
            .numberFormat = "@"
            .ClearContents
        End With


    End If


    ' =========================================================================
    ' PRIMARY SOURCE: HEDGE RISCO
    ' =========================================================================


    Application.StatusBar = _
        "Step 2: importing Hedge Risco Tx Juro..."


    ClearCoverageSupportSheet
    RefreshCoverageSupportFromHedgeRisco


    Application.StatusBar = _
        "Step 2: building SwapMap..."


    mappedSwapCount = BuildSwapMapFromCoverageSupport()


    Application.StatusBar = _
        "Step 2: building CoverageFutures..."


    mappedFutureCount = _
        BuildCoverageFuturesMapFromCoverageSupport(wsCfg)


    ' The temporary Coverage support copy is no longer needed after both maps
    ' have been built.
    ClearCoverageSupportSheet


    ' =========================================================================
    ' SECOND SOURCE: HEDGE RISCO TOTAL
    '
    ' Appends underneath the Tx Juro rows, so the Futures sheet reads as one
    ' book followed by the other.  A failure here must not lose the Tx Juro
    ' rows that are already mapped: the whole point of loading two files is
    ' that one of them being unreachable is a partial load, not a dead run.
    ' =========================================================================
    Application.StatusBar = _
        "Step 2: importing Hedge Risco Total..."


    totalFutureCount = 0


    On Error Resume Next
    ClearCoverageTotalSupportSheet
    RefreshCoverageTotalSupportFromHedgeRiscoTotal
    totalBookError = Err.Description
    On Error GoTo 0


    If Len(totalBookError) = 0 Then
        Application.StatusBar = _
            "Step 2: appending CoverageFutures from Hedge Risco Total..."
        totalFutureCount = AppendCoverageFuturesMapFromCoverageTotal()
    End If


    ClearCoverageTotalSupportSheet


    mappedFutureCount = mappedFutureCount + totalFutureCount


    ' =========================================================================
    ' BUILD FINAL SWAPS POPULATION FROM SWAPMAP
    ' =========================================================================


    Application.StatusBar = _
        "Step 2: loading swaps from Hedge Risco..."


    wsSw.Range( _
        WCOL_DEALID & DATA_ROW & ":" & _
        WCOL_FLOATFAMILY_STATUS & _
        CStr(SheetClearLastRow(wsSw, DATA_ROW, LastSwapDataRow(wsSw))) _
    ).ClearContents


    swapCount = AppendSwapMapRowsToSwaps(wsSw, 0)


    ' =========================================================================
    ' BUILD FINAL FUTURES POPULATION FROM COVERAGEFUTURES
    ' =========================================================================


    Application.StatusBar = _
        "Step 2: loading futures from Hedge Risco..."


    futCount = AppendCoverageFuturesRowsToFutures(wsFut, wsBnd)


    ' =========================================================================
    ' OPICS ENRICHMENT
    '
    ' Hedge Risco has already determined:
    '   - which swaps and futures exist;
    '   - their LinkedISIN;
    '   - their main economics;
    '   - their classification.
    '
    ' OPICS is now used only to enrich approved missing fields.
    ' =========================================================================


    Application.StatusBar = _
        "Step 2: enriching missing operational fields from OPICS..."


    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")


    opicsOpened = OpenOPICS(conn, wsCfg)


    If opicsOpened Then


        ' ---------------------------------------------------------------------
        ' SWAP ENRICHMENT
        '
        ' Temporary transitional rule:
        '   Use OPICS only to retrieve PayFixed.
        '
        ' No OPICS swap rows are written into the Swaps sheet.
        ' ---------------------------------------------------------------------
        Set payFixedExact = CreateObject("Scripting.Dictionary")
        Set payFixedDates = CreateObject("Scripting.Dictionary")
        Set payFixedEndDate = CreateObject("Scripting.Dictionary")
        Set payFixedByCoverage = CreateObject("Scripting.Dictionary")


        rs.Open _
            GetSwapSelectQuery( _
                allPorts, _
                isinList, _
                isinType, _
                asOf _
            ), _
            conn, _
            0, _
            1


        Do While Not rs.EOF


            netPayRaw = UCase$(Trim$( _
                NullToStr(Fld(rs, "PayFixed")) _
            ))


            payFixedYN = ""


            Select Case netPayRaw


                Case "Y", "1", "PAY", "PAYFIXED", "PAYER"
                    payFixedYN = "Y"


                Case "R", "N", "0", "REC", "RECEIVE", "RECEIVER"
                    payFixedYN = "N"


                Case Else
                    If Len(netPayRaw) > 0 Then
                        If Left$(netPayRaw, 1) = "P" Then
                            payFixedYN = "Y"
                        End If
                    End If


            End Select


            If Len(payFixedYN) > 0 And Len(SwapMatchText(Fld(rs, "LinkedISIN"))) > 0 Then


                opicsExactKey = SwapPayFixedKeyExact( _
                    Fld(rs, "LinkedISIN"), _
                    Fld(rs, "StartDate"), _
                    Fld(rs, "EndDate"), _
                    Fld(rs, "CptyName") _
                )


                opicsDatesKey = SwapPayFixedKeyDates( _
                    Fld(rs, "LinkedISIN"), _
                    Fld(rs, "StartDate"), _
                    Fld(rs, "EndDate") _
                )


                opicsEndDateKey = SwapPayFixedKeyEndDate( _
                    Fld(rs, "LinkedISIN"), _
                    Fld(rs, "EndDate") _
                )


                AddPayFixedCandidate _
                    payFixedExact, _
                    opicsExactKey, _
                    payFixedYN


                AddPayFixedCandidate _
                    payFixedDates, _
                    opicsDatesKey, _
                    payFixedYN


                AddPayFixedCandidate _
                    payFixedEndDate, _
                    opicsEndDateKey, _
                    payFixedYN


            End If


            rs.MoveNext


        Loop


        rs.Close


       ' ---------------------------------------------------------------------
       ' FIRST PASS
       '
       ' Match only PLAIN swaps against OPICS.
       ' Store each valid PayFixed result by Hedge Risco coverage group.
       ' ---------------------------------------------------------------------


        wsSw.Range( _
            WCOL_PAYFIXED & DATA_ROW & ":" & _
            WCOL_PAYFIXED & CStr(DATA_ROW + swapCount - 1) _
        ).numberFormat = "@"


        For pfRow = DATA_ROW To DATA_ROW + swapCount - 1


            If UCase$(CleanText( _
                wsSw.Cells( _
                    pfRow, _
                    colNum(WCOL_SWAP_ID_SOURCE) _
                ).value _
            )) = "PLAIN" Then


                payFixedResult = ""
                payFixedStatus = ""


                swapExactKey = SwapPayFixedKeyExact( _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_LINKEDISIN) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_STARTDATE) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_ENDDATE) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_CPTY) _
                    ).value _
                )


                swapDatesKey = SwapPayFixedKeyDates( _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_LINKEDISIN) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_STARTDATE) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_ENDDATE) _
                    ).value _
                )


                swapEndDateKey = SwapPayFixedKeyEndDate( _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_LINKEDISIN) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_ENDDATE) _
                    ).value _
                )


                If payFixedExact.Exists(swapExactKey) Then


                    If payFixedExact(swapExactKey) <> "AMBIGUOUS" Then
                        payFixedResult = payFixedExact(swapExactKey)
                            payFixedStatus = "PayFixed OPICS exact match"
                    Else
                        payFixedStatus = "PayFixed ambiguous exact match"
                    End If


                End If


                If Len(payFixedResult) = 0 And Len(payFixedStatus) = 0 Then


                    If payFixedDates.Exists(swapDatesKey) Then


                        If payFixedDates(swapDatesKey) <> "AMBIGUOUS" Then
                            payFixedResult = payFixedDates(swapDatesKey)
                            payFixedStatus = "PayFixed OPICS date match"
                        Else
                            payFixedStatus = "PayFixed ambiguous date match"
                        End If


                    End If


                End If


                If Len(payFixedResult) = 0 And _
                    Len(payFixedStatus) = 0 Then


                    If payFixedEndDate.Exists(swapEndDateKey) Then


                        If payFixedEndDate(swapEndDateKey) <> "AMBIGUOUS" Then
                            payFixedResult = payFixedEndDate(swapEndDateKey)
                            payFixedStatus = "PayFixed OPICS maturity match"
                        Else
                            payFixedStatus = "PayFixed ambiguous maturity match"
                        End If


                    End If


                End If


                If Len(payFixedResult) = 0 And _
                    Len(payFixedStatus) = 0 Then


                    payFixedStatus = "PayFixed OPICS no match"


                End If


                wsSw.Cells( _
                    pfRow, _
                    colNum(WCOL_PAYFIXED) _
                ).value = payFixedResult


                AppendSwapEnrichmentStatus _
                    wsSw, _
                    pfRow, _
                    payFixedStatus


                ' -------------------------------------------------------------
                ' Store a valid PLAIN result by its Hedge Risco coverage group.
                '
                ' The group uses:
                '   SwapMap source row
                '   Coverage relationship
                '   LinkedISIN
                ' -------------------------------------------------------------
                If Len(payFixedResult) > 0 Then


                    coverageKey = SwapCoverageGroupKey( _
                        wsSw.Cells( _
                            pfRow, _
                            colNum(WCOL_SWAPMAP_SOURCEROW) _
                        ).value, _
                        wsSw.Cells( _
                            pfRow, _
                            colNum(WCOL_COVERAGERELATION) _
                        ).value, _
                        wsSw.Cells( _
                            pfRow, _
                            colNum(WCOL_LINKEDISIN) _
                        ).value _
                    )


                    If Len(coverageKey) > 0 Then


                        If Not payFixedByCoverage.Exists(coverageKey) Then


                            payFixedByCoverage.Add _
                                coverageKey, _
                                payFixedResult


                        ElseIf payFixedByCoverage(coverageKey) <> payFixedResult Then


                            payFixedByCoverage(coverageKey) = "AMBIGUOUS"


                        End If


                    End If


                End If


            End If


        Next pfRow
        
        ' ---------------------------------------------------------------------
        ' SECOND PASS
        '
        ' Copy the PLAIN PayFixed result to the SYNTHETIC row belonging to
        ' the same Hedge Risco coverage group.
        ' ---------------------------------------------------------------------


        For pfRow = DATA_ROW To DATA_ROW + swapCount - 1


            If UCase$(CleanText( _
                wsSw.Cells( _
                    pfRow, _
                    colNum(WCOL_SWAP_ID_SOURCE) _
                ).value _
            )) = "SYNTHETIC" Then


                coverageKey = SwapCoverageGroupKey( _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_SWAPMAP_SOURCEROW) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_COVERAGERELATION) _
                    ).value, _
                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_LINKEDISIN) _
                    ).value _
                )


                If payFixedByCoverage.Exists(coverageKey) Then


                    If payFixedByCoverage(coverageKey) <> "AMBIGUOUS" Then


                        wsSw.Cells( _
                            pfRow, _
                            colNum(WCOL_PAYFIXED) _
                        ).value = payFixedByCoverage(coverageKey)


                        AppendSwapEnrichmentStatus _
                            wsSw, _
                            pfRow, _
                            "PayFixed inherited from PLAIN coverage relationship"


                   Else


                        wsSw.Cells( _
                            pfRow, _
                            colNum(WCOL_PAYFIXED) _
                        ).ClearContents


                        AppendSwapEnrichmentStatus _
                            wsSw, _
                            pfRow, _
                            "PayFixed ambiguous in PLAIN coverage relationship"


                    End If


                Else


                    wsSw.Cells( _
                        pfRow, _
                        colNum(WCOL_PAYFIXED) _
                    ).ClearContents


                    AppendSwapEnrichmentStatus _
                        wsSw, _
                        pfRow, _
                        "PayFixed missing from PLAIN coverage relationship"


                End If


            End If


        Next pfRow
        
        


        ' ---------------------------------------------------------------------
        ' FUTURES ENRICHMENT
        '
        ' Enrich only the final Hedge Risco futures rows.
        ' ---------------------------------------------------------------------
        If futCount > 0 Then


            EnrichCoverageFuturesFromOPICS _
                conn, _
                wsFut, _
                wsCfg, _
                isinType, _
                asOf


        End If


        conn.Close


    End If


    ' =========================================================================
    ' FORMULA RESPONSIBILITY
    '
    ' Do not call WriteSwapsCalculatedFormulas here.
    '
    ' Button 3 writes Bloomberg and market-dependent formulas.
    ' Button 5 later rebuilds the final PNL attribution.
    ' =========================================================================


    wsCfg.Range(CFG_SWAP_CNT).value = swapCount
    wsCfg.Range(CFG_FUT_CNT).value = futCount


    Application.StatusBar = False


    ' The two coverage books are reported separately, not as one total.  A run
    ' where Hedge Risco Total was unreachable still succeeds and still loads
    ' every Tx Juro position - so the only way to notice is a count that says
    ' so, in the box the desk actually reads.
    If silent Then GoTo CleanExit

    MsgBox _
        CStr(swapCount) & _
        " swap rows loaded from Hedge Risco Tx Juro." & vbCrLf & _
        CStr(futCount) & _
        " futures rows loaded (both coverage books)." & vbCrLf & vbCrLf & _
        "SwapMap rows: " & CStr(mappedSwapCount) & vbCrLf & _
        "CoverageFutures rows: " & CStr(mappedFutureCount) & vbCrLf & _
        "  of which Hedge Risco Total: " & CStr(totalFutureCount) & vbCrLf & _
        IIf( _
            Len(totalBookError) > 0, _
            "  Hedge Risco Total NOT loaded: " & totalBookError & vbCrLf, _
            "" _
        ) & vbCrLf & _
        IIf( _
            opicsOpened, _
            "OPICS was used only to enrich approved missing fields.", _
            "OPICS enrichment was not completed." _
        ), _
        IIf(opicsOpened And Len(totalBookError) = 0, vbInformation, vbExclamation), _
        "Step 2 Complete"


CleanExit:


    On Error Resume Next


    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If


    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If


    ClearCoverageSupportSheet


    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.AskToUpdateLinks = oldAskLinks
    Application.StatusBar = False


    On Error GoTo 0
    Exit Sub


HedgeErr:


    Dim eInfo As String
    eInfo = CaptureErrInfo()


    On Error Resume Next


    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If


    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If


    ClearCoverageSupportSheet


    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.AskToUpdateLinks = oldAskLinks
    Application.StatusBar = False


    On Error GoTo 0


    MsgBox _
        "Step 2 hedge load failed." & vbCrLf & vbCrLf & _
        eInfo, _
        vbCritical, _
        "Load Hedges Failed"


End Sub






Private Sub ValidateCurrentCurveTickers(ByVal ws As Worksheet)


    ValidateOneCurveTickerRange _
        ws, _
        "C7:C19", _
        "EUR ESTR OIS"


    ValidateOneCurveTickerRange _
        ws, _
        "F7:F19", _
        "EUR Government"


    ValidateOneCurveTickerRange _
        ws, _
        "I7:I19", _
        "EUR EURIBOR Swap"


    ValidateOneCurveTickerRange _
        ws, _
        "T7:T19", _
        "USD OIS"


    ValidateOneCurveTickerRange _
        ws, _
        "W7:W19", _
        "USD Government"


    ValidateOneCurveTickerRange _
        ws, _
        "Z7:Z19", _
        "USD Swap"


    ValidateOneCurveTickerRange _
        ws, _
        "AK7:AK19", _
        "GBP OIS"


    ValidateOneCurveTickerRange _
        ws, _
        "AN7:AN19", _
        "GBP Government"


    ValidateOneCurveTickerRange _
        ws, _
        "AQ7:AQ19", _
        "GBP Swap"


End Sub


Private Function ValidateCurrentCurvePair( _
    ByVal ws As Worksheet, _
    ByVal tickerCol As String, _
    ByVal outputCol As String) As Boolean


    Dim rowNum As Long
    Dim tickerText As String
    Dim outputCell As Range


    For rowNum = CURVE_FIRST_ROW To CURVE_LAST_ROW


        tickerText = CleanText( _
            ws.Range(tickerCol & rowNum).value)


        Set outputCell = ws.Range(outputCol & rowNum)


        If Len(tickerText) > 0 Then


            If CurveOutputCellHasProblem(outputCell) Then


                Debug.Print _
                    "Curve problem at " & _
                    ws.Name & "!" & _
                    outputCell.Address(False, False) & _
                    ", ticker " & tickerText & _
                    ", result " & CStr(outputCell.Text)


                ValidateCurrentCurvePair = True
                Exit Function


            End If


        End If


    Next rowNum


    ValidateCurrentCurvePair = False


End Function






Private Sub ValidateOneCurveTickerRange( _
    ByVal ws As Worksheet, _
    ByVal rangeAddress As String, _
    ByVal curveName As String)


    Dim c As Range
    Dim populatedCount As Long


    For Each c In ws.Range(rangeAddress).Cells


        If Len(CleanText(c.value)) > 0 Then
            populatedCount = populatedCount + 1
        End If


    Next c


    If populatedCount < 2 Then


        Err.Raise _
            9623, _
            "ValidateOneCurveTickerRange", _
            curveName & _
            " has fewer than two populated Bloomberg tickers in " & _
            ws.Name & "!" & rangeAddress & "."


    End If


End Sub


Private Sub ValidateCurveTickerRange( _
    ByVal ws As Worksheet, _
    ByVal rangeAddress As String, _
    ByRef missingCells As String)


    Dim c As Range
    Dim tickerText As String


    For Each c In ws.Range(rangeAddress).Cells


        tickerText = CleanText(c.value)


        If Len(tickerText) = 0 Then


            If Len(missingCells) > 0 Then
                missingCells = missingCells & ", "
            End If


            missingCells = missingCells & c.Address(False, False)


        End If


    Next c


End Sub




















Private Sub TriggerBloombergRefreshOnly()
    If KeepFormulasMode() Then Exit Sub   ' offline/inspection: no Bloomberg terminal
    On Error Resume Next
    Application.Run "RefreshAllStaticData"
    Application.Run "BLP.RefreshAll"
    On Error GoTo 0
End Sub


' =============================================================================
' EFFICIENT CURVE WRITERS
' =============================================================================


Private Sub WriteCurves_T1_BDP_Efficient( _
    ByVal ws As Worksheet)


    mLastFormulaTarget = ws.Name & "!E7:E19"
    WriteBDPDown_Efficient _
        ws, _
        "C", _
        "E", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!H7:H19"
    WriteBDPDown_Efficient _
        ws, _
        "F", _
        "H", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_YLD_YTM_MID


    mLastFormulaTarget = ws.Name & "!K7:K19"
    WriteBDPDown_Efficient _
        ws, _
        "I", _
        "K", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!H7"
    WriteBDPCell_Efficient _
        ws, _
        "F", _
        "H", _
        CURVE_FIRST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!V7:V19"
    WriteBDPDown_Efficient _
        ws, _
        "T", _
        "V", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!Y7:Y19"
    WriteBDPDown_Efficient _
        ws, _
        "W", _
        "Y", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!AB7:AB19"
    WriteBDPDown_Efficient _
        ws, _
        "Z", _
        "AB", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!AM7:AM19"
    WriteBDPDown_Efficient _
        ws, _
        "AK", _
        "AM", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!AP7:AP19"
    WriteBDPDown_Efficient _
        ws, _
        "AN", _
        "AP", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_YLD_YTM_MID


    mLastFormulaTarget = ws.Name & "!AS7:AS19"
    WriteBDPDown_Efficient _
        ws, _
        "AQ", _
        "AS", _
        CURVE_FIRST_ROW, _
        CURVE_LAST_ROW, _
        BBG_PX_LAST


    mLastFormulaTarget = ws.Name & "!AP7"
    WriteBDPCell_Efficient _
        ws, _
        "AN", _
        "AP", _
        CURVE_FIRST_ROW, _
        BBG_PX_LAST


    WriteCurveDerivedFormulas_Efficient ws


    mLastFormulaTarget = ""


End Sub




Private Sub WriteBDPDown_Efficient( _
    ByVal ws As Worksheet, _
    ByVal tickerCol As String, _
    ByVal outCol As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal fieldName As String)


    Dim tickerColNum As Long
    Dim formulaText As String


    tickerColNum = ws.Columns(tickerCol).Column


    formulaText = _
        "=IF(" & _
            "RC" & tickerColNum & "="""",""""," & _
            "BDP(" & _
                "RC" & tickerColNum & "," & _
                XlText(fieldName) & _
            ")" & _
        ")"


    mLastFormulaTarget = _
        ws.Name & "!" & _
        outCol & firstRow & ":" & _
        outCol & lastRow


    ws.Range( _
        outCol & firstRow & ":" & _
        outCol & lastRow _
    ).FormulaR1C1 = formulaText


End Sub


Private Sub WriteBDPCell_Efficient( _
    ByVal ws As Worksheet, _
    ByVal tickerCol As String, _
    ByVal outCol As String, _
    ByVal rowNum As Long, _
    ByVal fieldName As String)


    Dim tickerColNum As Long
    Dim formulaText As String


    tickerColNum = ws.Columns(tickerCol).Column


    formulaText = _
        "=IF(" & _
            "RC" & tickerColNum & "="""",""""," & _
            "BDP(" & _
                "RC" & tickerColNum & "," & _
                XlText(fieldName) & _
            ")" & _
        ")"


    mLastFormulaTarget = _
        ws.Name & "!" & _
        outCol & rowNum


    ws.Range( _
        outCol & rowNum _
    ).FormulaR1C1 = formulaText


End Sub










Private Function FXT1FormulaR1C1(ByVal cfgFxFix As String) As String
    Dim fml As String


    fml = "=IF(" & RC(BCOL_ISIN) & "="""","""",LET("
    fml = fml & "_ccy,UPPER(TRIM(" & RC(BCOL_CCY) & ")),"
    fml = fml & "IF(_ccy="""","""","
    fml = fml & "IF(_ccy=""EUR"",1,"
    fml = fml & "IFERROR("
    fml = fml & "1/BDP(""EUR""&_ccy&"" Curncy"",""PX_LAST""),"
    fml = fml & "IFERROR("
    fml = fml & "BDP(_ccy&""EUR Curncy"",""PX_LAST""),"
    fml = fml & """"")"
    fml = fml & ")"
    fml = fml & ")))"
    fml = fml & ")"




    FXT1FormulaR1C1 = fml
End Function


Private Function FXT0FormulaR1C1(ByVal cfgFxFix As String, ByVal cfgT0 As String) As String
    Dim secEURCCY As String
    Dim secCCYEUR As String
    Dim fml As String


    secEURCCY = """EUR""&_ccy&"" Curncy"""
    secCCYEUR = "_ccy&""EUR Curncy"""


    fml = "=IF(" & RC(BCOL_ISIN) & "="""","""",LET("
    fml = fml & "_ccy,UPPER(TRIM(" & RC(BCOL_CCY) & ")),"
    fml = fml & "IF(_ccy="""","""","
    fml = fml & "IF(_ccy=""EUR"",1,"
    fml = fml & "IFERROR("
    fml = fml & "1/" & BQLPointExprR1C1(secEURCCY, XlText(BBG_PX_LAST), cfgT0) & ","
    fml = fml & "IFERROR("
    fml = fml & BQLPointExprR1C1(secCCYEUR, XlText(BBG_PX_LAST), cfgT0) & ","
    fml = fml & """"")"
    fml = fml & ")"
    fml = fml & ")))"
    fml = fml & ")"


    FXT0FormulaR1C1 = fml
End Function


Private Function PnlGetOrCreateSheet(ByVal sheetName As String) As Worksheet


    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If


    Set PnlGetOrCreateSheet = ws


End Function


Private Function CurrentCurveResultHasProblem( _
    ByVal ws As Worksheet) As Boolean


    If ValidateCurrentCurvePair(ws, "C", "E") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "F", "H") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "I", "K") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "T", "V") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "W", "Y") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "Z", "AB") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "AK", "AM") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "AN", "AP") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    If ValidateCurrentCurvePair(ws, "AQ", "AS") Then
        CurrentCurveResultHasProblem = True
        Exit Function
    End If


    CurrentCurveResultHasProblem = False


End Function


Private Function CurveOutputCellHasProblem( _
    ByVal c As Range) As Boolean


    Dim txt As String


    txt = Trim$(CStr(c.Text))


    If IsError(c.value) Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If InStr(1, txt, "Requesting", vbTextCompare) > 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If InStr(1, txt, "#N/A", vbTextCompare) > 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If InStr(1, txt, "Connection", vbTextCompare) > 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If InStr(1, txt, "Authorization", vbTextCompare) > 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If InStr(1, txt, "Invalid", vbTextCompare) > 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If Len(txt) = 0 Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    If Not IsNumeric(c.value) Then


        CurveOutputCellHasProblem = True
        Exit Function


    End If


    CurveOutputCellHasProblem = False


End Function














' =============================================================================
' BLOOMBERG SCALABLE SECURITY-CANDIDATE FORMULA HELPERS
'
' Logic:
'   1) Candidate security strings are written in Bonds!BV:CE.
'   2) AP resolves the first parseable Bloomberg security.
'   3) BDP fields try AP first, then selected candidate columns from BV:CE.
' =============================================================================


Private Sub WriteBondSecurityCandidates(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < BOND_DATA_ROW Then Exit Sub


    Dim f As Long
    f = BOND_DATA_ROW


    ws.Range(BCOL_BBG_CAND_ISIN & f & ":" & BCOL_BBG_CAND_ISIN & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&"" ISIN""")


    ws.Range(BCOL_BBG_CAND_SLASHISIN & f & ":" & BCOL_BBG_CAND_SLASHISIN & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), """/isin/""&" & RC(BCOL_ISIN))


    ws.Range(BCOL_BBG_CAND_CORP & f & ":" & BCOL_BBG_CAND_CORP & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&"" Corp""")


    ws.Range(BCOL_BBG_CAND_BVAL_CORP & f & ":" & BCOL_BBG_CAND_BVAL_CORP & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&""@BVAL Corp""")


    ws.Range(BCOL_BBG_CAND_BGN_CORP & f & ":" & BCOL_BBG_CAND_BGN_CORP & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&""@BGN Corp""")


    ws.Range(BCOL_BBG_CAND_GOVT & f & ":" & BCOL_BBG_CAND_GOVT & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&"" Govt""")


    ws.Range(BCOL_BBG_CAND_BVAL_GOVT & f & ":" & BCOL_BBG_CAND_BVAL_GOVT & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&""@BVAL Govt""")


    ws.Range(BCOL_BBG_CAND_BGN_GOVT & f & ":" & BCOL_BBG_CAND_BGN_GOVT & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&""@BGN Govt""")


    ws.Range(BCOL_BBG_CAND_MTGE & f & ":" & BCOL_BBG_CAND_MTGE & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&"" Mtge""")


    ws.Range(BCOL_BBG_CAND_MMKT & f & ":" & BCOL_BBG_CAND_MMKT & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_ISIN) & "&"" M-Mkt""")


End Sub


Private Function BBGParseFormulaR1C1() As String
    Dim expr As String
    Dim i As Long


    expr = XlText("UNKNOWN")


    For i = COL_BBG_CAND_LAST To COL_BBG_CAND_FIRST Step -1
        expr = "IFERROR(BDP(RC" & i & ",""PARSEKYABLE_DES"")," & expr & ")"
    Next i


    BBGParseFormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), expr)
End Function


' Guarded single-BDP fallback step.  Replaces a bare IFERROR(BDP(...),next).
'
' Why this exists: Bloomberg returns "#N/A Invalid Field" and "#N/A Authorization"
' as TEXT strings, not as real Excel error values, so IFERROR does NOT trap them - the
' error text "passes" as a valid result, short-circuits the fallback chain, and lands in
' the cell (this is the AY/AZ/BB/BC/BD bug).  This wrapper calls BDP once via LET and
' falls through to <nextExpr> when the result is unusable.
'
' Pending-safe: "#N/A Requesting Data" is ALSO returned as text and MUST keep bubbling up
' so RangeHasPendingBBG / WaitForBloombergReadyMany can still see the request in flight.
' The guard therefore falls through only for "#N..." text that is NOT a Requesting
' message (and for genuine non-text error values), leaving pending markers intact.
Private Function GuardedBDPExprR1C1(ByVal secRef As String, ByVal fieldExpr As String, ByVal nextExpr As String) As String
    GuardedBDPExprR1C1 = _
        "LET(_v,BDP(" & secRef & "," & fieldExpr & ")," & _
        "IF(IF(ISTEXT(_v),AND(LEFT(_v,2)=""#N"",NOT(ISNUMBER(SEARCH(""Requesting"",_v)))),ISERR(_v))," & _
        nextExpr & ",_v))"
End Function


Private Function GuardedBDPExprR1C1_WithOneOverride( _
    ByVal secRef As String, _
    ByVal fieldExpr As String, _
    ByVal overrideNameExpr As String, _
    ByVal overrideValueExpr As String, _
    ByVal nextExpr As String) As String


    GuardedBDPExprR1C1_WithOneOverride = _
        "LET(_v,BDP(" & secRef & "," & fieldExpr & "," & overrideNameExpr & "," & overrideValueExpr & ")," & _
        "IF(IF(ISTEXT(_v),AND(LEFT(_v,2)=""#N"",NOT(ISNUMBER(SEARCH(""Requesting"",_v)))),ISERR(_v))," & _
        nextExpr & ",_v))"


End Function


Private Function BBGTryBDPFieldExprR1C1(ByVal fieldName As String, ByVal fallbackExpr As String) As String
    Dim expr As String
    Dim i As Long


    expr = fallbackExpr


    For i = COL_BBG_CAND_LAST To COL_BBG_CAND_FIRST Step -1
        expr = GuardedBDPExprR1C1("RC" & i, XlText(fieldName), expr)
    Next i


    expr = GuardedBDPExprR1C1("RC" & COL_BBG_TICKER, XlText(fieldName), expr)


    BBGTryBDPFieldExprR1C1 = expr
End Function


' Adaptive variant used by MULTI-FIELD formulas.  A multi-field formula expands
' per field AND per candidate, so the full 11-candidate chain overflows Excel's hard
' 64-level function-nesting cap and makes .FormulaR1C1 raise run-time error 1004.
' BondFallbackColCount below works out how much of the chain fits for the field count
' actually being asked for, instead of truncating every column to the same three.
'
' The security ids one Bloomberg field is tried against, in order.
'
' The desk starts a BDP lookup with the ISIN-suffixed form - `XS123... ISIN` -
' so that leads the chain, behind only the ticker PARSEKYABLE_DES already
' resolved into BCOL_BBG_TICKER.  The rest follow in the order the module header
' states: /isin/, then the Corp / Govt / Mtge / M-Mkt venue forms.
'
' Built from the column constants.  It used to be Array(COL_BBG_TICKER, 75, 77)
' - two bare RC numbers, with a comment that misdescribed all three of them: it
' called RC42 the ticker (RC42 is Ticker_Status; the ticker is RC41), RC75 the
' /isin/ candidate (RC75 is @BVAL Corp; /isin/ is RC73) and RC77 @BVAL Corp
' (RC77 is Govt).  So the chain was ticker -> @BVAL Corp -> Govt, and NEITHER
' ISIN form was in it, while the module header said the resolver tries /isin/
' first.  Moving a candidate column would not have moved the pull either.
Private Function BondSecurityFallbackCols() As Variant

    ' ColIdx, not colNum: this is pure arithmetic on a letter and must not go
    ' through the object model.  It runs once per field per column, and it has
    ' to be executable away from a worksheet to be tested.
    BondSecurityFallbackCols = Array( _
        COL_BBG_TICKER, _
        ColIdx(BCOL_BBG_CAND_ISIN), _
        ColIdx(BCOL_BBG_CAND_SLASHISIN), _
        ColIdx(BCOL_BBG_CAND_CORP), _
        ColIdx(BCOL_BBG_CAND_BVAL_CORP), _
        ColIdx(BCOL_BBG_CAND_BGN_CORP), _
        ColIdx(BCOL_BBG_CAND_GOVT), _
        ColIdx(BCOL_BBG_CAND_BVAL_GOVT), _
        ColIdx(BCOL_BBG_CAND_BGN_GOVT), _
        ColIdx(BCOL_BBG_CAND_MTGE), _
        ColIdx(BCOL_BBG_CAND_MMKT))

End Function


' How many of them fit, given how many FIELDS the column is trying.
'
' Excel stops at 64 levels of nested functions.  Every (candidate x field) pair
' costs two levels - the LET and the IF that chooses between this answer and the
' next - and the surrounding wrapper costs six, so:
'
'       nesting = 6 + 2 * candidates * fields
'
' The old code answered this by hard-coding THREE candidates for every column
' whatever it was asking for, which was over-cautious for the one-field columns
' and arbitrary for the rest.  Deriving it lets a one- or two-field column try
' the whole chain and only shortens it where the field list is genuinely long.
'
' Capped at 60 rather than 64 so a future field added to a list cannot silently
' push a formula over the limit.  The constant itself lives with the other
' module constants at the top of the file; see BBG_MAX_FORMULA_NESTING.
Private Function BondFallbackColCount(ByVal fieldCount As Long) As Long

    Dim cols As Long
    Dim available As Long

    If fieldCount < 1 Then fieldCount = 1

    available = UBound(BondSecurityFallbackCols()) + 1
    cols = (BBG_MAX_FORMULA_NESTING - 6) \ (2 * fieldCount)

    If cols < 1 Then cols = 1
    If cols > available Then cols = available

    BondFallbackColCount = cols

End Function


Private Function BBGTryBDPFieldExprR1C1_Short( _
    ByVal fieldName As String, _
    ByVal fallbackExpr As String, _
    ByVal fieldCount As Long) As String

    Dim allCols As Variant
    Dim expr As String
    Dim k As Long
    Dim useCols As Long

    allCols = BondSecurityFallbackCols()
    useCols = BondFallbackColCount(fieldCount)

    expr = fallbackExpr

    For k = useCols - 1 To 0 Step -1
        expr = GuardedBDPExprR1C1("RC" & allCols(k), XlText(fieldName), expr)
    Next k

    BBGTryBDPFieldExprR1C1_Short = expr

End Function


Private Function BBGFirstBDPFormulaR1C1(ByVal fieldName As String, Optional ByVal fallbackExpr As String = "") As String
    If fallbackExpr = "" Then fallbackExpr = XlBlank()


    BBGFirstBDPFormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), BBGTryBDPFieldExprR1C1(fieldName, fallbackExpr))
End Function


Private Function BBGFirstBDPMultiFieldFormulaR1C1(ByVal fields As Variant, Optional ByVal fallbackExpr As String = "") As String
    Dim expr As String
    Dim i As Long


    If fallbackExpr = "" Then fallbackExpr = XlBlank()
    expr = fallbackExpr


    ' Use the compact (3-security) per-field expansion so nesting stays well under
    ' Excel's 64-level limit even for 6-field columns like AY.
    For i = UBound(fields) To LBound(fields) Step -1
        expr = BBGTryBDPFieldExprR1C1_Short( _
            CStr(fields(i)), expr, UBound(fields) - LBound(fields) + 1)
    Next i


    BBGFirstBDPMultiFieldFormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), expr)
End Function


Private Function BBGAccruedInterestFormulaR1C1(ByVal cfgT1 As String) As String


    Dim expr As String
    Dim settleExpr As String
    Dim i As Long


    ' Bloomberg SETTLE_DT expects a real Excel date value, not TEXT(...).
    ' INT() strips any time component from Config!B5 while preserving the Excel date serial.
    settleExpr = "INT(" & cfgT1 & ")"


    ' Final fallback:
    '   V = AccruedInterest_T0
    '   T = RC20 = CleanPx_T0
    '   U = RC21 = DirtyPx_T0
    '
    ' If Bloomberg accrued interest is unavailable, use Dirty - Clean.
    expr = "IF(AND(ISNUMBER(" & RC(BCOL_DIRTYPX_T0) & "),ISNUMBER(" & RC(BCOL_CLEANPX_T0) & "))," & RC(BCOL_DIRTYPX_T0) & "-" & RC(BCOL_CLEANPX_T0) & "," & XlBlank() & ")"


    For i = COL_BBG_CAND_LAST To COL_BBG_CAND_FIRST Step -1
        expr = GuardedBDPExprR1C1_WithOneOverride( _
            "RC" & i, _
            XlText("ACCRUED_INTEREST"), _
            XlText("SETTLE_DT"), _
            settleExpr, _
            expr)
    Next i


    expr = GuardedBDPExprR1C1_WithOneOverride( _
        "RC" & COL_BBG_TICKER, _
        XlText("ACCRUED_INTEREST"), _
        XlText("SETTLE_DT"), _
        settleExpr, _
        expr)


    BBGAccruedInterestFormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), expr)


End Function
Private Sub WriteCurveDerivedFormulas_Efficient(ws As Worksheet)
    ws.Range(CVCOL_EUR_G_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_EUR_G_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_EUR_GOV_TM1), RC(CVCOL_EUR_ESTR_OIS_TM1))
    ws.Range(CVCOL_EUR_G_T0 & CURVE_FIRST_ROW & ":" & CVCOL_EUR_G_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_EUR_GOV_T0), RC(CVCOL_EUR_ESTR_OIS_T0))
    ws.Range(CVCOL_EUR_Q_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_EUR_Q_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_EUR_EURIBOR_SWAP_TM1), RC(CVCOL_EUR_GOV_TM1))
    ws.Range(CVCOL_EUR_Q_T0 & CURVE_FIRST_ROW & ":" & CVCOL_EUR_Q_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_EUR_EURIBOR_SWAP_T0), RC(CVCOL_EUR_GOV_T0))
    ws.Range(CVCOL_EUR_STATUS & CURVE_FIRST_ROW & ":" & CVCOL_EUR_STATUS & CURVE_LAST_ROW).FormulaR1C1 = "=IF(COUNT(" & RC(CVCOL_EUR_ESTR_OIS_TM1) & ":" & RC(CVCOL_EUR_Q_T0) & ")=0,""Missing"",""OK"")"
    ws.Range(CVCOL_USD_G_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_USD_G_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_USD_GOV_TM1), RC(CVCOL_USD_OIS_TM1))
    ws.Range(CVCOL_USD_G_T0 & CURVE_FIRST_ROW & ":" & CVCOL_USD_G_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_USD_GOV_T0), RC(CVCOL_USD_OIS_T0))
    ws.Range(CVCOL_USD_Q_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_USD_Q_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_USD_SWAP_TM1), RC(CVCOL_USD_GOV_TM1))
    ws.Range(CVCOL_USD_Q_T0 & CURVE_FIRST_ROW & ":" & CVCOL_USD_Q_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_USD_SWAP_T0), RC(CVCOL_USD_GOV_T0))
    ws.Range(CVCOL_USD_STATUS & CURVE_FIRST_ROW & ":" & CVCOL_USD_STATUS & CURVE_LAST_ROW).FormulaR1C1 = "=IF(" & RC(CVCOL_USD_GOV_TICKER) & "=""""," & """USD Gov node unavailable - skipped by interpolation""," & "IF(OR(" & "NOT(ISNUMBER(" & RC(CVCOL_USD_GOV_T0) & "))," & "NOT(ISNUMBER(" & RC(CVCOL_USD_GOV_TM1) & "))" & ")," & """Missing USD Gov data""," & """OK""" & ")" & ")"
    ws.Range(CVCOL_GBP_G_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_GBP_G_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_GBP_GOV_TM1), RC(CVCOL_GBP_OIS_TM1))
    ws.Range(CVCOL_GBP_G_T0 & CURVE_FIRST_ROW & ":" & CVCOL_GBP_G_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_GBP_GOV_T0), RC(CVCOL_GBP_OIS_T0))
    ws.Range(CVCOL_GBP_Q_TM1 & CURVE_FIRST_ROW & ":" & CVCOL_GBP_Q_TM1 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_GBP_SWAP_TM1), RC(CVCOL_GBP_GOV_TM1))
    ws.Range(CVCOL_GBP_Q_T0 & CURVE_FIRST_ROW & ":" & CVCOL_GBP_Q_T0 & CURVE_LAST_ROW).FormulaR1C1 = DiffFormula(RC(CVCOL_GBP_SWAP_T0), RC(CVCOL_GBP_GOV_T0))
    ws.Range(CVCOL_GBP_STATUS & CURVE_FIRST_ROW & ":" & CVCOL_GBP_STATUS & CURVE_LAST_ROW).FormulaR1C1 = "=IF(COUNT(" & RC(CVCOL_GBP_OIS_TM1) & ":" & RC(CVCOL_GBP_Q_T0) & ")=0,""Missing"",""OK"")"
End Sub




' Tracked formula write: records the target range in mLastFormulaTarget before the
' assignment so an over-nested / invalid formula (Excel run-time 1004) is pinpointed by
' the step error handler instead of guessed at.
Private Sub PutBondF(ByVal ws As Worksheet, ByVal col As String, ByVal f As Long, _
                     ByVal lastRow As Long, ByVal formula As String)
    mLastFormulaTarget = ws.Name & "!" & col & f & ":" & col & lastRow
    ws.Range(col & f & ":" & col & lastRow).FormulaR1C1 = formula
End Sub


Private Sub WriteBonds_T1_BDP_Efficient(ByVal ws As Worksheet, ByVal lastRow As Long)


    If lastRow < BOND_DATA_ROW Then Exit Sub


    Dim f As Long
    Dim cfgT1 As String


    f = BOND_DATA_ROW
    cfgT1 = CfgR1C1(CFG_T1_DATE)


    ' =============================================================================
    ' CURRENT / T0 BOND BLOOMBERG WRITER
    '
    ' Legacy function name: WriteBonds_T1_BDP_Efficient
    '
    ' User-facing convention:
    '   T0  = current reporting/as-of date
    '   T-1 = prior snapshot date
    '
    ' This procedure writes current/T0 Bloomberg formulas into Bonds current-market
    ' columns. The function name still contains T1 for legacy reasons only.
    ' =============================================================================


    ' R = FX_T0
    ws.Range(BCOL_FX_T0 & f & ":" & BCOL_FX_T0 & lastRow).FormulaR1C1 = _
        FXT1FormulaR1C1(CfgR1C1(CFG_FX_FIX_SOURCE))


    ' T = CleanPx_T0
    ws.Range(BCOL_CLEANPX_T0 & f & ":" & BCOL_CLEANPX_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_PX_LAST))


    ' U = DirtyPx_T0
    ws.Range(BCOL_DIRTYPX_T0 & f & ":" & BCOL_DIRTYPX_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_PX_DIRTY_MID, "PX_DIRTY"))


    ' V = AccruedInterest_T0
    ws.Range(BCOL_ACCRUEDINTEREST_T0 & f & ":" & BCOL_ACCRUEDINTEREST_T0 & lastRow).FormulaR1C1 = _
        BBGAccruedInterestFormulaR1C1(cfgT1)


    ' W = YTM_T0
    ws.Range(BCOL_YTM_T0 & f & ":" & BCOL_YTM_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_YLD_YTM_MID, "YLD_YTM_BID", "YLD_YTM_ASK", "YLD_YTM"))




    ' X = ModDur_T0
    '
    ' Use YAS_MOD_DUR first if this is the desk/company standard.
    ' Then fallback to generic Bloomberg duration fields.
    ws.Range(BCOL_MODDUR_T0 & f & ":" & BCOL_MODDUR_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array( _
            "YAS_MOD_DUR", _
            BBG_DUR_ADJ_MID, _
            "MOD_DUR_MID", _
            "DUR_MID", _
            "DURATION" _
        ))
    ' Y = Convexity
    ' Do NOT pull Bloomberg CONVEXITY here.
    ' Y is owned by WriteBondConvexityBumpFormulas in Step 3E.
    ' This avoids #N/A Authorization diagnostics from Bloomberg CONVEXITY.
    ws.Range(BCOL_CONVEXITY & f & ":" & BCOL_CONVEXITY & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), """""")
    
    ' Z = ZSprd_T0
    ws.Range(BCOL_ZSPRD_T0 & f & ":" & BCOL_ZSPRD_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_ZSPREAD, "Z_SPRD", "Z_SPREAD_MID", "Z_SPREAD"))


    ' AA = ASW_T0
    ws.Range(BCOL_ASW_T0 & f & ":" & BCOL_ASW_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_ASW, "ASSET_SWAP_SPD", "ASW_SPREAD_MID", "ASW_SPREAD"))


    ' AR = OAS_T0
    ws.Range(BCOL_OAS_T0 & f & ":" & BCOL_OAS_T0 & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_OAS, "OAS_SPREAD", "OAS_MID"))
    
    ' -------------------------------------------------------------------------
    ' AU = DV01_Unit
    '
    ' DV01_Unit means EUR DV01 per one unit of nominal.
    '
    ' Duration hierarchy:
    '   1. SpreadDuration_Used
    '   2. ModDur_T0
    '
    ' Formula:
    '   Duration * DirtyPrice / 100 * FX * 0.0001
    '
    ' AU is derived. It is not retrieved from Bloomberg and must not be frozen
    ' as part of the Step 3C Bloomberg output.
    ' -------------------------------------------------------------------------
    
    ws.Range(BCOL_DV01_UNIT & f & ":" & BCOL_DV01_UNIT & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
            "IF(AND(" & _
                "ISNUMBER(" & RC(BCOL_DIRTYPX_T0) & ")," & _
                "ISNUMBER(" & RC(BCOL_FX_T0) & ")," & _
                "OR(ISNUMBER(" & RC(BCOL_SPREADDURATION_USED) & "),ISNUMBER(" & RC(BCOL_MODDUR_T0) & "))" & _
            ")," & _
                "PRODUCT(" & _
                    "ABS(IF(ISNUMBER(" & RC(BCOL_SPREADDURATION_USED) & ")," & RC(BCOL_SPREADDURATION_USED) & "," & RC(BCOL_MODDUR_T0) & "))," & _
                    "ABS(" & RC(BCOL_DIRTYPX_T0) & ")/100," & _
                    "ABS(" & RC(BCOL_FX_T0) & ")," & _
                    "0.0001" & _
                ")," & _
                """""" & _
            ")" & _
        ")"




    ' AW = OAS_ModDuration_Raw
    ws.Range(BCOL_OAS_MODDURATION_RAW & f & ":" & BCOL_OAS_MODDURATION_RAW & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_DUR_ADJ_OAS_MID, "OAS_MOD_DUR", "OAS_DUR_MID", "DUR_ADJ_OAS"))


    ' AX = OAS_Convexity
    ws.Range(BCOL_OAS_CONVEXITY & f & ":" & BCOL_OAS_CONVEXITY & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array("CONVEXITY", BBG_CONVEXITY_OAS, "OAS_CONVEXITY", "CONVEXITY_OAS_MID"))


    ' AY = Pricing_Source
    ws.Range(BCOL_PRICING_SOURCE & f & ":" & BCOL_PRICING_SOURCE & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array(BBG_PRICING_SOURCE, "PX_SOURCE", "PRICING_SOURCE_DES"))


    ' AZ = Last_Pricing_Date
    ws.Range(BCOL_LAST_PRICING_DATE & f & ":" & BCOL_LAST_PRICING_DATE & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array("PX_DATE", "LAST_UPDATE_DT", "LAST_UPDATE_DATE", "LAST_PRICE_TIME"))


    ' BA = Benchmark_Bond
    
    ws.Range(BCOL_BENCHMARK_BOND & f & ":" & BCOL_BENCHMARK_BOND & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1( _
            Array( _
                "YAS_BOND_BENCHMARK", _
                "YAS_BENCHMARK_BOND", _
                "BENCHMARK_BOND"))


    ' BP = FundTkr
    ws.Range(BCOL_FUNDTKR & f & ":" & BCOL_FUNDTKR & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), _
        "IF(" & RC(BCOL_CCY) & "=""EUR"",""EUR003M Index"",IF(" & RC(BCOL_CCY) & "=""USD"",""US0003M Index"",IF(" & RC(BCOL_CCY) & "=""GBP"",""BP0003M Index"","""")))")


    ' BQ = FundRate_T0
    ws.Range(BCOL_FUNDRATE_T0 & f & ":" & BCOL_FUNDRATE_T0 & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""","""",IF(" & RC(BCOL_FUNDTKR) & "="""",""""," & _
            GuardedBDPExprR1C1(RC(BCOL_FUNDTKR), XlText(BBG_PX_LAST), XlBlank()) & _
        "))"


    ' CL = DAY_CNT_DES
    ws.Range(BCOL_DAY_CNT_DES & f & ":" & BCOL_DAY_CNT_DES & lastRow).FormulaR1C1 = _
        BBGFirstBDPMultiFieldFormulaR1C1(Array("DAY_CNT_DES", "DAY_COUNT_DES", "DAY_CNT"))


    ' CM = BondDCC_Code
    ws.Range(BCOL_BONDDCC_CODE & f & ":" & BCOL_BONDDCC_CODE & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IF(" & RC(BCOL_DAY_CNT_DES) & "="""","""",BloombergDayCountToDCC(" & RC(BCOL_DAY_CNT_DES) & "))")


    ' CN = BondDCC_Name
    ws.Range(BCOL_BONDDCC_NAME & f & ":" & BCOL_BONDDCC_NAME & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""","""",IFERROR(SWITCH(" & RC(BCOL_BONDDCC_CODE) & "," & _
            "0,""ACT/ACT ICMA""," & _
            "1,""ACT/ACT ISDA""," & _
            "2,""ACT/365F""," & _
            "3,""ACT/360""," & _
            "4,""30/360 BOND""," & _
            "5,""30E/360""," & _
            "6,""30E/360 ISDA""," & _
            """Unknown DCC"")," & RC(BCOL_DAY_CNT_DES) & "))"


End Sub
' =============================================================================
' CURRENT/T0 AND PRIOR/T-1 DERIVED BOND FORMULAS
'
' Legacy code naming:
'   useT1=True  = user-facing T0 / current
'   useT1=False = user-facing T-1 / prior
'
' Key Bonds columns:
'   N  = OIS_T0/current
'   O  = OIS_T-1/prior
'   P  = DF_T0/current
'   Q  = DF_T-1/prior
'   AK = ISpread_T0/current
'   AL = ISpread_T-1/prior
'   AM = GSpread_T0/current
'   AN = GSpread_T-1/prior
'   BC = Gov_T0/current
'   BD = Gov_T-1/prior
'   BE = Swap_T0/current
'   BF = Swap_T-1/prior
'   BG = g_T0/current
'   BH = g_T-1/prior
'   BJ = q_T0/current
'   BK = q_T-1/prior
' =============================================================================


Private Sub WriteBondsCalculatedFormulas_Efficient(ByVal ws As Worksheet, ByVal lastRow As Long)


    If lastRow < BOND_DATA_ROW Then Exit Sub


    Dim f As Long
    f = BOND_DATA_ROW


    Dim cfgT1 As String
    cfgT1 = CfgR1C1(CFG_T1_DATE)


    ' M = DaysLeft
    ws.Range(BCOL_DAYSLEFT & f & ":" & BCOL_DAYSLEFT & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "MAX(0," & RC(BCOL_MATURITY) & "-INT(" & cfgT1 & "))")


    ' N = OIS_T0
    ws.Range(BCOL_OIS_T0 & f & ":" & BCOL_OIS_T0 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpOIS(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,TRUE),"""")")


    ' O = OIS_T-1
    ws.Range(BCOL_OIS_TM1 & f & ":" & BCOL_OIS_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpOIS(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,FALSE),"""")")


    ' P = DF_T0
    ws.Range(BCOL_DF_T0 & f & ":" & BCOL_DF_T0 & lastRow).FormulaR1C1 = _
        "=IFERROR(EXP(-(" & RC(BCOL_OIS_T0) & "/100)*(" & RC(BCOL_DAYSLEFT) & "/365)),"""")"


    ' Q = DF_T-1
    ws.Range(BCOL_DF_TM1 & f & ":" & BCOL_DF_TM1 & lastRow).FormulaR1C1 = _
        "=IFERROR(EXP(-(" & RC(BCOL_OIS_TM1) & "/100)*(" & RC(BCOL_DAYSLEFT) & "/365)),"""")"


    ' BC = Gov_T0
    ws.Range(BCOL_GOV_T0 & f & ":" & BCOL_GOV_T0 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpGov(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,TRUE),"""")")


    ' BD = Gov_T-1
    ws.Range(BCOL_GOV_TM1 & f & ":" & BCOL_GOV_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpGov(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,FALSE),"""")")


    ' BE = Swap_T0
    ws.Range(BCOL_SWAP_T0 & f & ":" & BCOL_SWAP_T0 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpSwap(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,TRUE),"""")")


    ' BF = Swap_T-1
    ws.Range(BCOL_SWAP_TM1 & f & ":" & BCOL_SWAP_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IFERROR(InterpSwap(" & RC(BCOL_CCY) & "," & RC(BCOL_DAYSLEFT) & "/365,FALSE),"""")")


    ' BG = g_T0 = Gov_T0 - OIS_T0
    ws.Range(BCOL_G_T0 & f & ":" & BCOL_G_T0 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_GOV_T0), RC(BCOL_OIS_T0))


    ' BH = g_T-1 = Gov_T-1 - OIS_T-1
    ws.Range(BCOL_G_TM1 & f & ":" & BCOL_G_TM1 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_GOV_TM1), RC(BCOL_OIS_TM1))


    ' BI = Delta_g_bp
    ws.Range(BCOL_DELTA_G_BP & f & ":" & BCOL_DELTA_G_BP & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_G_T0), RC(BCOL_G_TM1), True)


    ' BJ = q_T0 = Swap_T0 - Gov_T0
    ws.Range(BCOL_Q_T0 & f & ":" & BCOL_Q_T0 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_SWAP_T0), RC(BCOL_GOV_T0))


    ' BK = q_T-1 = Swap_T-1 - Gov_T-1
    ws.Range(BCOL_Q_TM1 & f & ":" & BCOL_Q_TM1 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_SWAP_TM1), RC(BCOL_GOV_TM1))


    ' BL = Delta_q_bp
    ws.Range(BCOL_DELTA_Q_BP & f & ":" & BCOL_DELTA_Q_BP & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_Q_T0), RC(BCOL_Q_TM1), True)


    ' AK = ISpread_T0 = YTM_T0 - Swap_T0
    ws.Range(BCOL_ISPREAD_T0 & f & ":" & BCOL_ISPREAD_T0 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_YTM_T0), RC(BCOL_SWAP_T0), True)


    ' AL = ISpread_T-1 = YTM_T-1 - Swap_T-1
    ws.Range(BCOL_ISPREAD_TM1 & f & ":" & BCOL_ISPREAD_TM1 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_YTM_TM1), RC(BCOL_SWAP_TM1), True)


    ' AM = GSpread_T0 = YTM_T0 - Gov_T0
    ws.Range(BCOL_GSPREAD_T0 & f & ":" & BCOL_GSPREAD_T0 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_YTM_T0), RC(BCOL_GOV_T0), True)


    ' AN = GSpread_T-1 = YTM_T-1 - Gov_T-1
    ws.Range(BCOL_GSPREAD_TM1 & f & ":" & BCOL_GSPREAD_TM1 & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_YTM_TM1), RC(BCOL_GOV_TM1), True)


    ' BM = Delta_i_bp
    ws.Range(BCOL_DELTA_I_BP & f & ":" & BCOL_DELTA_I_BP & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_ISPREAD_T0), RC(BCOL_ISPREAD_TM1))


    ' BN = Delta_y_bp
    ws.Range(BCOL_DELTA_Y_BP & f & ":" & BCOL_DELTA_Y_BP & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_YTM_T0), RC(BCOL_YTM_TM1), True)


    ' DV01_EUR

    ws.Range(BCOL_DV01_EUR & f & ":" & _
             BCOL_DV01_EUR & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
            "IF(AND(" & _
                "ISNUMBER(" & RC(BCOL_NOTIONAL) & ")," & _
                "ISNUMBER(" & RC(BCOL_DV01_UNIT) & ")" & _
            ")," & _
                RC(BCOL_NOTIONAL) & "*" & _
                RC(BCOL_DV01_UNIT) & "," & _
                """""" & _
            ")" & _
        ")"


    ' CM = DV01_Opening_EUR
    '
    ' The same DV01, struck on T-1 price and T-1 FX instead of T0.
    '
    ' AF DV01_EUR is the risk the book carries NOW, which is the right number
    ' for a hedge ratio, a hedge gap or a coverage weight - all statements
    ' about today's position.  It is the WRONG number to attribute yesterday's
    ' move to.  An attribution says 'I began the day holding this much risk,
    ' the curve moved that much, so this is what it earned', and that sentence
    ' is about the OPENING risk.  Using closing risk credits the desk with a
    ' position size it only reached after the move it is being paid for.
    '
    ' It also has to be opening risk for the second-order term to be right.
    ' The Taylor expansion  dP = -DV01*dy + 0.5*MV*C*dy^2  carries the PLUS on
    ' convexity only when both terms are anchored at the START of the move;
    ' anchored at the end the convexity term flips sign.  PnL_Convexity is
    ' struck on DirtyMV_T-1, so the duration leg beside it had to be too, or
    ' the two legs are expansions around different points and the convexity is
    ' being added when it should be subtracted.
    '
    ' Duration itself is taken at T0: a one-day change in modified duration is
    ' third-order here, and Bloomberg is not asked for a second, dated pull.
    ' What matters - the price and the FX rate - are both taken at T-1.
    ws.Range(BCOL_DV01_OPENING_EUR & f & ":" & _
             BCOL_DV01_OPENING_EUR & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
            "IF(AND(" & _
                "ISNUMBER(" & RC(BCOL_NOTIONAL) & ")," & _
                "ISNUMBER(" & RC(BCOL_DIRTYPX_TM1) & ")," & _
                "ISNUMBER(" & RC(BCOL_FX_TM1) & ")," & _
                "OR(ISNUMBER(" & RC(BCOL_SPREADDURATION_USED) & "),ISNUMBER(" & RC(BCOL_MODDUR_T0) & "))" & _
            ")," & _
                RC(BCOL_NOTIONAL) & "*PRODUCT(" & _
                    "ABS(IF(ISNUMBER(" & RC(BCOL_SPREADDURATION_USED) & ")," & RC(BCOL_SPREADDURATION_USED) & "," & RC(BCOL_MODDUR_T0) & "))," & _
                    "ABS(" & RC(BCOL_DIRTYPX_TM1) & ")/100," & _
                    "ABS(" & RC(BCOL_FX_TM1) & ")," & _
                    "0.0001" & _
                ")," & _
                """""" & _
            ")" & _
        ")"


    ' AH = DirtyMV_T0_EUR
    ws.Range(BCOL_DIRTYMV_T0_EUR & f & ":" & _
             BCOL_DIRTYMV_T0_EUR & lastRow).FormulaR1C1 = _
        "=IF(AND(" & _
        "ISNUMBER(" & RC(BCOL_NOTIONAL) & ")," & _
        "ISNUMBER(" & RC(BCOL_DIRTYPX_T0) & ")," & _
        "ISNUMBER(" & RC(BCOL_FX_T0) & "))," & _
        RC(BCOL_NOTIONAL) & "*" & _
        RC(BCOL_DIRTYPX_T0) & "/100*" & _
        RC(BCOL_FX_T0) & ","""")"


    ' AI = DirtyMV_T-1_EUR
    ws.Range(BCOL_DIRTYMV_TM1_EUR & f & ":" & _
             BCOL_DIRTYMV_TM1_EUR & lastRow).FormulaR1C1 = _
        "=IF(AND(" & _
        "ISNUMBER(" & RC(BCOL_NOTIONAL) & ")," & _
        "ISNUMBER(" & RC(BCOL_DIRTYPX_TM1) & ")," & _
        "ISNUMBER(" & RC(BCOL_FX_TM1) & "))," & _
        RC(BCOL_NOTIONAL) & "*" & _
        RC(BCOL_DIRTYPX_TM1) & "/100*" & _
        RC(BCOL_FX_TM1) & ","""")"
        
    
    ' AJ = BookVal_EUR2
    ws.Range(BCOL_BOOKVAL_EUR2 & f & ":" & BCOL_BOOKVAL_EUR2 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), RC(BCOL_BOOKVAL))


    ' AT = DeltaOAS
    ws.Range(BCOL_DELTAOAS & f & ":" & BCOL_DELTAOAS & lastRow).FormulaR1C1 = _
        DiffFormula(RC(BCOL_OAS_T0), RC(BCOL_OAS_TM1))


    ' -------------------------------------------------------------------------
    ' AU = SpreadDuration_Used     (feeds AT = DV01_Unit, and nothing else)
    '
    ' PREFERENCE ORDER: ModDur_T0 first, OAS_ModDuration_Raw only as a fallback.
    '
    ' This was the other way round, and it was wrong for eight of the nine
    ' columns that consume it.  DV01_Unit is not a spread-only figure: it is the
    ' ONLY DV01 on PNL_Attribution, and every first-order leg multiplies it -
    '
    '     PnL_OIS            x Delta_r          a rate move
    '     PnL_GovBasis       x Delta_g          a rate move
    '     PnL_SwapGovBasis   x Delta_q          a rate move
    '     PnL_Credit_Ispread x Delta_i          derived from the YIELD
    '     PnL_GSpread        x Delta_GSpread    derived from the YIELD
    '     PnL_ZSpread        x Delta_Z          quoted, ~ yield duration
    '     PnL_ASW            x Delta_ASW        quoted, ~ yield duration
    '     Duration_Identity  x Delta_y          the yield move itself
    '     PnL_OAS            x Delta_OAS        the one genuine OAS move
    '
    ' - so the sensitivity wanted almost everywhere is the ORDINARY modified
    ' duration.  OAS duration is the option-adjusted one: on a callable bond it
    ' is deliberately shorter, and using it against a curve move understates
    ' every rate leg on exactly the bonds where the split matters most.
    '
    ' Why this was invisible: Duration_Identity_Check compares the chain against
    ' -DV01 * Delta_y using the SAME DV01 on both sides, so a wrong duration
    ' cancels and the check still ties.  And Bloomberg returns DUR_ADJ_OAS_MID
    ' for bullet bonds too, where it equals modified duration - so the error only
    ' ever showed on the callable minority and never tripped a check.
    '
    ' KNOWN REMAINING APPROXIMATION: PnL_OAS now also uses the modified duration.
    ' For a callable bond its own OAS duration (Bonds!AV, kept and still pulled)
    ' is the better sensitivity for that ONE leg.  Splitting it needs a second
    ' DV01 column, which is a deliberate decision for the desk rather than a
    ' side effect of this fix.
    '
    ' ABS() is kept: some Bloomberg duration fields arrive sign-flipped, and a
    ' negative duration here would invert the sign of every leg above.  The
    ' position's own sign lives in Notional, which DV01_EUR applies separately.
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_SPREADDURATION_USED & f & ":" & BCOL_SPREADDURATION_USED & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
            "IF(ISNUMBER(" & RC(BCOL_MODDUR_T0) & ")," & _
                "ABS(" & RC(BCOL_MODDUR_T0) & ")," & _
                "IF(ISNUMBER(" & RC(BCOL_OAS_MODDURATION_RAW) & "),ABS(" & RC(BCOL_OAS_MODDURATION_RAW) & "),"""")" & _
            ")" & _
        ")"


    ' AO = ISpread_Status
    ws.Range(BCOL_ISPREAD_STATUS & f & ":" & BCOL_ISPREAD_STATUS & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), _
        "IF(OR(NOT(ISNUMBER(" & RC(BCOL_ISPREAD_T0) & ")),NOT(ISNUMBER(" & RC(BCOL_ISPREAD_TM1) & "))),""Missing ISpread T0/T1"",""OK"")")


    ' BS = Bond_Status
    ws.Range(BCOL_BOND_STATUS & f & ":" & BCOL_BOND_STATUS & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), _
        "IF(" & RC(BCOL_TICKER_STATUS) & "<>""OK""," & RC(BCOL_TICKER_STATUS) & ",IF(OR(NOT(ISNUMBER(" & RC(BCOL_DIRTYPX_T0) & ")),NOT(ISNUMBER(" & RC(BCOL_DIRTYPX_TM1) & ")),NOT(ISNUMBER(" & RC(BCOL_YTM_T0) & ")),NOT(ISNUMBER(" & _
        RC(BCOL_YTM_TM1) & "))),""Missing market data"",""OK""))")


End Sub






' =============================================================================
' EFFICIENT FUTURES WRITERS
'   - Generic ticker via FutMap lookup (col V=22), not blind code&"1 Comdty".
'   - CTD resolved from OPICS CTD_ISIN (col G=7) into AA=27, used for CTD px/DV01.
'   - All config refs use CfgR1C1 (fixes DaysToDeliv #NAME? and futures px pull).
'   Column map: ContCode=A1 CTD_ISIN=G7 CTD_CF=H8 DelivDate=F6 Contracts=D4
'               FaceValue=E5 FX=M13 FutPx_T1=N14 FutPx_T0=Q17 CTDtkr=AA27
' =============================================================================


Private Sub WriteFutures_T1_BDP_Efficient(ws As Worksheet, lastRow As Long)


    If lastRow < DATA_ROW Then Exit Sub


    Dim f As Long
    Dim cfgFutPx As String
    Dim fml As String


    f = DATA_ROW
    cfgFutPx = CfgR1C1(CFG_FUT_PRICE_FIELD)


    ' -------------------------------------------------------------------------
    ' V = BBG_Ticker
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), "IFERROR(VLOOKUP(" & RC(FCOL_CONTRACTCODE) & ",FutMapTable,2,FALSE),"""")")
    ws.Range(FCOL_BBG_TICKER & f & ":" & FCOL_BBG_TICKER & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' F = DelivDate
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_BBG_TICKER), "IFERROR(BDP(" & RC(FCOL_BBG_TICKER) & ",""LAST_TRADEABLE_DT""),IFERROR(BDP(" & RC(FCOL_BBG_TICKER) & ",""FUT_DLV_DT""),""""))")
    ws.Range(FCOL_DELIVDATE & f & ":" & FCOL_DELIVDATE & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' M = FX_T0
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_CONTRACTCODE) & "="""","""",LET("
    fml = fml & "_ccy,UPPER(TRIM(" & RC(FCOL_CCY) & ")),"
    fml = fml & "IF(_ccy="""","""","
    fml = fml & "IF(_ccy=""EUR"",1,"
    fml = fml & "IFERROR(1/BDP(""EUR""&_ccy&"" Curncy"",""PX_LAST""),"
    fml = fml & "IFERROR(BDP(_ccy&""EUR Curncy"",""PX_LAST""),""""))"
    fml = fml & ")))"
    fml = fml & ")"
    ws.Range(FCOL_FX_T0 & f & ":" & FCOL_FX_T0 & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' N = FutPx_T0
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_BBG_TICKER) & "="""","""",IFERROR(BDP(" & RC(FCOL_BBG_TICKER) & ","
    fml = fml & cfgFutPx
    fml = fml & "),IFERROR(BDP(" & RC(FCOL_BBG_TICKER) & ",""PX_LAST""),"""")))"
    ws.Range(FCOL_FUTPX_T0 & f & ":" & FCOL_FUTPX_T0 & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' Y = FUT_VAL_PT
    '
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_BBG_TICKER), "IFERROR(BDP(" & RC(FCOL_BBG_TICKER) & ",""FUT_VAL_PT""),"""")")
    ws.Range(FCOL_FUT_VAL_PT & f & ":" & FCOL_FUT_VAL_PT & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' Z = CTD_Ticker
    '
    ' CTD_Ticker shifted from old AA to new Z.
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_CTD_ISIN) & "="""","""",LET("
    fml = fml & "_a,BDP(""/isin/""&" & RC(FCOL_CTD_ISIN) & ",""PARSEKYABLE_DES""),"
    fml = fml & "IF(IF(ISTEXT(_a),AND(LEFT(_a,2)=""#N"",NOT(ISNUMBER(SEARCH(""Requesting"",_a)))),ISERR(_a)),"
    fml = fml & "LET(_b,BDP(" & RC(FCOL_CTD_ISIN) & "&"" Govt"",""PARSEKYABLE_DES""),"
    fml = fml & "IF(IF(ISTEXT(_b),AND(LEFT(_b,2)=""#N"",NOT(ISNUMBER(SEARCH(""Requesting"",_b)))),ISERR(_b)),"
    fml = fml & "LET(_c,BDP(" & RC(FCOL_CTD_ISIN) & "&"" Corp"",""PARSEKYABLE_DES""),"
    fml = fml & "IF(IF(ISTEXT(_c),AND(LEFT(_c,2)=""#N"",NOT(ISNUMBER(SEARCH(""Requesting"",_c)))),ISERR(_c)),"""",_c)),"
    fml = fml & "_b)),"
    fml = fml & "_a)))"
    ws.Range(FCOL_CTD_TICKER & f & ":" & FCOL_CTD_TICKER & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' O = CTD_DirtyPx_T0
    ' New CTD_Ticker is RC26 = column Z.
    ' -------------------------------------------------------------------------
    ws.Range(FCOL_CTD_DIRTYPX_T0 & f & ":" & FCOL_CTD_DIRTYPX_T0 & lastRow).FormulaR1C1 = _
        FutFirstBDPMultiFieldFormulaR1C1(RC(FCOL_CTD_TICKER), Array("PX_DIRTY_MID", "PX_LAST"))


    ' -------------------------------------------------------------------------
    ' P = CTD_Ticker_T-1
    ' New CTD_Ticker is RC26 = column Z.
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CTD_TICKER), RC(FCOL_CTD_TICKER))
    ws.Range(FCOL_CTD_TICKER_TM1 & f & ":" & FCOL_CTD_TICKER_TM1 & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' AA = ConvFactor
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), "IF(ISNUMBER(" & RC(FCOL_CTD_CF) & ")," & RC(FCOL_CTD_CF) & ","""")")
    ws.Range(FCOL_CONVFACTOR & f & ":" & FCOL_CONVFACTOR & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' AB = HedgeUnitDV01
    '
    ' New source:
    '   Bloomberg field FUT_PX_VAL_BP.
    '
    ' Fallback:
    '   Y = RC25 = FUT_VAL_PT.
    ' I took the fallback out but left  this placeholder here for future reference that if this is consistently wrong this might be an option
    ' -------------------------------------------------------------------------
    ws.Range(FCOL_HEDGEUNITDV01 & f & ":" & FCOL_HEDGEUNITDV01 & lastRow).FormulaR1C1 = _
        FutFirstBDPMultiFieldFormulaR1C1(RC(FCOL_BBG_TICKER), Array("FUT_PX_VAL_BP"), XlBlank())


    ' -------------------------------------------------------------------------
    ' AD = ImpliedRepo_BBG
    '
    ' Old AF shifted to new AD.
    ' -------------------------------------------------------------------------
    ws.Range(FCOL_IMPLIEDREPO_BBG & f & ":" & FCOL_IMPLIEDREPO_BBG & lastRow).FormulaR1C1 = _
        FutFirstBDPMultiFieldFormulaR1C1( _
            RC(FCOL_BBG_TICKER), _
            Array("FUT_IMPLIED_REPO_RT", "FUT_IMPLIED_REPO_RATE", "IMPLIED_REPO_RATE"), _
            "IF(ISNUMBER(" & RC(FCOL_IMPLIEDREPO_CALC) & ")," & RC(FCOL_IMPLIEDREPO_CALC) & "," & XlBlank() & ")" _
        )


    ' -------------------------------------------------------------------------
    ' AE = NetBasis_BBG
    ' Old AG shifted to new AE.
    ' -------------------------------------------------------------------------
    ws.Range(FCOL_NETBASIS_BBG & f & ":" & FCOL_NETBASIS_BBG & lastRow).FormulaR1C1 = _
        FutFirstBDPMultiFieldFormulaR1C1( _
            RC(FCOL_BBG_TICKER), _
            Array("FUT_NET_BASIS", "NET_BASIS"), _
            XlBlank() _
        )


    ' -------------------------------------------------------------------------
    ' AF = GrossBasis_BBG
    '
    ' Old AH shifted to new AF.
    ' -------------------------------------------------------------------------
    ws.Range(FCOL_GROSSBASIS_BBG & f & ":" & FCOL_GROSSBASIS_BBG & lastRow).FormulaR1C1 = _
        FutFirstBDPMultiFieldFormulaR1C1( _
            RC(FCOL_BBG_TICKER), _
            Array("FUT_GROSS_BASIS", "GROSS_BASIS"), _
            "IF(ISNUMBER(" & RC(FCOL_GROSSBASIS_CALC) & ")," & RC(FCOL_GROSSBASIS_CALC) & "," & XlBlank() & ")" _
        )


    ' -------------------------------------------------------------------------
    ' W and AG = Status
    '
    ' Old AI shifted to new AG.
    ' Old Z FUT_CONT_SIZE and old AB CTD_DV01 are no longer required.
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), _
        "IF(" & RC(FCOL_BBG_TICKER) & "="""",""Unmapped contract (FutMap)""," & "IF(" & RC(FCOL_FUTPX_T0) & "="""",""Missing T0 future price""," & "IF(" & RC(FCOL_CTD_ISIN) & "="""",""Missing CTD_ISIN""," & "IF(" & RC(FCOL_CTD_CF) & _
        "="""",""Missing CTD_CF""," & "IF(" & RC(FCOL_CTD_TICKER) & "="""",""Missing CTD ticker""," & "IF(" & RC(FCOL_CTD_DIRTYPX_T0) & "="""",""Missing CTD dirty price""," & "IF(" & RC(FCOL_FUT_VAL_PT) & "="""",""Missing FUT_VAL_PT""," & _
        "IF(" & RC(FCOL_HEDGEUNITDV01) & "="""",""Missing FUT_PX_VAL_BP / HedgeUnitDV01"",""OK""))))))))")


    ws.Range(FCOL_STATUS_T0 & f & ":" & FCOL_STATUS_T0 & lastRow).FormulaR1C1 = fml
    ws.Range(FCOL_STATUS & f & ":" & FCOL_STATUS & lastRow).FormulaR1C1 = fml


End Sub
Private Sub WriteSwaps_T1_BQL_Formulas(ByVal ws As Worksheet, ByVal lastRow As Long)


    If lastRow < DATA_ROW Then Exit Sub


    Dim f As Long
    f = DATA_ROW


    ' -------------------------------------------------------------------------
    ' AL:AP are mapping/economic fields from Coverage support.
    ' Do NOT call Bloomberg for notional.
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_MAP_NOTIONAL & f & ":" & WCOL_MAP_NOTIONAL & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), RC(WCOL_NOTIONAL))
    ws.Range(WCOL_MAP_CCY & f & ":" & WCOL_MAP_CCY & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), RC(WCOL_CCY))
    ws.Range(WCOL_MAP_COUNTERPARTY & f & ":" & WCOL_MAP_COUNTERPARTY & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), RC(WCOL_CPTY))
    ws.Range(WCOL_NOTIONAL_FINAL & f & ":" & WCOL_NOTIONAL_FINAL & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), "IF(ISNUMBER(" & RC(WCOL_NOTIONAL) & "),ABS(" & RC(WCOL_NOTIONAL) & "),"""")")
    ws.Range(WCOL_NOTIONAL_SOURCE & f & ":" & WCOL_NOTIONAL_SOURCE & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), """COVERAGE_SUPPORT""")


    ' -------------------------------------------------------------------------
    ' AQ:AS = T1 NPV from direct / fixed / float securities.
    '
    ' AD = RC30 = Direct Bloomberg ID
    ' AE = RC31 = Fixed leg Bloomberg ID
    ' AF = RC32 = Float leg Bloomberg ID
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_BQL_NPV_DIRECT_T0 & f & ":" & WCOL_BQL_NPV_DIRECT_T0 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_SWAP_DIRECT_ID), XlText(BQL_SWAP_MV_T1), CFG_T1_DATE)


    ws.Range(WCOL_BQL_NPV_FIXED_T0 & f & ":" & WCOL_BQL_NPV_FIXED_T0 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_FIXED_LEG_ID), XlText(BQL_SWAP_MV_T1), CFG_T1_DATE)


    ws.Range(WCOL_BQL_NPV_FLOAT_T0 & f & ":" & WCOL_BQL_NPV_FLOAT_T0 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_FLOAT_LEG_ID), XlText(BQL_SWAP_MV_T1), CFG_T1_DATE)


    ' BN = DV01_BBG.
    '
    ' Bloomberg-sourced swap BPV used directly by PNL_Attribution.
    ' The internal model DV01 in column X is not used for PNL attribution.
    
    ws.Range(WCOL_DV01_BBG & f & ":" & WCOL_DV01_BBG & lastRow).FormulaR1C1 = _
        "=IF(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & "="""",""""," & _
        "IFERROR(BDP(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & ",""SW_CNV_BPV""," & _
        """SW_RECEIVE_CURVE_MKT_SIDE"",""Mid""),""""))"




    ' AT = total T1 NPV.
    ws.Range(WCOL_BQL_NPV_TOTAL_T0 & f & ":" & WCOL_BQL_NPV_TOTAL_T0 & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), _
        "IF(ISNUMBER(" & RC(WCOL_BQL_NPV_DIRECT_T0) & ")," & RC(WCOL_BQL_NPV_DIRECT_T0) & ",IF(OR(ISNUMBER(" & RC(WCOL_BQL_NPV_FIXED_T0) & "),ISNUMBER(" & RC(WCOL_BQL_NPV_FLOAT_T0) & ")),SUM(" & RC(WCOL_BQL_NPV_FIXED_T0) & ":" & _
        RC(WCOL_BQL_NPV_FLOAT_T0) & "),""""))")


    ' -------------------------------------------------------------------------
    ' BA = PAY_FLT_RATE_IDX.
    '
    ' Prefer float leg ID in AF / RC32.
    ' Fallback to direct swap ID in AD / RC30.
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_PAY_FLT_RATE_IDX & f & ":" & WCOL_PAY_FLT_RATE_IDX & lastRow).FormulaR1C1 = _
        "=IF(" & RC(WCOL_DEALID) & "="""",""""," & _
            "IF(" & RC(WCOL_BBG_FLOAT_LEG_ID) & "<>"""",IFERROR(BDP(" & RC(WCOL_BBG_FLOAT_LEG_ID) & ",""" & BBG_PAY_FLT_RATE_IDX & """)," & _
                "IF(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & "<>"""",IFERROR(BDP(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & ",""" & BBG_PAY_FLT_RATE_IDX & """),""""),""""))," & _
            "IF(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & "<>"""",IFERROR(BDP(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & ",""" & BBG_PAY_FLT_RATE_IDX & """),""""),""""))" & _
        ")"


    ' -------------------------------------------------------------------------
    ' D / E / F = swap economics DERIVED from Bloomberg (float index = BA / PAY_FLT_RATE_IDX).
    '   E = FloatIndex  : mirror BA.
    '   D = FixedRate   : entry rate = the float index's yield at the swap StartDate (G),
    '                     i.e. the floater's level at inception (par assumption).  BDH point.
    '   F = FloatSpread : current float-index yield (BDP) minus D (drift since inception).
    '
    ' idxRef normalises the BA ticker so BDP/BDH can resolve it: if BA already carries a
    ' yellow-key suffix (Index/Curncy/Comdty) use it as-is, otherwise append " Index".
    ' -------------------------------------------------------------------------
    Dim idxRef As String
    idxRef = "IF(OR(ISNUMBER(SEARCH("" Index""," & RC(WCOL_PAY_FLT_RATE_IDX) & "))," & _
                 "ISNUMBER(SEARCH("" Curncy""," & RC(WCOL_PAY_FLT_RATE_IDX) & "))," & _
                 "ISNUMBER(SEARCH("" Comdty""," & RC(WCOL_PAY_FLT_RATE_IDX) & ")))," & _
                 RC(WCOL_PAY_FLT_RATE_IDX) & "," & RC(WCOL_PAY_FLT_RATE_IDX) & "&"" Index"")"


    ' E = FloatIndex (mirror BA / PAY_FLT_RATE_IDX)
    ws.Range(WCOL_FLOATINDEX & f & ":" & WCOL_FLOATINDEX & lastRow).FormulaR1C1 = _
        WrapIfPresent(RC(WCOL_DEALID), RC(WCOL_PAY_FLT_RATE_IDX))


    ' D = FixedRate = float-index yield at StartDate (last value in the week ending at G)
    ws.Range(WCOL_FIXEDRATE & f & ":" & WCOL_FIXEDRATE & lastRow).FormulaR1C1 = _
        "=IF(OR(" & RC(WCOL_DEALID) & "=""""," & RC(WCOL_PAY_FLT_RATE_IDX) & "=""""," & RC(WCOL_STARTDATE) & "=""""),""""," & _
            BDHLastPointExprR1C1(idxRef, XlText(BBG_PX_LAST), "INT(" & RC(WCOL_STARTDATE) & ")-7", "INT(" & RC(WCOL_STARTDATE) & ")") & _
        ")"


    ' F = FloatSpread = current float-index yield (BDP) - D (FixedRate)
    ws.Range(WCOL_FLOATSPREAD & f & ":" & WCOL_FLOATSPREAD & lastRow).FormulaR1C1 = _
        "=IF(" & RC(WCOL_DEALID) & "="""",""""," & _
            "IF(ISNUMBER(" & RC(WCOL_FIXEDRATE) & "),IFERROR(BDP(" & idxRef & ",""" & BBG_PX_LAST & """)-" & RC(WCOL_FIXEDRATE) & ",""""),"""")" & _
        ")"


End Sub
' Futures T0 status only.  FutPx_T0 (col Q) is now pulled on the hidden staging
' sheet (WriteFutureT0Staging) and copied back as a static value, so no BDH array
' formula is written into the live Futures rows anymore.
Private Sub WriteFutures_T0_Status(ws As Worksheet, lastRow As Long)
    If lastRow < DATA_ROW Then Exit Sub


    ws.Range(FCOL_STATUS_TM1 & DATA_ROW & ":" & FCOL_STATUS_TM1 & lastRow).FormulaR1C1 = _
        "=IF(" & RC(FCOL_FUTPX_TM1) & "="""",""Missing T0 future price"",""OK"")"
End Sub


Private Sub WriteFuturesCalculatedFormulas_Efficient(ws As Worksheet, lastRow As Long)


    If lastRow < DATA_ROW Then Exit Sub


    Dim f As Long
    Dim cfgT1 As String
    Dim fml As String


    f = DATA_ROW
    cfgT1 = CfgR1C1(CFG_T1_DATE)


    ' -------------------------------------------------------------------------
    ' L = DaysToDeliv
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), "IF(ISNUMBER(" & RC(FCOL_DELIVDATE) & "),MAX(0," & RC(FCOL_DELIVDATE) & "-" & cfgT1 & "),"""")")
    ws.Range(FCOL_DAYSTODELIV & f & ":" & FCOL_DAYSTODELIV & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' AP = Hedge_Class
    '
    ' The book is EUR/USD hedged, so the Futures sheet does not hold only bond
    ' futures: the EUR/USD contracts that exist purely to translate the USD
    ' positions back to EUR arrive through the same coverage import.  Nothing
    ' distinguished them before, so FutureCodeFromCoverageLabel stripped their
    ' label down to a contract code like any other row and they were priced as
    ' government futures - contributing a CTD-derived DV01 to the rates hedge of
    ' whatever bond they were linked to, which is risk the book does not have.
    '
    ' Classification is by LABEL, and the default is RATES, so no row this rule
    ' does not recognise changes behaviour.  Config!B19 takes extra tokens for
    ' whatever the desk's coverage file calls its FX hedges, so extending the
    ' list is a config edit rather than a code change.
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), FxHedgeClassExpr())
    ws.Range(FCOL_HEDGE_CLASS & f & ":" & FCOL_HEDGE_CLASS & lastRow).FormulaR1C1 = fml




    ' -------------------------------------------------------------------------
    ' R = Implied Repo using the Public VBA UDF directly
    ' -------------------------------------------------------------------------
    fml = WrapIfPresent(RC(FCOL_CONTRACTCODE), _
        "IF(" & "AND(" & "ISNUMBER(" & RC(FCOL_CCY) & ")," & "ISNUMBER(" & RC(FCOL_CONTRACTS) & ")," & "ISNUMBER(" & RC(FCOL_CTD_CF) & ")," & "ISNUMBER(" & RC(FCOL_PORTFOLIO) & ")," & "ISNUMBER(" & RC(FCOL_LINKEDISIN) & ")," & "ISNUMBER(" & _
        RC(FCOL_AVGENTRYPX) & ")," & "ISNUMBER(" & RC(FCOL_FX_T0) & ")," & "ISNUMBER(" & RC(FCOL_FUTPX_T0) & ")," & "ISNUMBER(" & RC(FCOL_FUTURESPNL_EUR) & ")," & "ISNUMBER(" & RC(FCOL_BBG_TICKER) & ")," & RC(FCOL_CTD_CF) & "<>0," & _
        RC(FCOL_LINKEDISIN) & ">0," & RC(FCOL_FX_T0) & ">" & RC(FCOL_AVGENTRYPX) & ")," & "IFERROR(" & "ImpliedRepoBloomberg(" & RC(FCOL_CCY) & "," & RC(FCOL_CONTRACTS) & "," & RC(FCOL_CTD_CF) & "," & RC(FCOL_PORTFOLIO) & "," & _
        RC(FCOL_LINKEDISIN) & "," & RC(FCOL_AVGENTRYPX) & "," & RC(FCOL_FX_T0) & "," & RC(FCOL_FUTPX_T0) & "," & RC(FCOL_FUTURESPNL_EUR) & "," & RC(FCOL_BBG_TICKER) & "),"""")," & """"")")


    With ws.Range(FCOL_IMPLIEDREPO_CALC & f & ":" & FCOL_IMPLIEDREPO_CALC & lastRow)
        .FormulaR1C1 = fml
    End With


    ' -------------------------------------------------------------------------
    ' S = GrossBasis_Calc
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_CONTRACTCODE) & "="""","""","
    fml = fml & "IF(AND(ISNUMBER(" & RC(FCOL_CTD_DIRTYPX_T0) & "),ISNUMBER(" & RC(FCOL_FUTPX_T0) & "),ISNUMBER(" & RC(FCOL_CTD_CF) & ")),"
    fml = fml & RC(FCOL_CTD_DIRTYPX_T0) & "-" & RC(FCOL_FUTPX_T0) & "*" & RC(FCOL_CTD_CF) & ",""""))"
    ws.Range(FCOL_GROSSBASIS_CALC & f & ":" & FCOL_GROSSBASIS_CALC & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' AB = HedgeUnitDV01
    '
    ' Do not write AB here.
    ' AB is now written in WriteFutures_T1_BDP_Efficient using FUT_PX_VAL_BP.
    ' -------------------------------------------------------------------------


    ' -------------------------------------------------------------------------
    ' AC = Futures_DV01_EUR
    '
    ' New physical layout:
    '   AB = RC28 = HedgeUnitDV01 / FUT_PX_VAL_BP
    '   AC = Futures_DV01_EUR
    '
    ' Formula:
    '   Contracts * HedgeUnitDV01 * FX_T0
    '
    ' RC4  = Contracts
    ' RC13 = FX_T0
    ' RC28 = AB = HedgeUnitDV01
    ' -------------------------------------------------------------------------
    ' FUT_VAL_PT is part of the product but was missing from the ISNUMBER
    ' guard, so a future whose BDP(FUT_VAL_PT) came back "" or #N/A made the
    ' whole product an error rather than a blank.  PNL_Attribution wraps the
    ' futures DV01 sum in IFERROR(...,0), so that single bad contract silently
    ' zeroed the futures DV01 of every bond it was mapped to.
    fml = "=IF(" & RC(FCOL_CONTRACTCODE) & "="""","""","
    fml = fml & "IF(AND(ISNUMBER(" & RC(FCOL_CONTRACTS) & "),ISNUMBER(" & RC(FCOL_HEDGEUNITDV01) & "),ISNUMBER(" & RC(FCOL_FX_T0) & "),ISNUMBER(" & RC(FCOL_FUT_VAL_PT) & ")),"
    fml = fml & RC(FCOL_CONTRACTS) & "*" & RC(FCOL_HEDGEUNITDV01) & "*" & RC(FCOL_FX_T0) & "*" & RC(FCOL_FUT_VAL_PT) & ",""""))"
    ws.Range(FCOL_FUTURES_DV01_EUR & f & ":" & FCOL_FUTURES_DV01_EUR & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' T = NotionalValue_EUR
    '
    ' New logic after removing old FUT_CONT_SIZE:
    '
    '   Notional / futures market exposure =
    '       Contracts * FUT_VAL_PT * FutPx_T0 * FX_T0
    '
    ' RC4  = Contracts
    ' RC13 = FX_T0
    ' RC14 = FutPx_T0
    ' RC25 = Y = FUT_VAL_PT
    '
    ' Note:
    '   This is signed if Contracts is signed.
    '   If you want gross notional/exposure, use ABS(RC4) instead of RC4.
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_CONTRACTCODE) & "="""","""","
    fml = fml & "IF(AND(ISNUMBER(" & RC(FCOL_CONTRACTS) & "),ISNUMBER(" & RC(FCOL_FUT_VAL_PT) & "),ISNUMBER(" & RC(FCOL_FUTPX_T0) & "),ISNUMBER(" & RC(FCOL_FX_T0) & ")),"
    fml = fml & RC(FCOL_CONTRACTS) & "*" & RC(FCOL_FUT_VAL_PT) & "*" & RC(FCOL_FUTPX_T0) & "*" & RC(FCOL_FX_T0) & ",""""))"
    ws.Range(FCOL_NOTIONALVALUE_EUR & f & ":" & FCOL_NOTIONALVALUE_EUR & lastRow).FormulaR1C1 = fml


    ' -------------------------------------------------------------------------
    ' U = FuturesPnL_EUR
    '
    ' Formula:
    '   Contracts * FUT_VAL_PT * price change * FX_T0
    '
    ' RC4  = Contracts
    ' RC13 = FX_T0
    ' RC14 = FutPx_T0
    ' RC17 = FutPx_T-1
    ' RC25 = Y = FUT_VAL_PT
    ' -------------------------------------------------------------------------
    fml = "=IF(" & RC(FCOL_CONTRACTCODE) & "="""","""","
    fml = fml & "IF(AND(ISNUMBER(" & RC(FCOL_CONTRACTS) & "),ISNUMBER(" & RC(FCOL_FUT_VAL_PT) & "),ISNUMBER(" & RC(FCOL_FUTPX_T0) & "),ISNUMBER(" & RC(FCOL_FUTPX_TM1) & "),ISNUMBER(" & RC(FCOL_FX_T0) & ")),"
    fml = fml & RC(FCOL_CONTRACTS) & "*" & RC(FCOL_FUT_VAL_PT) & "*(" & RC(FCOL_FUTPX_T0) & "-" & RC(FCOL_FUTPX_TM1) & ")*" & RC(FCOL_FX_T0) & ",""""))"
    ws.Range(FCOL_FUTURESPNL_EUR & f & ":" & FCOL_FUTURESPNL_EUR & lastRow).FormulaR1C1 = fml


End Sub
' =============================================================================
' BONDS T0 WRITER
'   Missing T0 is preserved as ""  (NEVER coerced to 0) so spread deltas blank
'   out instead of producing fake spread PnL.  Spread T0 fields derive from
'   curves only when curve nodes exist, else "".
'   Targets: CleanPx_T0=AD30 DirtyPx_T0=AE31 YTM_T0=AF32 ZSprd_T0=AG33
'            ASW_T0=AH34 ISpread_T0=AN40 GSpread_T0=AP42 OAS_T0=AU47
'            FX_T0=U21 FundRate_T0=BT72   (ticker AR=44)
' =============================================================================


' Bonds T0 DERIVED spreads only.  These reference model cells (YTM_T0, Gov_T0,
' Swap_T0), never Bloomberg, so they are written AFTER the staged T0 values have been
' copied into the model.  The market T0 pulls (U / AD:AH / AU / BT) are produced on the
' hidden staging sheet (WriteBondT0Staging) and copied back as static values.
Private Sub WriteBonds_T0_DerivedSpreads(ws As Worksheet, lastRow As Long)


    If lastRow < BOND_DATA_ROW Then Exit Sub


    ' AL = ISpread_T0 = YTM_T0 - Swap_T0
    ws.Range(BCOL_ISPREAD_TM1 & BOND_DATA_ROW & ":" & BCOL_ISPREAD_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), _
        "IF(AND(ISNUMBER(" & RC(BCOL_YTM_TM1) & "),ISNUMBER(" & RC(BCOL_SWAP_TM1) & ")),(" & RC(BCOL_YTM_TM1) & "-" & RC(BCOL_SWAP_TM1) & ")*100,"""")")


    ' AN = GSpread_T0 = YTM_T0 - Gov_T0
    ws.Range(BCOL_GSPREAD_TM1 & BOND_DATA_ROW & ":" & BCOL_GSPREAD_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), _
        "IF(AND(ISNUMBER(" & RC(BCOL_YTM_TM1) & "),ISNUMBER(" & RC(BCOL_GOV_TM1) & ")),(" & RC(BCOL_YTM_TM1) & "-" & RC(BCOL_GOV_TM1) & ")*100,"""")")


End Sub






' =============================================================================
' BOND PRIOR / T-1 BQL WRITER
'
' Legacy function name: WriteBondT0BQLFormulas
'
' User-facing convention:
'   T0  = current reporting/as-of date
'   T-1 = prior snapshot date
'
' This procedure writes BQL formulas into prior/T-1 model cells.
' =============================================================================




Private Function WriteBondT0BQLFormulas( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long _
) As Range


    If lastRow < BOND_DATA_ROW Then Exit Function


    Dim f As Long
    f = BOND_DATA_ROW


    ' Column targets are the BCOL_ constants, never literal letters.  These
    ' writes used to name "S", "AB", "AS" ... while the matching AddToUnion
    ' below each of them used the constant.  The two agreed only by luck: the
    ' moment a column moved, the formula was written to one column and the
    ' Bloomberg wait registered a different one - so the T-1 pull silently
    ' landed on top of a T0 column and the intended column stayed blank.

    Dim cfgT0 As String
    Dim cfgFxFix As String


    cfgT0 = CfgR1C1(CFG_T0_DATE)
    cfgFxFix = CfgR1C1(CFG_FX_FIX_SOURCE)


    Dim targetRng As Range


    ' -------------------------------------------------------------------------
    ' S = FX_T0
    '
    ' New layout:
    '   Bonds!S = FX_T0
    '
    ' Formula helper:
    '   FXT0FormulaR1C1(cfgFxFix, cfgT0)
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_FX_TM1, f, lastRow, FXT0FormulaR1C1(cfgFxFix, cfgT0)
    AddToUnion targetRng, ws.Range(BCOL_FX_TM1 & f & ":" & BCOL_FX_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AB = CleanPx_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AB = CleanPx_T0
    '
    ' Field:
    '   PX_LAST
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_CLEANPX_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText(BBG_PX_LAST), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_CLEANPX_TM1 & f & ":" & BCOL_CLEANPX_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AC = DirtyPx_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AC = DirtyPx_T0
    '
    ' Field:
    '   PX_DIRTY_MID
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_DIRTYPX_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText(BQL_DIRTY_PX), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_DIRTYPX_TM1 & f & ":" & BCOL_DIRTYPX_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AD = YTM_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AD = YTM_T0
    '
    ' Field:
    '   YIELD(YIELD_TYPE='YTM')
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_YTM_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText("YIELD(YIELD_TYPE='YTM')"), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_YTM_TM1 & f & ":" & BCOL_YTM_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AE = ZSprd_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AE = ZSprd_T0
    '
    ' Field:
    '   SPREAD(SPREAD_TYPE='Z')
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_ZSPRD_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText("SPREAD(SPREAD_TYPE='Z')"), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_ZSPRD_TM1 & f & ":" & BCOL_ZSPRD_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AF = ASW_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AF = ASW_T0
    '
    ' Field:
    '   SPREAD(SPREAD_TYPE='ASW')
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_ASW_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText("SPREAD(SPREAD_TYPE='ASW')"), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_ASW_TM1 & f & ":" & BCOL_ASW_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' AS = OAS_T0
    '
    ' New layout:
    '   Bonds!AP = RC42 = resolved Bloomberg ticker
    '   Bonds!AS = OAS_T0
    '
    ' Field:
    '   SPREAD(SPREAD_TYPE='OAS')
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_OAS_TM1, f, lastRow, _
        BQLPointFormulaR1C1(RC(BCOL_BBG_TICKER), XlText("SPREAD(SPREAD_TYPE='OAS')"), cfgT0)


    AddToUnion targetRng, ws.Range(BCOL_OAS_TM1 & f & ":" & BCOL_OAS_TM1 & lastRow)


    ' -------------------------------------------------------------------------
    ' BR = FundRate_T0
    '
    ' New layout:
    '   Bonds!BP = RC68 = funding ticker
    '   Bonds!BR = FundRate_T0
    '
    ' Field:
    '   PX_LAST
    '
    ' Logic:
    '   Looks back 7 days up to T0 and returns the last available point.
    ' -------------------------------------------------------------------------


    PutBondF ws, BCOL_FUNDRATE_TM1, f, lastRow, _
        "=IF(" & RC(BCOL_ISIN) & "="""","""",IF(" & RC(BCOL_FUNDTKR) & "="""",""""," & _
            BDHLastPointExprR1C1( _
                RC(BCOL_FUNDTKR), _
                XlText(BBG_PX_LAST), _
                "INT(" & cfgT0 & ")-7", _
                "INT(" & cfgT0 & ")" _
            ) & _
        "))"


    AddToUnion targetRng, ws.Range(BCOL_FUNDRATE_TM1 & f & ":" & BCOL_FUNDRATE_TM1 & lastRow)


    Set WriteBondT0BQLFormulas = targetRng


End Function




' =============================================================================
' FUTURES T0 BQL WRITER
'
' Replaces:
'   WriteFutureT0Staging
'
' Writes BQL directly into:
'   Futures!Q = FutPx_T0
'
' Security:
'   Futures!V = RC22 = generic Bloomberg futures ticker
'
' Field:
'   Config B29 = CFG_FUT_PRICE_FIELD
' =============================================================================


Private Function WriteFutureT0BQLFormulas( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long _
) As Range


    If lastRow < DATA_ROW Then Exit Function


    Dim cfgT0 As String
    Dim cfgFutPx As String


    cfgT0 = CfgR1C1(CFG_T0_DATE)
    cfgFutPx = CfgR1C1(CFG_FUT_PRICE_FIELD)


    Dim rg As Range
    Set rg = ws.Range(FCOL_FUTPX_TM1 & DATA_ROW & ":" & FCOL_FUTPX_TM1 & lastRow)


    rg.FormulaR1C1 = _
        "=IF(" & RC(FCOL_BBG_TICKER) & "="""",""""," & _
            BQLPointExprR1C1(RC(FCOL_BBG_TICKER), cfgFutPx, cfgT0) & _
        ")"


    Set WriteFutureT0BQLFormulas = rg
End Function






' =============================================================================
' SWAPS FORMULAS  (unchanged economics; PnL only when OPICS supplies economics)
' =============================================================================


Private Sub WriteSwapsCalculatedFormulas(ws As Worksheet)


    Dim r As Long
    Dim lastRow As Long


    lastRow = LastSwapDataRow(ws)
    If lastRow < DATA_ROW Then Exit Sub


    For r = DATA_ROW To lastRow


        If CleanText(ws.Cells(r, WCOL_DEALID).value) <> "" Then


            ' -----------------------------------------------------------------
            ' K = explicit front-table curve identifier.
            '
            ' Source:
            '   BA = PAY_FLT_RATE_IDX from Bloomberg
            '   E  = fallback OPICS/source FloatIndex
            '
            ' Output:
            '   ESTR / EURIBOR / UNKNOWN
            
            '
            ' Currency conventions:
            '   EUR OIS = ESTR
            '   USD OIS = SOFR
            '
            ' UNKNOWN is retained when the floating index cannot be classified.


            ws.Cells(r, WCOL_FLOATCURVE_TYPE).formula = _
                "=IF($" & WCOL_DEALID & r & "="""","""",SwapFloatFamily(IF($" & WCOL_PAY_FLT_RATE_IDX & r & "="""",$" & WCOL_FLOATINDEX & r & ",$" & WCOL_PAY_FLT_RATE_IDX & r & "),$" & WCOL_CCY & r & "))"


            ' -----------------------------------------------------------------
            ' Time / OIS / discount / annuity / FX block.
            ' -----------------------------------------------------------------
            ws.Cells(r, WCOL_YEARFRAC).formula = _
                SwapYearFracFml("$" & WCOL_DEALID & r, "$" & WCOL_ENDDATE & r)


            ws.Cells(r, WCOL_OIS_T0).formula = _
                SwapOisRateFml("$" & WCOL_CCY & r, "$" & WCOL_YEARFRAC & r, True)


            ws.Cells(r, WCOL_OIS_TM1).formula = _
                SwapOisRateFml("$" & WCOL_CCY & r, "$" & WCOL_YEARFRAC & r, False)


            ws.Cells(r, WCOL_DF_T0).formula = _
                DiscountFactorFml(WCOL_OIS_T0 & r, WCOL_YEARFRAC & r)


            ws.Cells(r, WCOL_DF_TM1).formula = _
                DiscountFactorFml(WCOL_OIS_TM1 & r, WCOL_YEARFRAC & r)


            ws.Cells(r, WCOL_ANNUITY_T0).formula = _
                AnnuityFml(WCOL_DF_T0 & r, WCOL_OIS_T0 & r)


            ws.Cells(r, WCOL_ANNUITY_TM1).formula = _
                AnnuityFml(WCOL_DF_TM1 & r, WCOL_OIS_TM1 & r)


            ws.Cells(r, WCOL_FX).formula = _
                FxToBaseFml("$" & WCOL_DEALID & r, "$" & WCOL_CCY & r)


            ' -----------------------------------------------------------------
            ' Float-index diagnostics.
            ' -----------------------------------------------------------------
            ws.Cells(r, WCOL_FLOATINDEX_FAMILY).formula = _
                "=IF($" & WCOL_DEALID & r & "="""",""""," & WCOL_FLOATCURVE_TYPE & r & ")"
                
            ws.Cells(r, WCOL_FLOATFAMILY_STATUS).formula = _
                "=IF($" & WCOL_DEALID & r & "="""",""""," & _
                    "SwapFloatFamilyStatus(" & _
                        "IF($" & WCOL_PAY_FLT_RATE_IDX & r & "=""""," & _
                            "$" & WCOL_FLOATINDEX & r & "," & _
                            "$" & WCOL_PAY_FLT_RATE_IDX & r & ")," & _
                        "$" & WCOL_CCY & r & "," & _
                        "$" & WCOL_FLOATINDEX_FAMILY & r & "))"


            ws.Cells(r, WCOL_FLOATINDEX_TENOR).formula = _
                "=IF($" & WCOL_DEALID & r & "="""","""",SwapFloatTenor(IF($" & WCOL_PAY_FLT_RATE_IDX & r & "="""",$" & WCOL_FLOATINDEX & r & ",$" & WCOL_PAY_FLT_RATE_IDX & r & "),$" & WCOL_FLOATINDEX_FAMILY & r & "))"


            ' BD = selected floating reference curve, current T0.
            ws.Cells(r, WCOL_FLOATCURVE_T0).formula = _
                "=IF($" & WCOL_DEALID & r & "="""",""""," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""ESTR""," & _
                        WCOL_OIS_T0 & r & "," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""SOFR""," & _
                        WCOL_OIS_T0 & r & "," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""EURIBOR""," & _
                        "InterpSwap($" & WCOL_CCY & r & ",$" & _
                        WCOL_YEARFRAC & r & ",TRUE)," & _
                    """""))))"


            ' BE = selected floating reference curve, prior T-1.
            ws.Cells(r, WCOL_FLOATCURVE_TM1).formula = _
                "=IF($" & WCOL_DEALID & r & "="""",""""," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""ESTR""," & _
                        WCOL_OIS_TM1 & r & "," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""SOFR""," & _
                        WCOL_OIS_TM1 & r & "," & _
                    "IF($" & WCOL_FLOATINDEX_FAMILY & r & "=""EURIBOR""," & _
                        "InterpSwap($" & WCOL_CCY & r & ",$" & _
                        WCOL_YEARFRAC & r & ",FALSE)," & _
                    """""))))"


            ' BF/BG = model spread versus selected floating reference curve.
            ws.Cells(r, WCOL_MODELSPREAD_T0).formula = _
                "=IF($" & WCOL_DEALID & r & "="""","""",IF(AND(ISNUMBER($" & WCOL_FIXEDRATE & r & "),ISNUMBER($" & WCOL_FLOATSPREAD & r & "),ISNUMBER($" & WCOL_FLOATCURVE_T0 & r & ")),$" & WCOL_FIXEDRATE & r & "+$" & WCOL_FLOATSPREAD & r & "-$" & WCOL_FLOATCURVE_T0 & r & ",""""))"


            ws.Cells(r, WCOL_MODELSPREAD_TM1).formula = _
                "=IF($" & WCOL_DEALID & r & "="""","""",IF(AND(ISNUMBER($" & WCOL_FIXEDRATE & r & "),ISNUMBER($" & WCOL_FLOATSPREAD & r & "),ISNUMBER($" & WCOL_FLOATCURVE_TM1 & r & ")),$" & WCOL_FIXEDRATE & r & "+$" & WCOL_FLOATSPREAD & r & "-$" & WCOL_FLOATCURVE_TM1 & r & ",""""))"


            ws.Cells(r, WCOL_DELTA_FLOATCURVE_BP).formula = _
                "=IF(AND(ISNUMBER(" & WCOL_FLOATCURVE_T0 & r & "),ISNUMBER(" & WCOL_FLOATCURVE_TM1 & r & ")),(" & WCOL_FLOATCURVE_T0 & r & "-" & WCOL_FLOATCURVE_TM1 & r & ")*100,"""")"


            ws.Cells(r, WCOL_DELTA_MODELSPREAD_BP).formula = _
                "=IF(AND(ISNUMBER(" & WCOL_MODELSPREAD_T0 & r & "),ISNUMBER(" & WCOL_MODELSPREAD_TM1 & r & ")),(" & WCOL_MODELSPREAD_T0 & r & "-" & WCOL_MODELSPREAD_TM1 & r & ")*100,"""")"


            ' -----------------------------------------------------------------
            ' Model PVs and model PnL.
            ' -----------------------------------------------------------------
            ws.Cells(r, WCOL_PV_T0_MODEL).formula = _
                SwapModelPVFml("$" & WCOL_CCY & r, "$" & WCOL_NOTIONAL & r, "$" & WCOL_FLOATCURVE_T0 & r, _
                               "$" & WCOL_PAYFIXED & r, "$" & WCOL_MODELSPREAD_T0 & r, "$" & WCOL_ANNUITY_T0 & r, "$" & WCOL_FX & r)


            ws.Cells(r, WCOL_PV_TM1_MODEL).formula = _
                SwapModelPVFml("$" & WCOL_CCY & r, "$" & WCOL_NOTIONAL & r, "$" & WCOL_FLOATCURVE_TM1 & r, _
                               "$" & WCOL_PAYFIXED & r, "$" & WCOL_MODELSPREAD_TM1 & r, "$" & WCOL_ANNUITY_TM1 & r, "$" & WCOL_FX & r)




            ws.Cells(r, WCOL_SWAP_DV01_EUR).formula = _
                SwapDV01Fml("$" & WCOL_CCY & r, "$" & WCOL_NOTIONAL & r, "$" & WCOL_PAYFIXED & r, "$" & WCOL_ANNUITY_T0 & r, "$" & WCOL_FX & r)




            ws.Cells(r, WCOL_MODEL_PNL).formula = _
                ModelPnLFml(WCOL_PV_T0_MODEL & r, WCOL_PV_TM1_MODEL & r)


            ' -----------------------------------------------------------------
            ' Status.
            ' -----------------------------------------------------------------
            ws.Cells(r, WCOL_STATUS).formula = _
                "=IF(OR($" & WCOL_CCY & r & "="""",$" & WCOL_NOTIONAL & r & "=""""),""Missing economics""," & _
                "IF(AND(UPPER(TRIM($" & WCOL_CCY & r & "))=""EUR"",$" & WCOL_FLOATCURVE_TYPE & r & "=""UNKNOWN""),""Unknown float index""," & _
                "IF(OR($" & WCOL_FLOATCURVE_T0 & r & "="""",$" & WCOL_FLOATCURVE_TM1 & r & "=""""),""Missing float curve"",""OK"")))"


            ws.Cells(r, WCOL_PNL_SOURCE).formula = _
                "=IF(" & WCOL_STATUS & r & "=" & Chr(34) & "OK" & Chr(34) & ",""MODEL"",""NONE"")"


            ' -----------------------------------------------------------------
            ' Final swap PnL and final status.
            '
            ' AB = final swap PnL used by PNL_Attribution.
            ' AC = final swap status.
            ' -----------------------------------------------------------------
            ws.Cells(r, WCOL_PNL).formula = _
                "=IF(ISNUMBER(" & WCOL_BQL_SWAP_PNL & r & ")," & WCOL_BQL_SWAP_PNL & r & ",IF(ISNUMBER(" & WCOL_MODEL_PNL & r & ")," & WCOL_MODEL_PNL & r & ",""""))"


            ws.Cells(r, WCOL_AF_STATUS).formula = _
                "=IF(" & WCOL_BQL_SWAP_STATUS & r & "<>""""," & WCOL_BQL_SWAP_STATUS & r & "," & WCOL_STATUS & r & ")"


        End If


    Next r


End Sub


' =============================================================================
' PNL ATTRIBUTION
'   Full header + row rebuild removes any stale #REF! formulas in the sheet.
'   Run RebuildPNLOnly for a fast repair without touching Bloomberg.
' =============================================================================


Public Sub RebuildPNLOnly()


    Dim wb As Workbook
    Set wb = ThisWorkbook


    Dim wsBnd As Worksheet
    Dim wsPnl As Worksheet


    Set wsBnd = wb.Worksheets(SH_BONDS)
    Set wsPnl = wb.Worksheets(SH_PNL)


    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldStatusBar As Variant


    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldStatusBar = Application.StatusBar


    On Error GoTo CleanFail


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Rebuilding PnL attribution only..."


    Dim tStart As Double
    tStart = Timer


    Dim lastBondRow As Long
    lastBondRow = LastBondDataRow(wsBnd)


    SetupPNLAttributionHeaders wsPnl


    If lastBondRow >= BOND_DATA_ROW Then
        WritePNLAttributionRows wsPnl, wsBnd
        wsPnl.Range(PCOL_ISIN & DATA_ROW & ":" & PNL_LAST_COL & _
        CStr(DATA_ROW + PnlRowCount(lastBondRow) - 1)).Calculate
    End If


    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar


    MsgBox "PNL attribution rebuilt." & vbCrLf & _
           "Elapsed seconds: " & Format(Timer - tStart, "0.0"), vbInformation


    Exit Sub


CleanFail:
    Dim eInfo As String
    eInfo = CaptureErrInfo()


    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar


    MsgBox "RebuildPNLOnly failed:" & vbCrLf & eInfo, vbCritical


End Sub


' =============================================================================
' THE PNL_ATTRIBUTION COLUMN CONTRACT  -  one table, three consumers
'
' Every column of PNL_Attribution is declared exactly once, here, as three
' things that used to be scattered across three procedures that could disagree
' with each other:
'
'   COLUMN   where it sits            (the PCOL_ constant)
'   KEY      what the Dashboard asks for   (stable, machine-facing, never
'                                          translated, never reworded)
'   LABEL    what the desk reads in row 4  (yours to change freely)
'
' KEY and LABEL are separated on purpose, and it is the single most important
' thing in this module to understand before touching the sheet.
'
' The Dashboard used to find its columns by SEARCHING row 4 for the header
' text - Rows(4).Find("Bond_DV01_Current").  So the header row was not a row of
' labels at all: it was the interface between the two modules, and every string
' in it was load-bearing.  Renaming L4 to "Bond BPVs" - an obviously harmless
' thing to do to a spreadsheet, and the right name for the desk to read - meant
' the search returned Nothing and the whole Dashboard build died on error 9901
' before drawing a single cell.  There was no way to tell from the sheet that
' this would happen.
'
' Now nothing searches for header text.  PublishPnlColumnNames walks this table
' and publishes one workbook name per KEY - Pnl_Bond_DV01_Current and so on -
' pointing at that column by LETTER.  The Dashboard refers only to those names.
' Row 4 is free to say whatever the desk wants it to say, in any language.
'
' TO CHANGE A COLUMN, EDIT THIS TABLE AND NOTHING ELSE.  See docs/COLUMNS.md.
' =============================================================================

Private Sub AddPnlCol( _
    ByRef spec As Collection, _
    ByVal colLetter As String, _
    ByVal contractKey As String, _
    ByVal displayLabel As String)

    spec.Add Array(colLetter, contractKey, displayLabel)

End Sub


' The layout, in sheet order.  Adding a row here adds the column everywhere:
' the header writer writes it, the contract check verifies it, and the
' Dashboard can name it.  Writing the VALUE into the column is still a
' separate job - see WritePNLRow.
Private Function PnlLayout() As Collection

    Dim spec As Collection
    Set spec = New Collection

    ' --- identity, period and the bond's own risk measures ---
    AddPnlCol spec, PCOL_ISIN, "ISIN", "ISIN"
    AddPnlCol spec, PCOL_NAME, "Name", "Name"
    AddPnlCol spec, PCOL_CCY, "CCY", "CCY"
    AddPnlCol spec, PCOL_PORTFOLIO, "Portfolio", "Portfolio"
    AddPnlCol spec, PCOL_ACCTGCAT, "AcctgCat", "AcctgCat"
    AddPnlCol spec, PCOL_NOTIONAL, "Notional", "Notional"
    AddPnlCol spec, PCOL_MODDUR, "ModDur", "ModDur"
    AddPnlCol spec, PCOL_CONVEXITY, "Convexity", "Convexity"
    AddPnlCol spec, PCOL_SPREADDURATION, "SpreadDuration", "SpreadDuration"
    AddPnlCol spec, PCOL_DAYS, "Days", "Days"
    AddPnlCol spec, PCOL_YEARFRAC, "YearFrac", "YearFrac"

    ' --- DV01 / BPV block.  Futures risk is split by SOURCE BOOK, not by
    ' instrument: the desk's futures hedges live in two files, so a book whose
    ' futures hedge stops arriving from one of them is visible rather than netted
    ' away inside a single total. ---
    AddPnlCol spec, PCOL_BOND_DV01_CURRENT, "Bond_DV01_Current", "Bond BPVs"
    AddPnlCol spec, PCOL_BOND_DV01_CREDIT_SPREAD, "Bond_DV01_Credit_Spread", "Bond DV01 Credit Spread"
    AddPnlCol spec, PCOL_HEDGE_DV01, "Actual_Hedge_DV01", "Hedge (BPVs)"
    AddPnlCol spec, PCOL_PLAINSWAP_DV01, "PlainSwap_DV01", "PlainSwap (BPVs)"
    AddPnlCol spec, PCOL_FUTURES_RTJ_DV01, "FuturesRTJ_DV01", "Futures RTJ (BPVs)"
    AddPnlCol spec, PCOL_FUTURES_RT_DV01, "FuturesRT_DV01", "Futures RT (BPVs)"
    AddPnlCol spec, PCOL_HEDGE_DV01_GAP, "Hedge_DV01_Gap", "Hedge_DV01_Gap"
    AddPnlCol spec, PCOL_SYNTHETICSWAP_DV01, "SyntheticSwap_DV01", "SyntheticSwap_DV01"
    AddPnlCol spec, PCOL_TARGET_HEDGE_DV01, "Target_Hedge_DV01", "Target_Hedge_DV01"

    ' --- market movements, in basis points ---
    AddPnlCol spec, PCOL_DELTA_DIRTY_MV_EUR, "Delta_Dirty_MV_EUR", "VariaÃ§Ã£o Bond (EUR)"
    AddPnlCol spec, PCOL_DELTA_Y_BP, "Delta_Y_bp", "VariaÃ§Ã£o yield bond (bps)"
    AddPnlCol spec, PCOL_DELTA_R_BP, "Delta_r_bp", "VariaÃ§Ã£o yield rf (bps)"
    AddPnlCol spec, PCOL_DELTA_GOV_BP, "Delta_Gov_bp", "VariaÃ§Ã£o yield Gov (bps)"
    AddPnlCol spec, PCOL_DELTA_G_BP, "Delta_g_bp", "VariaÃ§Ã£o Govi basis bp"
    AddPnlCol spec, PCOL_DELTA_SWAP_BP, "Delta_Swap_bp", "VariaÃ§Ã£o Swap bp"
    AddPnlCol spec, PCOL_DELTA_Q_BP, "Delta_q_bp", "VariaÃ§Ã£o Gov/Swap basis bp"
    AddPnlCol spec, PCOL_DELTA_I_BP, "Delta_i_bp", "VariaÃ§Ã£o I spread bp"
    AddPnlCol spec, PCOL_DELTA_Z_BP, "Delta_Z_bp", "VariaÃ§Ã£o Z spread bp"
    AddPnlCol spec, PCOL_DELTA_G_BP_T, "Delta_GSpread_bp", "VariaÃ§Ã£o G spread bp"
    AddPnlCol spec, PCOL_DELTA_ASW_BP, "Delta_ASW_bp", "VariaÃ§Ã£o ASW bp"
    AddPnlCol spec, PCOL_DELTA_OAS_BP, "Delta_OAS_bp", "VariaÃ§Ã£o OAS bp"

    ' --- main duration PnL chain ---
    AddPnlCol spec, PCOL_PNL_DURATION_TOTAL, "PnL_Duration_Total", "PnL_Duration_Total"
    AddPnlCol spec, PCOL_PNL_OIS, "PnL_OIS", "PnL_OIS"
    AddPnlCol spec, PCOL_PNL_GOVBASIS, "PnL_GovBasis", "PnL_GovBasis"
    AddPnlCol spec, PCOL_PNL_SWAPGOVBASIS, "PnL_SwapGovBasis", "PnL_SwapGovBasis"
    AddPnlCol spec, PCOL_PNL_CREDIT_ISPREAD, "PnL_Credit_Ispread", "PnL_Credit_Ispread"
    AddPnlCol spec, PCOL_PNL_CONVEXITY, "PnL_Convexity", "PnL_Convexity"

    ' --- carry, alternative spread PnLs and FX ---
    AddPnlCol spec, PCOL_CARRY_COUPON, "Carry_Coupon", "Carry_Coupon"
    AddPnlCol spec, PCOL_CARRY_ROLLTOPAR, "Carry_RollToPar", "Carry_RollToPar"
    AddPnlCol spec, PCOL_CARRY_FUNDING, "Funding_Carry_Memo", "Funding_Carry_Memo"
    AddPnlCol spec, PCOL_CARRY_TOTAL, "Carry_Total", "Carry_Total"
    AddPnlCol spec, PCOL_PNL_ZSPREAD, "PnL_ZSpread", "PnL_ZSpread"
    AddPnlCol spec, PCOL_PNL_GSPREAD, "PnL_GSpread", "PnL_GSpread"
    AddPnlCol spec, PCOL_PNL_ASW, "PnL_ASW", "PnL_ASW"
    AddPnlCol spec, PCOL_PNL_OAS, "PnL_OAS", "PnL_OAS"
    AddPnlCol spec, PCOL_SPREADPNL_USED, "SpreadPnL_Used", "SpreadPnL_Used"
    AddPnlCol spec, PCOL_PNL_FX, "PnL_FX", "PnL_FX"

    ' --- hedge model PnL, explained and residual ---
    AddPnlCol spec, PCOL_PNL_FUTURES, "Futures_Gov_Model_PnL", "Futures_Gov_Model_PnL"
    AddPnlCol spec, PCOL_PNL_SWAP, "Swap_Curve_Model_PnL", "Swap_Curve_Model_PnL"
    AddPnlCol spec, PCOL_TOTAL_HEDGE, "Hedge_Curve_Model_PnL", "Hedge_Curve_Model_PnL"
    AddPnlCol spec, PCOL_BASIS_PNL, "Hedge_Model_Residual_PnL", "Hedge_Model_Residual_PnL"
    AddPnlCol spec, PCOL_TOTAL_EXPLAINED, "Total_Model_Explained", "Total_Model_Explained"
    AddPnlCol spec, PCOL_OFFICIAL_PNL, "Official_Total_PnL", "Official_Total_PnL"
    AddPnlCol spec, PCOL_RESIDUAL, "Unexplained_Residual_PnL", "Unexplained_Residual_PnL"
    AddPnlCol spec, PCOL_RESIDUAL_AX, "Unexplained_Residual_Pct", "Unexplained_Residual_Pct"

    ' --- hedge diagnostics ---
    AddPnlCol spec, PCOL_RESIDUAL_DV01, "Residual_DV01", "Residual_DV01"
    AddPnlCol spec, PCOL_HEDGE_RATIO, "Hedge_Ratio", "Hedge_Ratio"
    AddPnlCol spec, PCOL_HEDGE_EFFICIENCY, "Hedge_Efficiency", "Hedge_Efficiency"
    AddPnlCol spec, PCOL_ATTRIBUTION_STATUS, "Attribution_Status", "Attribution_Status"
    AddPnlCol spec, PCOL_THEORETICALHEDGE_DV01, "Synthetic_Alternative_DV01", "Synthetic_Alternative_DV01"

    ' --- matching counts, actual hedge PnL and framework diagnostics ---
    AddPnlCol spec, PCOL_FUT_MATCH_COUNT, "Futures_Match_Count", "Futures_Match_Count"
    AddPnlCol spec, PCOL_FUT_PNL_RAW, "Actual_Futures_PnL", "Actual_Futures_PnL"
    AddPnlCol spec, PCOL_SWAP_ALL_MATCH_COUNT, "Swap_All_Match_Count", "Swap_All_Match_Count"
    AddPnlCol spec, PCOL_SWAP_PLAIN_MATCH_COUNT, "PlainSwap_Match_Count", "PlainSwap_Match_Count"
    AddPnlCol spec, PCOL_SWAP_PLAIN_PNL_RAW, "Actual_PlainSwap_PnL", "Actual_PlainSwap_PnL"
    AddPnlCol spec, PCOL_SWAP_SYNTHETIC_MATCH_COUNT, "SyntheticSwap_Match_Count", "SyntheticSwap_Match_Count"
    AddPnlCol spec, PCOL_SWAP_SYNTHETIC_PNL_RAW, "Actual_SyntheticSwap_PnL", "Actual_SyntheticSwap_PnL"
    AddPnlCol spec, PCOL_SWAP_ALL_PNL_RAW, "AllSwapRows_PnL_Diagnostic", "AllSwapRows_PnL_Diagnostic"
    AddPnlCol spec, PCOL_ACTUAL_FUTURES_PNL_RAW, "Actual_Futures_PnL_Check", "Actual_Futures_PnL_Check"
    AddPnlCol spec, PCOL_FUTURES_GOV_MODEL_PNL, "Futures_Gov_Model_PnL_Check", "Futures_Gov_Model_PnL_Check"
    AddPnlCol spec, PCOL_FUTURES_BASIS_PNL, "Futures_Model_Residual_PnL", "Futures_Model_Residual_PnL"
    AddPnlCol spec, PCOL_ACTUAL_SWAP_PNL_RAW, "Actual_PlainSwap_PnL_Check", "Actual_PlainSwap_PnL_Check"
    AddPnlCol spec, PCOL_SWAP_CURVE_MODEL_PNL, "Swap_Curve_Model_PnL_Check", "Swap_Curve_Model_PnL_Check"
    AddPnlCol spec, PCOL_SWAP_BASIS_PNL, "Swap_Model_Residual_PnL", "Swap_Model_Residual_PnL"
    AddPnlCol spec, PCOL_ACTUAL_HEDGE_PNL_RAW, "Actual_Hedge_PnL", "Actual_Hedge_PnL"
    AddPnlCol spec, PCOL_HEDGE_BASIS_PNL, "Hedge_Model_Residual_PnL_Check", "Hedge_Model_Residual_PnL_Check"
    AddPnlCol spec, PCOL_SPREAD_FRAMEWORK_AUTO, "Spread_Framework_Auto", "Spread_Framework_Auto"
    AddPnlCol spec, PCOL_SPREAD_FRAMEWORK_REASON, "Spread_Framework_Reason", "Spread_Framework_Reason"
    AddPnlCol spec, PCOL_DURATION_IDENTITY_CHECK, "Duration_Identity_Check", "Duration_Identity_Check"

    ' --- row quarantine, FX exposure and the opening-risk anchor ---
    AddPnlCol spec, PCOL_ROW_VALID, "Row_Valid", "Row_Valid"
    AddPnlCol spec, PCOL_ROW_EXCLUSION_REASON, "Row_Exclusion_Reason", "Row_Exclusion_Reason"
    AddPnlCol spec, PCOL_FX_EXPOSURE_EUR, "FX_Exposure_EUR", "FX_Exposure_EUR"
    AddPnlCol spec, PCOL_BOND_DV01_OPENING, "Bond_DV01_Opening", "Bond_DV01_Opening"
    AddPnlCol spec, PCOL_RISK_TIMING_BIAS, "Risk_Timing_Bias", "Risk_Timing_Bias"
    AddPnlCol spec, PCOL_COUPON_PAID_EUR, "Coupon_Paid_EUR", "Coupon_Paid_EUR"

    Set PnlLayout = spec

End Function


Private Sub SetupPNLAttributionHeaders(ws As Worksheet)

    ' Clear as far as this sheet has ever been written, not to a fixed 600 rows:
    ' the bond count changes every load and stale rows below the new last bond
    ' would keep reporting positions the desk no longer holds.
    Dim clearLastRow As Long
    clearLastRow = SheetClearLastRow(ws, DATA_ROW, DATA_ROW - 1)

    ' -------------------------------------------------------------------------
    ' Preserve formatting:
    '   - Clear only cell contents/formulas.
    '   - Do NOT clear formats.
    '   - Do NOT apply font/alignment/autofit here.
    '
    ' PNL_LAST_COL is the active layout end.
    ' PNL_CLEAR_LAST_COL reaches past it so a column REMOVED from the layout
    ' cannot leave a stale header and a column of stale values behind it -
    ' which is exactly what happened to the three DV01 columns dropped here.
    ' -------------------------------------------------------------------------

    ws.Range(CELL_PNL_A1).value = "PNL Attribution"

    ' Clear old headers/output values/formulas, including old unused columns.
    ' Formatting is preserved.
    ws.Range(PCOL_ISIN & "3" & ":" & _
             PNL_CLEAR_LAST_COL & clearLastRow).ClearContents

    ' Row 4 carries the DISPLAY LABEL from PnlLayout, not the contract key.
    ' Nothing downstream reads these strings; PublishPnlColumnNames publishes
    ' the keys as workbook names against the same table.  Reword them freely.
    Dim spec As Collection
    Dim item As Variant

    Set spec = PnlLayout()

    For Each item In spec
        ws.Range(CStr(item(0)) & PNL_HEADER_ROW).value = CStr(item(2))
    Next item

End Sub


' =============================================================================
' PublishPnlColumnNames  -  the interface the Dashboard actually uses
'
' One workbook name per contract key, covering rows DATA_ROW..lastRow of that
' key's column:  Pnl_Bond_DV01_Current  ->  'PNL_Attribution'!$L$5:$L$212
'
' Public because modDashboard calls it before it writes a single formula.  The
' Dashboard has no other way to find a column, which is the point: a name either
' exists and is right, or does not exist and fails loudly at the one call site
' that creates it, rather than silently resolving to whatever text happens to
' sit in row 4 today.
'
' Republished on every run, because the row count changes with the book.
' =============================================================================
Public Sub PublishPnlColumnNames(ByVal lastRow As Long)

    Dim spec As Collection
    Dim item As Variant
    Dim colLetter As String
    Dim refersTo As String

    If lastRow < DATA_ROW Then lastRow = DATA_ROW

    Set spec = PnlLayout()

    For Each item In spec

        colLetter = CStr(item(0))

        refersTo = "='" & SH_PNL & "'!$" & colLetter & "$" & CStr(DATA_ROW) & _
                   ":$" & colLetter & "$" & CStr(lastRow)

        On Error Resume Next
        ThisWorkbook.Names(PNL_NAME_PREFIX & CStr(item(1))).Delete
        On Error GoTo 0

        ThisWorkbook.Names.Add _
            Name:=PNL_NAME_PREFIX & CStr(item(1)), _
            refersTo:=refersTo

    Next item

End Sub


Private Sub WritePNLAttributionRows(wsPnl As Worksheet, wsBnd As Worksheet)


    Dim lastBondRow As Long
    Dim bondRow As Long
    Dim pnlRow As Long
    Dim clearLastRow As Long


    lastBondRow = LastBondDataRow(wsBnd)


    ' Size the hedge lookup ranges from the hedge sheets themselves.  Reaching a
    ' little past the last live hedge (SheetClearLastRow) is deliberate: SUMIFS
    ' over trailing blanks is harmless, missing a hedge is not.
    SetHedgeLookupExtents


    ' PNL_Attribution still uses:
    '   Row 4 = headers
    '   Row 5 = first data row
    ' Active final column is CG.
    clearLastRow = SheetClearLastRow(wsPnl, DATA_ROW, DATA_ROW - 1)
    wsPnl.Range(PCOL_ISIN & DATA_ROW & ":" & PNL_CLEAR_LAST_COL & clearLastRow).ClearContents


    pnlRow = DATA_ROW


    ' Bonds layout: rows 1..BOND_COMMENT_ROWS are the desk's reserved comment
    ' band, BOND_HEADER_ROW is the header, BOND_DATA_ROW is the first bond.
    For bondRow = BOND_DATA_ROW To lastBondRow
        If CleanText(wsBnd.Cells(bondRow, colNum(BCOL_ISIN)).value) <> "" Then
            WritePNLRow wsPnl, bondRow, pnlRow
            pnlRow = pnlRow + 1
        End If
    Next bondRow


End Sub


' =============================================================================
' WritePNLRow  -  one PNL_Attribution row per BOND position
'
' Every cell written here is an Excel formula, never a value, so the sheet stays
' auditable and recalculates without re-running the macro.  Missing inputs
' produce "" rather than 0, so a gap never masquerades as a zero PnL.
'
' THE BRIDGE
'   AV Official_Total_PnL  = change in dirty MV + ACTUAL hedge PnL
'   AU Total_Model_Explained
'                          = AA Duration      (framework chain, ties to -DV01*dy)
'                          + AF Convexity
'                          + AJ Carry         (coupon + roll-to-par)
'                          + AP FX
'                          + AS Hedge model PnL
'                          + AT Hedge basis   (actual hedge - model hedge)
'   AW Unexplained_Residual_PnL = AV - AU
'
' AS + AT telescopes to the actual hedge PnL, which is what AV contains, so the
' bridge closes by construction and AW isolates bond-leg model error.
'
' SIGN CONVENTION
'   PositionSide (Bonds!L, PNL!H) is the ONLY +/-1 driver and is always paired
'   with ABS(Notional).  Bonds!AG DV01_EUR and Bonds!AH/AI DirtyMV are already
'   signed by it - do not multiply them by PositionSide again.
'
' KNOWN LIMITATIONS  (real, deliberate, and NOT bugs to be "fixed" locally)
'
'   1. No T-1 position snapshot.  Positions come from a single OPICS pull of
'      TODAY's book (GetBondSelectQuery), and DirtyMV_T-1 is today's notional at
'      yesterday's price.  Therefore:
'        - a position opened during the period has its whole change in MV
'          attributed to market moves, since there is no entry-price term;
'        - a position closed during the period is absent from the query and its
'          realised PnL is silently missing from every total;
'        - an intraday size change is mis-attributed in proportion to the change.
'      Fixing this needs a second, dated position pull - not a formula change.
'
'   2. No coupon-payment handling.  If a coupon pays inside the period, accrued
'      interest resets and the change in dirty MV drops by the coupon, producing
'      a large one-day residual on that bond.  Carry_Coupon accrues smoothly and
'      does not offset it.
'
'   3. Hedges are matched to bonds by ISIN ONLY (Futures!J and Swaps!L
'      LinkedISIN).  Consequences:
'        - a future or swap with a blank LinkedISIN appears on NO row here, so
'          its PnL is absent from every PNL_Attribution total;
'        - if two bonds carry the same hedge ISIN, each row's SUMIFS claims the
'          FULL hedge PnL and DV01, double counting it.
'      The Dashboard surfaces both as explicit reconciliation figures.
'
' TWO UPSTREAM CONVENTIONS THIS SHEET DEPENDS ON
'
' Both were tested by rebuilding this workbook's formulas in a spreadsheet
' engine and running the wrong convention on purpose.  Neither fails silently -
' each one lights up Attribution_Status - so the desk check is quick.
'
'   a. Futures!AC Futures_DV01_EUR = Contracts * HedgeUnitDV01 * FX_T0 * FUT_VAL_PT
'
'      HedgeUnitDV01 is BDP FUT_PX_VAL_BP, which Bloomberg returns in PRICE
'      POINTS per bp (~0.0605 for a Bund), not in cash.  Multiplying by
'      FUT_VAL_PT (EUR 1000 per full point) is therefore CORRECT and gives
'      ~60.5 EUR per contract per bp.  The formula is right; only the comment
'      above it was incomplete.
'
'      Desk check: Futures!AC / Futures!D should be roughly 60-90 for a Bund.
'      If FUT_PX_VAL_BP ever came back already in cash the result would be
'      1000x too large, hedge ratios would run into the hundreds and every
'      futures-hedged bond would read "Over-hedged" - measured at ratio 563
'      against a correct 0.56 on the same test position.
'
'   b. Swaps!BN DV01_BBG is raw Bloomberg SW_CNV_BPV with NO sign normalisation,
'      and the internal model DV01 (Swaps!X) is deliberately not used.  AZ
'      Residual_DV01 = W + AY only nets if BN is SIGNED opposite to the bond.
'
'      Desk check: if BN comes back as a magnitude instead, the hedge ratio
'      flips sign and every hedged bond reads "Wrong-direction hedge" - measured
'      at -0.84 against a correct +0.84 on the same test position.  If that is
'      what you see, normalise BN by PayFixed at the source rather than here.
'
' RISK IS STRUCK AT T0, NOT AT T-1  (known bias, quantified)
'
'   Bonds!AU DV01_Unit is built from DirtyPx_T0, so the DV01 used to explain the
'   move from T-1 to T0 is the risk as it stands AFTER that move.  Attribution
'   convention is to strike risk at the START of the period being explained.
'
'   The bias is first order in the size of the move.  On a self-consistent test
'   book - prices derived from the yield move so the bridge SHOULD close exactly
'   - a 10bp day left a residual of +288 EUR on a -66,170 EUR duration PnL, and
'   +278 of that 288 (96.5%) was this timing convention alone.  Scale it up and
'   a 50bp day carries roughly a 3% standing residual that is pure convention,
'   not model error.
'
'   Fixing it means introducing a separate attribution DV01 struck on
'   DirtyPx_T-1.  DV01_EUR itself should stay as is: hedging needs CURRENT risk,
'   and the hedge ratio and efficiency columns rightly use it.  This is a
'   deliberate open item, not an oversight.
' =============================================================================
' Measure the Futures and Swaps sheets once, for the lookup ranges baked into
' every PNL_Attribution row.
Private Sub SetHedgeLookupExtents()

    Dim wsFut As Worksheet
    Dim wsSw As Worksheet

    Set wsFut = ThisWorkbook.Worksheets(SH_FUTURES)
    Set wsSw = ThisWorkbook.Worksheets(SH_SWAPS)

    mHedgeLookupFutLast = SheetClearLastRow( _
        wsFut, DATA_ROW, LastFutureDataRow(wsFut))

    mHedgeLookupSwLast = SheetClearLastRow( _
        wsSw, DATA_ROW, LastSwapDataRow(wsSw))

End Sub


' Number of PNL_Attribution rows implied by the bond count, floored at 0.
Private Function PnlRowCount(ByVal lastBondRow As Long) As Long

    If lastBondRow < BOND_DATA_ROW Then
        PnlRowCount = 0
    Else
        PnlRowCount = lastBondRow - BOND_DATA_ROW + 1
    End If

End Function


Private Sub WritePNLRow(ws As Worksheet, bondRow As Long, pnlRow As Long)


    Dim B As String: B = "'" & SH_BONDS & "'!"
    Dim c As String: c = "'" & SH_CONFIG & "'!"
    Dim fu As String: fu = "'" & SH_FUTURES & "'!"
    Dim sw As String: sw = "'" & SH_SWAPS & "'!"


    Dim p As String
    Dim b2 As String
    Dim futLast As String
    Dim swLast As String


    p = CStr(pnlRow)
    b2 = CStr(bondRow)


    ' Hedge lookup ranges for this row's SUMIFS/COUNTIFS.  Sized from the hedge
    ' sheets themselves - the number of futures and swaps is whatever Hedge Risco
    ' loaded today, and a fixed 220/200 either missed hedges or summed blanks.
    futLast = CStr(mHedgeLookupFutLast)
    swLast = CStr(mHedgeLookupSwLast)


    ' -------------------------------------------------------------------------
    ' A:J = Identity / period
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_ISIN & p).formula = "=UPPER(TRIM(" & B & BCOL_ISIN & b2 & "))"
    ws.Range(PCOL_NAME & p).formula = "=" & B & BCOL_NAME & b2
    ws.Range(PCOL_CCY & p).formula = "=" & B & BCOL_CCY & b2
    ws.Range(PCOL_PORTFOLIO & p).formula = "=" & B & BCOL_PORTFOLIO & b2
    ws.Range(PCOL_ACCTGCAT & p).formula = "=" & B & BCOL_ACCTGCAT & b2
    ws.Range(PCOL_NOTIONAL & p).formula = "=" & B & BCOL_NOTIONAL & b2
    ws.Range(PCOL_DAYS & p).formula = _
        "=INT(" & c & CFG_T1_DATE & ")-INT(" & c & CFG_T0_DATE & ")"
    ws.Range(PCOL_YEARFRAC & p).formula = "=IFERROR(" & PCOL_DAYS & p & "/365,"""")"


    ' -------------------------------------------------------------------------
    ' K:V = Delta-only block
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_DELTA_DIRTY_MV_EUR & p).formula = MarketValueChangeFml( _
        B & BCOL_DIRTYMV_T0_EUR & b2, B & BCOL_DIRTYMV_TM1_EUR & b2)


    ws.Range(PCOL_DELTA_Y_BP & p).formula = "=" & B & BCOL_DELTA_Y_BP & b2


    ws.Range(PCOL_DELTA_R_BP & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_OIS_T0 & b2 & "),ISNUMBER(" & B & BCOL_OIS_TM1 & b2 & "))," & _
        "(" & B & BCOL_OIS_T0 & b2 & "-" & B & BCOL_OIS_TM1 & b2 & ")*100,"""")"


    ws.Range(PCOL_DELTA_GOV_BP & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_GOV_T0 & b2 & "),ISNUMBER(" & B & BCOL_GOV_TM1 & b2 & "))," & _
        "(" & B & BCOL_GOV_T0 & b2 & "-" & B & BCOL_GOV_TM1 & b2 & ")*100,"""")"


    ws.Range(PCOL_DELTA_G_BP & p).formula = "=" & B & BCOL_DELTA_G_BP & b2


    ws.Range(PCOL_DELTA_SWAP_BP & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_SWAP_T0 & b2 & "),ISNUMBER(" & B & BCOL_SWAP_TM1 & b2 & "))," & _
        "(" & B & BCOL_SWAP_T0 & b2 & "-" & B & BCOL_SWAP_TM1 & b2 & ")*100,"""")"


    ws.Range(PCOL_DELTA_Q_BP & p).formula = "=" & B & BCOL_DELTA_Q_BP & b2
    ws.Range(PCOL_DELTA_I_BP & p).formula = "=" & B & BCOL_DELTA_I_BP & b2


    ws.Range(PCOL_DELTA_Z_BP & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_ZSPRD_T0 & b2 & "),ISNUMBER(" & B & BCOL_ZSPRD_TM1 & b2 & "))," & _
        B & BCOL_ZSPRD_T0 & b2 & "-" & B & BCOL_ZSPRD_TM1 & b2 & ","""")"


    ws.Range(PCOL_DELTA_G_BP_T & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_GSPREAD_T0 & b2 & "),ISNUMBER(" & B & BCOL_GSPREAD_TM1 & b2 & "))," & _
        B & BCOL_GSPREAD_T0 & b2 & "-" & B & BCOL_GSPREAD_TM1 & b2 & ","""")"


    ws.Range(PCOL_DELTA_ASW_BP & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_ASW_T0 & b2 & "),ISNUMBER(" & B & BCOL_ASW_TM1 & b2 & "))," & _
        B & BCOL_ASW_T0 & b2 & "-" & B & BCOL_ASW_TM1 & b2 & ","""")"


    ws.Range(PCOL_DELTA_OAS_BP & p).formula = "=" & B & BCOL_DELTAOAS & b2


    ' -------------------------------------------------------------------------
    ' W:Z = Risk fields
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_MODDUR & p).formula = "=" & B & BCOL_MODDUR_T0 & b2
    ws.Range(PCOL_CONVEXITY & p).formula = "=" & B & BCOL_CONVEXITY & b2
    ws.Range(PCOL_SPREADDURATION & p).formula = "=" & B & BCOL_SPREADDURATION_USED & b2


    ' -------------------------------------------------------------------------
    ' AA:AE = Main duration attribution
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PNL_OIS & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_R_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_R_BP & p & ","""")"


    ws.Range(PCOL_PNL_GOVBASIS & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_G_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_G_BP & p & ","""")"


    ws.Range(PCOL_PNL_SWAPGOVBASIS & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Q_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_Q_BP & p & ","""")"


    ws.Range(PCOL_PNL_CREDIT_ISPREAD & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_I_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_I_BP & p & ","""")"
        
    ' -------------------------------------------------------------------------
    ' AA = PnL_Duration_Total
    '
    ' The selected framework decides how the yield move is SPLIT; every chain
    ' sums back to the same -DV01 * Delta_y.  See the SPREAD FRAMEWORK LIBRARY
    ' block near the top of this module for the identities.
    '
    ' A SWITCH default of "" is mandatory here.  Without it a framework code
    ' with no arm (ASW and Z were both reachable from CF but unhandled)
    ' evaluated to #N/A, which blanked Total_Model_Explained and left the bond
    ' with no attribution at all.
    ' -------------------------------------------------------------------------
    Dim chainG As String
    Dim chainI As String


    chainG = ChainSumFml(p, Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_GSPREAD))
    chainI = ChainSumFml(p, Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_SWAPGOVBASIS, PCOL_PNL_CREDIT_ISPREAD))


    ws.Range(PCOL_PNL_DURATION_TOTAL & p).formula = _
        "=SWITCH(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & "," & _
        """G""," & chainG & "," & _
        """I""," & chainI & "," & _
        """ASW""," & ChainSumFml(p, Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_SWAPGOVBASIS, PCOL_PNL_ASW)) & "," & _
        """Z""," & ChainSumFml(p, Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_SWAPGOVBASIS, PCOL_PNL_ZSPREAD)) & "," & _
        """OAS""," & ChainSumFml(p, Array(PCOL_PNL_OIS, PCOL_PNL_GOVBASIS, PCOL_PNL_SWAPGOVBASIS, PCOL_PNL_OAS)) & "," & _
        """OIS""," & YieldOnlyPnLFml(p) & "," & _
        """SOFR""," & YieldOnlyPnLFml(p) & "," & _
        """MIXED""," & DV01BlendFml(p, chainG, chainI) & "," & _
        """REVIEW"",""""," & _
        """"")"


    ' -------------------------------------------------------------------------
    ' AM = PnL_Convexity
    '
    ' Second-order term of the same expansion the duration leg is the first
    ' order of, so it is anchored at the same point: DirtyMV at T-1, beside
    ' Bond_DV01_Opening.  The PLUS sign is only correct at that anchor.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PNL_CONVEXITY & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_DIRTYMV_TM1_EUR & b2 & ")," & _
        "ISNUMBER(" & PCOL_CONVEXITY & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Y_BP & p & "))," & _
        "0.5*" & B & BCOL_DIRTYMV_TM1_EUR & b2 & "*" & _
        PCOL_CONVEXITY & p & "*(" & PCOL_DELTA_Y_BP & p & "/10000)^2,"""")"




    ' -------------------------------------------------------------------------
    ' Roll-to-par uses prior CLEAN price Bonds!AB, not prior DIRTY price
    ' Bonds!AC. This prevents accrued interest from leaking into pull-to-par
    ' while coupon accrual is already being captured in Carry_Coupon.
    '
    ' Sign convention:
    '   PositionSide (H) is the single +/-1 sign driver and is ALWAYS paired
    '   with ABS(Notional), exactly as Bonds!AG DV01_EUR and Bonds!AH/AI
    '   DirtyMV do.  Multiplying PositionSide by a raw or already-signed
    '   quantity double-applies the sign and silently flips shorts.
    '
    ' FX convention:
    '   Carry accrues on the position as held at the START of the period, so
    '   both coupon and roll-to-par translate at the prior fix Bonds!FX_T-1.
    '   Using FX_T0 for one and FX_T-1 for the other mixes an FX effect into
    '   carry that PnL_FX (AP) is already accounting for separately.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_CARRY_COUPON & p).formula = _
        "=IF(AND(" & _
        "ISNUMBER(" & PCOL_NOTIONAL & p & ")," & _
        "ISNUMBER(" & B & BCOL_COUPON & b2 & ")," & _
        "ISNUMBER(" & PCOL_YEARFRAC & p & ")," & _
        "ISNUMBER(" & B & BCOL_FX_TM1 & b2 & "))," & _
        PCOL_NOTIONAL & p & "*(" & _
        B & BCOL_COUPON & b2 & "/100)*" & _
        PCOL_YEARFRAC & p & "*" & _
        B & BCOL_FX_TM1 & b2 & ","""")"


    ws.Range(PCOL_CARRY_ROLLTOPAR & p).formula = RollToParFml( _
        PCOL_NOTIONAL & p, _
        B & BCOL_FX_TM1 & b2, _
        PullToParPriceFml( _
            c & CFG_T0_DATE, _
            c & CFG_T1_DATE, _
            B & BCOL_MATURITY & b2, _
            B & BCOL_COUPON & b2, _
            B & BCOL_COUPONFREQ_NUM & b2, _
            B & BCOL_BONDDCC_CODE & b2, _
            PCOL_CCY & p, _
            PCOL_SPREAD_FRAMEWORK_AUTO & p, _
            PriorSpreadFml( _
                PCOL_SPREAD_FRAMEWORK_AUTO & p, _
                B & BCOL_GSPREAD_TM1 & b2, _
                B & BCOL_ISPREAD_TM1 & b2, _
                B & BCOL_ASW_TM1 & b2, _
                B & BCOL_ZSPRD_TM1 & b2, _
                B & BCOL_OAS_TM1 & b2)))


    ws.Range(PCOL_CARRY_FUNDING & p).formula = _
        "=IF(AND(ISNUMBER(" & B & BCOL_DIRTYMV_TM1_EUR & b2 & ")," & _
        "ISNUMBER(" & PCOL_YEARFRAC & p & "))," & _
        "LET(" & _
        "_r0," & B & BCOL_FUNDRATE_T0 & b2 & "," & _
        "_r1," & B & BCOL_FUNDRATE_TM1 & b2 & "," & _
        "_fr,IF(AND(ISNUMBER(_r0),ISNUMBER(_r1)),AVERAGE(_r0,_r1)," & _
        "IF(ISNUMBER(_r1),_r1,IF(ISNUMBER(_r0),_r0,"""")))," & _
        "IF(ISNUMBER(_fr),-" & B & BCOL_DIRTYMV_TM1_EUR & b2 & "*" & _
        "IF(ABS(_fr)>1,_fr/100,_fr)*" & PCOL_YEARFRAC & p & ","""")),"""")"


    ' AJ = Carry_Total = coupon + roll-to-par.
    '
    ' AI Funding_Carry_Memo is deliberately NOT included.  AV Official_Total_PnL
    ' is a mark-to-market proxy (change in dirty MV + actual hedge PnL) and
    ' contains no financing leg, so adding funding to the explained side would
    ' open a residual exactly equal to the funding cost.  AI is reported as an
    ' economic-carry memo instead and is shown outside the bridge subtotal on the
    ' Dashboard.  If the official PnL feed ever becomes funded rather than pure
    ' MTM, add AI here AND to AV in the same edit.
    ws.Range(PCOL_CARRY_TOTAL & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_CARRY_COUPON & p & ")," & _
        "ISNUMBER(" & PCOL_CARRY_ROLLTOPAR & p & "))," & _
        PCOL_CARRY_COUPON & p & "+" & PCOL_CARRY_ROLLTOPAR & p & ","""")"


    ' -------------------------------------------------------------------------
    ' AK:AO = Alternative spread PnLs
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PNL_ZSPREAD & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Z_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_Z_BP & p & ","""")"


    ws.Range(PCOL_PNL_GSPREAD & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_G_BP_T & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_G_BP_T & p & ","""")"


    ws.Range(PCOL_PNL_ASW & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_ASW_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_ASW_BP & p & ","""")"

    ws.Range(PCOL_PNL_OAS & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_OAS_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_OAS_BP & p & ","""")"
        
        
    ' -------------------------------------------------------------------------
    ' AO = SpreadPnL_Used
    '
    ' The credit/spread SUB-COMPONENT of the selected framework's chain.  This is
    ' a REPORTING column: it is already inside AA PnL_Duration_Total and must NOT
    ' be added to Total_Model_Explained, or the spread leg is counted twice.
    '
    ' For OIS/SOFR there is no explicit spread leg, so the spread is imputed as
    ' the part of the yield move not explained by the OIS curve.
    '
    ' As with AA, the SWITCH carries a "" default so an unhandled framework code
    ' can never produce #N/A.
    ' -------------------------------------------------------------------------
    Dim spreadOverOis As String


    spreadOverOis = _
        "IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Y_BP & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_R_BP & p & "))," & _
        "-" & PCOL_BOND_DV01_OPENING & p & "*(" & _
        PCOL_DELTA_Y_BP & p & "-" & _
        PCOL_DELTA_R_BP & p & "),"""")"


    ws.Range(PCOL_SPREADPNL_USED & p).formula = _
        "=SWITCH(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & "," & _
        """G""," & PCOL_PNL_GSPREAD & p & "," & _
        """I""," & PCOL_PNL_CREDIT_ISPREAD & p & "," & _
        """ASW""," & PCOL_PNL_ASW & p & "," & _
        """Z""," & PCOL_PNL_ZSPREAD & p & "," & _
        """OAS""," & PCOL_PNL_OAS & p & "," & _
        """OIS""," & spreadOverOis & "," & _
        """SOFR""," & spreadOverOis & "," & _
        """MIXED""," & DV01BlendFml(p, PCOL_PNL_GSPREAD & p, PCOL_PNL_CREDIT_ISPREAD & p) & "," & _
        """REVIEW"",""""," & _
        """"")"


    ' -------------------------------------------------------------------------
    ' AP = FX
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PNL_FX & p).formula = _
        "=IFERROR(IF(" & PCOL_CCY & p & "=" & Chr(34) & "EUR" & Chr(34) & ",0,(" & B & BCOL_DIRTYMV_TM1_EUR & b2 & "/" & _
        B & BCOL_FX_TM1 & b2 & ")*(" & B & BCOL_FX_T0 & b2 & "-" & B & BCOL_FX_TM1 & b2 & ")),0)"


    ' -------------------------------------------------------------------------
    ' CF = FX_Exposure_EUR
    '
    ' The opening EUR value of the position that is exposed to a currency move -
    ' i.e. the amount the EUR/USD hedge is there to neutralise.  Zero for EUR
    ' bonds, which have nothing to translate.
    '
    ' PnL_FX above is this number times the proportional FX move, so the two
    ' together let the Dashboard state the FX hedge in the only terms that mean
    ' anything: how much exposure there was, how much of it the EUR/USD
    ' contracts actually covered, and what the mismatch cost.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_FX_EXPOSURE_EUR & p).formula = _
        "=IF(" & PCOL_ISIN & p & "="""",""""," & _
        "IF(" & PCOL_CCY & p & "=""EUR"",0," & _
        "IFERROR(" & B & BCOL_DIRTYMV_TM1_EUR & b2 & ",0)))"


    ' -------------------------------------------------------------------------
    ' BF = Futures_DV01
    ' BD = PlainSwap_DV01
    ' BE = SyntheticSwap_DV01
    '
    ' These are written before AQ/AR/AT because AQ/AR now use the hedge DV01s.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PLAINSWAP_DV01 & p).formula = _
        "=IFERROR(SUMIFS(" & _
            sw & "$" & WCOL_DV01_BBG & "$" & DATA_ROW & ":$" & WCOL_DV01_BBG & "$" & swLast & "," & _
            sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
            sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""PLAIN""),0)"




    ws.Range(PCOL_SYNTHETICSWAP_DV01 & p).formula = _
        "=IFERROR(SUMIFS(" & _
            sw & "$" & WCOL_DV01_BBG & "$" & DATA_ROW & ":$" & WCOL_DV01_BBG & "$" & swLast & "," & _
            sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
            sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""SYNTHETIC""),0)"


    
    ' Futures risk, split by the coverage file it came from.  Both are SUMIFS on
    ' the same Futures! column with an extra criterion on Hedge_Source, so they
    ' partition the sheet exactly and their sum is the whole futures hedge.
    ws.Range(PCOL_FUTURES_RTJ_DV01 & p).formula = _
        "=IFERROR(SUMIFS(" & fu & "$" & FCOL_FUTURES_DV01_EUR & "$" & DATA_ROW & ":$" & FCOL_FUTURES_DV01_EUR & "$" & futLast & "," & _
        fu & "$" & FCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & FCOL_LINKEDISIN & "$" & futLast & "," & PCOL_ISIN & p & "," & _
        fu & "$" & FCOL_HEDGE_SOURCE & "$" & DATA_ROW & ":$" & FCOL_HEDGE_SOURCE & "$" & futLast & ",""" & HEDGE_SOURCE_RTJ & """" & _
        FutClassCrit(fu, futLast, HEDGE_CLASS_RATES) & "),0)"


    ws.Range(PCOL_FUTURES_RT_DV01 & p).formula = _
        "=IFERROR(SUMIFS(" & fu & "$" & FCOL_FUTURES_DV01_EUR & "$" & DATA_ROW & ":$" & FCOL_FUTURES_DV01_EUR & "$" & futLast & "," & _
        fu & "$" & FCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & FCOL_LINKEDISIN & "$" & futLast & "," & PCOL_ISIN & p & "," & _
        fu & "$" & FCOL_HEDGE_SOURCE & "$" & DATA_ROW & ":$" & FCOL_HEDGE_SOURCE & "$" & futLast & ",""" & HEDGE_SOURCE_RT & """" & _
        FutClassCrit(fu, futLast, HEDGE_CLASS_RATES) & "),0)"
        
    ' -------------------------------------------------------------------------
    ' CA = Spread_Framework_Auto      CB = Spread_Framework_Reason
    '
    ' Which spread the bond's credit leg is measured against.  Precedence:
    '
    '   1. SpreadOverride sheet   per-bond, keyed on ISIN (manual, persistent)
    '   2. automatic              from the hedge DV01 mix actually in the book
    '   3. Config!B18             LAST RESORT, only when the automatic rule
    '                             cannot resolve (the bond has no spread leg at
    '                             all, so there is nothing to pick between)
    '
    ' B18 used to sit at step 2, ahead of the automatic rule.  That made it a
    ' blanket override: setting it to (say) "I" for one problem bond silently
    ' re-based every OTHER bond in the book onto the swap curve, including the
    ' futures-hedged ones whose residual risk is against governments, and it
    ' switched off the automatic MIXED classification entirely.  The whole point
    ' of the automatic rule is that the framework follows the hedge; a global
    ' cell cannot know a per-bond fact.  Per-bond corrections belong on the
    ' SpreadOverride sheet, which is what step 1 is for.
    '
    ' The automatic rule follows what the hedge actually neutralises:
    '   - Futures-dominated hedge -> G-spread.  Bond futures hedge the deliverable
    '     government curve via the CTD, so what is left unhedged is the bond's
    '     spread OVER GOVERNMENTS.
    '   - Swap-dominated hedge -> I-spread.  A swap hedges the swap curve, so what
    '     is left unhedged is the bond's spread OVER SWAPS.
    '   - BOTH materially present -> MIXED.  When neither leg carries less than
    '     MIXED_MIN_SHARE of the combined hedge BPV the residual risk is genuinely
    '     against both curves, and forcing it onto one of them books the other
    '     curve's move as credit.  MIXED runs both chains and weights them by the
    '     same DV01 split (DV01BlendFml), so the blend is the hedge's own mix
    '     rather than a judgement call.  It needs both the G and the I leg to be
    '     numeric - with only one of them the blend degenerates to that leg
    '     anyway, so the rule falls through to the dominance test instead.
    '   - Unhedged -> prefer I, then G, then ASW, then Z, then OAS by availability.
    '   - No spread leg numeric at all -> "" , i.e. the automatic rule abstains
    '     and B18 gets its turn.
    '
    ' Dominance is measured on ABS DV01, so it is unaffected by the sign
    ' convention of either hedge leg.  SyntheticSwap_DV01 deliberately does NOT
    ' vote: a synthetic is the hedge the desk COULD have put on, not the one
    ' whose risk is actually in the book.
    '
    ' An unrecognised override code resolves to REVIEW rather than being silently
    ' ignored, so a typo suppresses attribution loudly instead of quietly
    ' changing the answer.  So does an abstaining automatic rule with B18 blank:
    ' a bond with no usable spread leg is a data problem, not a framework choice.
    ' -------------------------------------------------------------------------


    ws.Range(PCOL_SPREAD_FRAMEWORK_AUTO & p).formula = SpreadFrameworkAutoFml(p, c)


    ws.Range(PCOL_SPREAD_FRAMEWORK_REASON & p).formula = SpreadFrameworkReasonFml(p, c)




    ' -------------------------------------------------------------------------
    ' AQ:AT = Expected hedge PnL and explicit hedge basis
    '
    ' AQ = expected futures hedge PnL from government curve move
    '      Futures are government bond futures / CTD-basket hedges.
    '
    ' AR = expected plain swap hedge PnL from selected swap/OIS curve.
    '      ESTR swaps use Delta_r_bp.
    '      EURIBOR swaps use Delta_Swap_bp.
    '
    ' AT = basis / hedge mismatch:
    '      actual futures + actual plain swap PnL
    '      minus expected futures + expected swap model PnL.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_PNL_FUTURES & p).formula = _
        "=IF(" & PCOL_FUT_MATCH_COUNT & p & "=0,0," & _
        "IF(AND(ISNUMBER" & FutDv01(p) & "," & _
        "ISNUMBER(" & PCOL_DELTA_GOV_BP & p & "))," & _
        "-" & FutDv01(p) & "*" & _
        PCOL_DELTA_GOV_BP & p & ",""""))"


    ws.Range(PCOL_PNL_SWAP & p).formula = _
        "=IF(" & PCOL_SWAP_PLAIN_MATCH_COUNT & p & "=0,0," & _
        "IF(COUNTIFS(" & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & _
        ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""PLAIN""," & _
        sw & "$" & WCOL_FLOATCURVE_TYPE & "$" & DATA_ROW & _
        ":$" & WCOL_FLOATCURVE_TYPE & "$" & swLast & ",""UNKNOWN"")>0,""""," & _
        "-SUMPRODUCT((" & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & WCOL_LINKEDISIN & "$" & swLast & "=" & PCOL_ISIN & p & ")*(" & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & _
        ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & "=""PLAIN"")*(" & _
        sw & "$" & WCOL_DV01_BBG & "$" & DATA_ROW & _
        ":$" & WCOL_DV01_BBG & "$" & swLast & ")*IF((" & _
        sw & "$" & WCOL_FLOATCURVE_TYPE & "$" & DATA_ROW & _
        ":$" & WCOL_FLOATCURVE_TYPE & "$" & swLast & "=""ESTR"")+(" & _
        sw & "$" & WCOL_FLOATCURVE_TYPE & "$" & DATA_ROW & _
        ":$" & WCOL_FLOATCURVE_TYPE & "$" & swLast & "=""SOFR"")," & _
        PCOL_DELTA_R_BP & p & ",IF(" & _
        sw & "$" & WCOL_FLOATCURVE_TYPE & "$" & DATA_ROW & _
        ":$" & WCOL_FLOATCURVE_TYPE & "$" & swLast & "=""EURIBOR""," & _
        PCOL_DELTA_SWAP_BP & p & ",0)))))"




ws.Range(PCOL_TOTAL_HEDGE & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_PNL_FUTURES & p & ")," & _
        "ISNUMBER(" & PCOL_PNL_SWAP & p & "))," & _
        PCOL_PNL_FUTURES & p & "+" & PCOL_PNL_SWAP & p & ","""")"
    
    ws.Range(PCOL_FUT_PNL_RAW & p).formula = _
        "=IF(" & PCOL_FUT_MATCH_COUNT & p & "=0,0," & _
        "IF(COUNTIFS(" & _
        fu & "$" & FCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & FCOL_LINKEDISIN & "$" & futLast & "," & PCOL_ISIN & p & "," & _
        fu & "$" & FCOL_FUTURESPNL_EUR & "$" & DATA_ROW & _
        ":$" & FCOL_FUTURESPNL_EUR & "$" & futLast & ","">=-1E+307""" & _
        FutClassCrit(fu, futLast, HEDGE_CLASS_RATES) & ")=0,""""," & _
        "SUMIFS(" & _
        fu & "$" & FCOL_FUTURESPNL_EUR & "$" & DATA_ROW & _
        ":$" & FCOL_FUTURESPNL_EUR & "$" & futLast & "," & _
        fu & "$" & FCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & FCOL_LINKEDISIN & "$" & futLast & "," & PCOL_ISIN & p & _
        FutClassCrit(fu, futLast, HEDGE_CLASS_RATES) & ")))"


    ws.Range(PCOL_SWAP_PLAIN_PNL_RAW & p).formula = _
        "=IF(" & PCOL_SWAP_PLAIN_MATCH_COUNT & p & "=0,0," & _
        "IF(COUNTIFS(" & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & _
        ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""PLAIN""," & _
        sw & "$" & WCOL_PNL & "$" & DATA_ROW & _
        ":$" & WCOL_PNL & "$" & swLast & ","">=-1E+307"")=0,""""," & _
        "SUMIFS(" & _
        sw & "$" & WCOL_PNL & "$" & DATA_ROW & _
        ":$" & WCOL_PNL & "$" & swLast & "," & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & _
        ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""PLAIN"")))"


    ws.Range(PCOL_BASIS_PNL & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_FUTURES_BASIS_PNL & p & ")," & _
        "ISNUMBER(" & PCOL_SWAP_BASIS_PNL & p & "))," & _
        PCOL_FUTURES_BASIS_PNL & p & "+" & _
        PCOL_SWAP_BASIS_PNL & p & ","""")"
    ' -------------------------------------------------------------------------
    ' AU:AX = explained / official / residual
    '
    ' AU = Total_Model_Explained
    '        = AA Duration
    '        + AF Convexity
    '        + AJ Carry            (coupon + roll-to-par; funding is a memo, see AI)
    '        + AP FX
    '        + AS Hedge_Curve_Model_PnL      (MODEL hedge PnL)
    '        + AT Hedge_Model_Residual_PnL   (hedge basis = ACTUAL - MODEL hedge)
    '
    ' AT belongs here.  AV Official_Total_PnL is built from the ACTUAL hedge PnL
    ' (CD), so leaving AT out guaranteed that the futures/CTD basis and the swap
    ' spread mismatch landed in AW Unexplained_Residual_PnL every single day,
    ' regardless of model quality.  With AT included, AS + AT telescopes to the
    ' actual hedge PnL and AW measures only bond-leg model error - which is what
    ' "unexplained" is supposed to mean.
    '
    ' AT is still reported separately (and charted on the Dashboard) so the desk
    ' can watch the basis in its own right.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_TOTAL_EXPLAINED & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_PNL_DURATION_TOTAL & p & ")," & _
        "ISNUMBER(" & PCOL_PNL_CONVEXITY & p & ")," & _
        "ISNUMBER(" & PCOL_CARRY_TOTAL & p & ")," & _
        "ISNUMBER(" & PCOL_PNL_FX & p & ")," & _
        "ISNUMBER(" & PCOL_TOTAL_HEDGE & p & ")," & _
        "ISNUMBER(" & PCOL_BASIS_PNL & p & "))," & _
        PCOL_PNL_DURATION_TOTAL & p & "+" & _
        PCOL_PNL_CONVEXITY & p & "+" & _
        PCOL_CARRY_TOTAL & p & "+" & _
        PCOL_PNL_FX & p & "+" & _
        PCOL_TOTAL_HEDGE & p & "+" & _
        PCOL_BASIS_PNL & p & ","""")"
    ' -------------------------------------------------------------------------
    ' CI = Coupon_Paid_EUR
    '
    ' Coupon CASH actually received between T-1 and T0, as opposed to the
    ' smooth accrual in Carry_Coupon.
    '
    ' Official_Total_PnL is built from the change in DIRTY market value, and
    ' dirty value carries accrued interest.  On a coupon date accrued resets to
    ' zero, so the dirty value drops by the whole coupon while the desk's cash
    ' rises by the same amount.  Without this term the drop was booked as a loss
    ' and the cash was booked nowhere: the bond's residual opened by exactly one
    ' coupon on every ex-date, and only on ex-dates, which reads like a model
    ' failure and is really a missing cash flow.
    '
    ' Carry_Coupon is NOT changed and NOT double counted: it stays the accrual
    ' over the period, which is the theoretical carry.  This is the practical
    ' side of the same money, and it belongs with the other practical numbers.
    ' It translates at the prior fix, like the rest of the opening position.
    '
    ' The "+1" on the end date is not a fudge.  CouponDatesBetween counts
    ' coupons in (start, end) with the end EXCLUSIVE, which is right for the
    ' futures carry it was written for.  Here the window has to be (T-1, T0]:
    ' if T0 is the ex-date, the T0 dirty price has already reset accrued to
    ' zero, so that coupon is exactly the one this term exists to put back.
    ' Widening the call is safer than changing the shared helper, which the
    ' futures implied-repo chain also depends on.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_COUPON_PAID_EUR & p).formula = _
        "=IF(" & PCOL_ISIN & p & "="""",""""," & _
        "IFERROR(" & PCOL_NOTIONAL & p & "*SumCouponsBetween(" & _
        c & CFG_T0_DATE & "," & c & CFG_T1_DATE & "+1," & _
        B & BCOL_MATURITY & b2 & "," & B & BCOL_COUPON & b2 & "/100," & _
        B & BCOL_COUPONFREQ_NUM & b2 & ")/100*" & B & BCOL_FX_TM1 & b2 & ",0))"

    ' -------------------------------------------------------------------------
    ' BC = Official_Total_PnL   (the PRACTICAL side of the bridge)
    '
    '     change in dirty market value
    '   + coupon cash received
    '   + actual hedge PnL
    '
    ' Config!B34 no longer gates this.  It used to wrap the whole formula in
    ' IF(B34="PROXY", ..., ""), so any other value - a typo, a lower-case
    ' "proxy", a trailing space, or someone naming the source they intended to
    ' wire up later - blanked the official PnL on EVERY row of the book.  With
    ' no official total there is no residual, no attribution status and no
    ' bridge: one config cell silently emptied the entire sheet.  B34 stays as
    ' the LABEL for where this number came from, which is all it ever was -
    ' there is only one construction implemented, and it is this one.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_OFFICIAL_PNL & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_DELTA_DIRTY_MV_EUR & p & ")," & _
        "ISNUMBER(" & PCOL_ACTUAL_HEDGE_PNL_RAW & p & "))," & _
        PCOL_DELTA_DIRTY_MV_EUR & p & "+" & _
        "N(" & PCOL_COUPON_PAID_EUR & p & ")+" & _
        PCOL_ACTUAL_HEDGE_PNL_RAW & p & ","""")"
    ws.Range(PCOL_RESIDUAL & p).formula = ResidualPnLFml( _
        PCOL_OFFICIAL_PNL & p, PCOL_TOTAL_EXPLAINED & p)

    ws.Range(PCOL_RESIDUAL_AX & p).formula = ResidualPercentFml( _
        PCOL_RESIDUAL & p, PCOL_OFFICIAL_PNL & p)
    ' -------------------------------------------------------------------------
    ' AY:BC = hedge efficiency / attribution status
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_HEDGE_DV01 & p).formula = _
        "=IF(AND(ISNUMBER" & FutDv01(p) & "," & _
        "ISNUMBER(" & PCOL_PLAINSWAP_DV01 & p & "))," & _
        FutDv01(p) & "+" & PCOL_PLAINSWAP_DV01 & p & ","""")"

    ws.Range(PCOL_RESIDUAL_DV01 & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_BOND_DV01_CURRENT & p & ")," & _
        "ISNUMBER(" & PCOL_HEDGE_DV01 & p & "))," & _
        PCOL_BOND_DV01_CURRENT & p & "+" & _
        PCOL_HEDGE_DV01 & p & ","""")"

    ws.Range(PCOL_HEDGE_RATIO & p).formula = _
        "=IFERROR(-" & PCOL_HEDGE_DV01 & p & "/" & _
        PCOL_BOND_DV01_CURRENT & p & ","""")"
    ' -------------------------------------------------------------------------
    ' BB = Hedge_Efficiency = 1 - |Actual_Hedge_DV01 - Target_Hedge_DV01|
    '                             / |Target_Hedge_DV01|
    '
    ' Deliberately NOT clamped to [0,1].  The old MAX(0,MIN(1,...)) reported a
    ' wrong-way hedge and a 3x oversized hedge both as 0%, which hid exactly the
    ' positions worth looking at.  A negative number here is meaningful: -1 means
    ' the hedge is as wrong as it could be at that size.
    '
    ' 1 = the hedge matches its target.  The target comes from BN, so for a bond
    ' with a synthetic this measures replication, and without one it measures
    ' flatness.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_HEDGE_EFFICIENCY & p).formula = HedgeEfficiencyFml( _
        PCOL_HEDGE_DV01 & p, PCOL_TARGET_HEDGE_DV01 & p)


    ' -------------------------------------------------------------------------
    ' BC = Attribution_Status
    '
    ' First failing check wins, ordered so that a cause is reported ahead of its
    ' symptom: bad market data before a missing framework total, a missing
    ' framework total before a high residual.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_ATTRIBUTION_STATUS & p).formula = _
        "=IF(" & B & BCOL_BOND_STATUS & b2 & "<>""OK""," & _
        B & BCOL_BOND_STATUS & b2 & "," & _
        "IF(" & PCOL_SPREAD_FRAMEWORK_AUTO & p & "=""REVIEW"",""Framework review required""," & _
        "IF(AND(" & PCOL_SWAP_PLAIN_MATCH_COUNT & p & ">0," & _
        "NOT(ISNUMBER(" & PCOL_SWAP_PLAIN_PNL_RAW & p & "))),""Missing actual swap PnL""," & _
        "IF(AND(" & PCOL_FUT_MATCH_COUNT & p & ">0," & _
        "NOT(ISNUMBER(" & PCOL_FUT_PNL_RAW & p & "))),""Missing actual futures PnL""," & _
        "IF(AND(ISNUMBER(" & PCOL_HEDGE_RATIO & p & ")," & _
        PCOL_HEDGE_RATIO & p & "<0),""Wrong-direction hedge""," & _
        "IF(AND(ISNUMBER(" & PCOL_HEDGE_RATIO & p & ")," & _
        PCOL_HEDGE_RATIO & p & ">1+" & HEDGE_RATIO_TOL & "),""Over-hedged""," & _
        "IF(NOT(ISNUMBER(" & PCOL_PNL_DURATION_TOTAL & p & ")),""Missing selected framework data""," & _
        "IF(AND(ISNUMBER(" & PCOL_DURATION_IDENTITY_CHECK & p & ")," & _
        "ABS(" & PCOL_DURATION_IDENTITY_CHECK & p & ")>MAX(" & IDENTITY_TOL_EUR & "," & _
        IDENTITY_TOL_PCT & "*ABS(" & PCOL_PNL_DURATION_TOTAL & p & "))),""Duration chain does not tie""," & _
        "IF(NOT(ISNUMBER(" & PCOL_OFFICIAL_PNL & p & ")),""Missing official PnL""," & _
        "IF(AND(ISNUMBER(" & PCOL_RESIDUAL_AX & p & ")," & _
        "ABS(" & PCOL_RESIDUAL_AX & p & ")>" & RESIDUAL_PCT_TOL & "),""High residual"",""OK""))))))))))"


    ' -------------------------------------------------------------------------
    ' BG:BJ = DV01 diagnostics
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_HEDGE_DV01 & p).formula = _
        "=IF(AND(ISNUMBER" & FutDv01(p) & "," & _
        "ISNUMBER(" & PCOL_PLAINSWAP_DV01 & p & "))," & _
        FutDv01(p) & "+" & PCOL_PLAINSWAP_DV01 & p & ","""")"
    ws.Range(PCOL_THEORETICALHEDGE_DV01 & p).formula = "=" & PCOL_SYNTHETICSWAP_DV01 & p
    ws.Range(PCOL_BOND_DV01_CURRENT & p).formula = "=" & B & BCOL_DV01_EUR & b2

    ' The opening-risk twin.  Bond_DV01_Current answers 'how much risk is in
    ' the book right now' and drives the hedge ratio, the hedge gap and the
    ' Dashboard's BPV weights.  Bond_DV01_Opening answers 'how much risk was in
    ' the book when the move started' and drives every attribution leg.  They
    ' are the same formula one day apart; see Bonds!DV01_Opening_EUR.
    ws.Range(PCOL_BOND_DV01_OPENING & p).formula = "=" & B & BCOL_DV01_OPENING_EUR & b2

    ' What the choice between the two is worth, in EUR, on this row.
    '
    ' It is the whole of the difference between attributing the day on opening
    ' risk and attributing it on closing risk, so it bounds the argument rather
    ' than leaving it to be argued.  It is normally small - a day of price and
    ' FX drift on the same duration - and if it is not, the position changed
    ' size during the day and no single-point attribution is honest about it.
    ws.Range(PCOL_RISK_TIMING_BIAS & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_BOND_DV01_CURRENT & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Y_BP & p & "))," & _
        "-(" & PCOL_BOND_DV01_OPENING & p & "-" & PCOL_BOND_DV01_CURRENT & p & ")*" & _
        PCOL_DELTA_Y_BP & p & ","""")"

    ' Bond_DV01_Initial_Approx, Bond_DV01_Change_Approx and DV01_Comparison_Ratio
    ' used to sit here.  All three were approximations of the OPENING bond risk
    ' and of the drift from it, rebuilt on this sheet from Bonds!DirtyMV_T-1 and a
    ' duration - a second, weaker copy of a quantity Bonds already carries.  They
    ' are gone: Hedge_DV01_Gap and Hedge_Efficiency below answer the same
    ' question ("is the hedge the size it should be") from the actual and target
    ' risk, without re-deriving anything.
    ' -------------------------------------------------------------------------
    ' BN = Target_Hedge_DV01     the DV01 the hedge is SUPPOSED to have
    '
    '   1. SyntheticSwap_DV01 when it is populated and non-zero.  A synthetic is
    '      the hedge the coverage relationship says should be on, so the right
    '      question is "did we replicate it".
    '   2. otherwise -Bond_DV01_Current, i.e. "are we flat".
    '
    ' This used to be unconditionally -Bond_DV01_Current here while the Dashboard
    ' applied the synthetic-first rule, so the same bond scored differently
    ' depending on which sheet you read.  The rule now lives here only; the
    ' Dashboard aggregates BB and must not recompute it.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_TARGET_HEDGE_DV01 & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_SYNTHETICSWAP_DV01 & p & ")," & _
        PCOL_SYNTHETICSWAP_DV01 & p & "<>0)," & _
        PCOL_SYNTHETICSWAP_DV01 & p & "," & _
        "IF(AND(ISNUMBER(" & PCOL_BOND_DV01_CURRENT & p & ")," & _
        PCOL_BOND_DV01_CURRENT & p & "<>0)," & _
        "-" & PCOL_BOND_DV01_CURRENT & p & ",""""))"
    ws.Range(PCOL_HEDGE_DV01_GAP & p).formula = HedgeDv01GapFml( _
        PCOL_HEDGE_DV01 & p, PCOL_TARGET_HEDGE_DV01 & p)


    ' -------------------------------------------------------------------------
    ' BP:BW = futures/swap raw match diagnostics
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_FUT_MATCH_COUNT & p).formula = _
        "=COUNTIFS(" & fu & "$" & FCOL_LINKEDISIN & "$" & DATA_ROW & _
        ":$" & FCOL_LINKEDISIN & "$" & futLast & "," & PCOL_ISIN & p & _
        FutClassCrit(fu, futLast, HEDGE_CLASS_RATES) & ")"


    ' BQ already written above as actual futures raw PnL.


    ws.Range(PCOL_SWAP_ALL_MATCH_COUNT & p).formula = "=COUNTIF(" & sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & ")"


    ws.Range(PCOL_SWAP_PLAIN_MATCH_COUNT & p).formula = _
        "=COUNTIFS(" & sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""PLAIN"")"


    ' BT already written above as actual plain swap raw PnL.


    ws.Range(PCOL_SWAP_SYNTHETIC_MATCH_COUNT & p).formula = _
        "=COUNTIFS(" & sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""SYNTHETIC"")"


    ws.Range(PCOL_SWAP_SYNTHETIC_PNL_RAW & p).formula = _
        "=SUMIFS(" & sw & "$" & WCOL_PNL & "$" & DATA_ROW & ":$" & WCOL_PNL & "$" & swLast & "," & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & "," & _
        sw & "$" & WCOL_SWAP_ID_SOURCE & "$" & DATA_ROW & ":$" & WCOL_SWAP_ID_SOURCE & "$" & swLast & ",""SYNTHETIC"")"


    ws.Range(PCOL_SWAP_ALL_PNL_RAW & p).formula = _
        "=SUMIFS(" & sw & "$" & WCOL_PNL & "$" & DATA_ROW & ":$" & WCOL_PNL & "$" & swLast & "," & _
        sw & "$" & WCOL_LINKEDISIN & "$" & DATA_ROW & ":$" & WCOL_LINKEDISIN & "$" & swLast & "," & PCOL_ISIN & p & ")"


    ' -------------------------------------------------------------------------
    ' BX:CE = explicit hedge-basis bridge diagnostics
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_ACTUAL_FUTURES_PNL_RAW & p).formula = "=" & PCOL_FUT_PNL_RAW & p
    ws.Range(PCOL_FUTURES_GOV_MODEL_PNL & p).formula = "=" & PCOL_PNL_FUTURES & p
    ws.Range(PCOL_FUTURES_BASIS_PNL & p).formula = _
        "=IF(" & PCOL_FUT_MATCH_COUNT & p & "=0,0," & _
        "IF(AND(ISNUMBER(" & PCOL_FUT_PNL_RAW & p & ")," & _
        "ISNUMBER(" & PCOL_PNL_FUTURES & p & "))," & _
        PCOL_FUT_PNL_RAW & p & "-" & PCOL_PNL_FUTURES & p & ",""""))"


    ws.Range(PCOL_ACTUAL_SWAP_PNL_RAW & p).formula = "=" & PCOL_SWAP_PLAIN_PNL_RAW & p
    ws.Range(PCOL_SWAP_CURVE_MODEL_PNL & p).formula = "=" & PCOL_PNL_SWAP & p
    ws.Range(PCOL_SWAP_BASIS_PNL & p).formula = _
        "=IF(" & PCOL_SWAP_PLAIN_MATCH_COUNT & p & "=0,0," & _
        "IF(AND(ISNUMBER(" & PCOL_SWAP_PLAIN_PNL_RAW & p & ")," & _
        "ISNUMBER(" & PCOL_PNL_SWAP & p & "))," & _
        PCOL_SWAP_PLAIN_PNL_RAW & p & "-" & PCOL_PNL_SWAP & p & ",""""))"


    ws.Range(PCOL_ACTUAL_HEDGE_PNL_RAW & p).formula = _
        "=IF(AND(ISNUMBER(" & PCOL_FUT_PNL_RAW & p & ")," & _
        "ISNUMBER(" & PCOL_SWAP_PLAIN_PNL_RAW & p & "))," & _
        PCOL_FUT_PNL_RAW & p & "+" & _
        PCOL_SWAP_PLAIN_PNL_RAW & p & ","""")"


    ws.Range(PCOL_HEDGE_BASIS_PNL & p).formula = "=" & PCOL_BASIS_PNL & p


    ' -------------------------------------------------------------------------
    ' CH = Duration_Identity_Check
    '
    ' AA is built by SUMMING the selected framework's component columns, but
    ' every chain is an identity that must reproduce -DV01 * Delta_y.  This
    ' column is that difference:
    '
    '     Duration_Identity_Check = PnL_Duration_Total - ( -DV01_EUR * Delta_y )
    '
    ' Read it as a data-quality signal, not a model signal:
    '   ~0        the chain ties; the split is trustworthy
    '   non-zero  a curve leg is stale, a spread is quoted on a convention that
    '             does not reconcile to YTM, or T0/T-1 inputs come from
    '             different snapshots
    '
    ' G and I tie exactly.  ASW, Z and OAS are quoted on their own conventions
    ' rather than derived from YTM, so a small standing difference there is
    ' expected; OIS and SOFR are -DV01 * Delta_y by construction and tie to zero.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_DURATION_IDENTITY_CHECK & p).formula = _
        "=IF(AND(ISNUMBER(" & _
        PCOL_PNL_DURATION_TOTAL & p & ")," & _
        "ISNUMBER(" & PCOL_BOND_DV01_OPENING & p & ")," & _
        "ISNUMBER(" & PCOL_DELTA_Y_BP & p & "))," & _
        PCOL_PNL_DURATION_TOTAL & p & "-(-" & _
        PCOL_BOND_DV01_OPENING & p & "*" & _
        PCOL_DELTA_Y_BP & p & "),"""")"


    ' -------------------------------------------------------------------------
    ' CD:CE = Row_Valid / Row_Exclusion_Reason  (the quarantine)
    '
    ' A bond whose attribution does not compute cleanly must not sit inside the
    ' book totals.  Before this existed, one bond with a missing T-1 price
    ' contributed a blank to some legs and a number to others, so the bridge
    ' lines on the Dashboard were each summed over a DIFFERENT set of bonds and
    ' could not tie - which reads as a broken model when it is one broken row.
    ' Row_Valid draws the line once, and every total on the Dashboard is taken
    ' over Row_Valid = 1, so the excluded row cannot reach any other bond's
    ' number.  Row_Exclusion_Reason says why, and the Dashboard prints it.
    '
    ' What counts as broken is the row being UNCOMPUTABLE or SELF-INCONSISTENT:
    ' missing inputs, an unresolved framework, a matched hedge whose actual PnL
    ' never arrived, a duration chain that does not tie.  What does NOT count is
    ' the book being badly hedged - Over-hedged, Wrong-direction hedge and High
    ' residual are real, fully computed states of the book, and hiding them
    ' would remove exactly the PnL the reader opened the Dashboard to find.
    ' Those stay in the totals and stay flagged in Attribution_Status.
    ' -------------------------------------------------------------------------
    ws.Range(PCOL_ROW_EXCLUSION_REASON & p).formula = RowExclusionReasonFml(p, B, b2)

    ws.Range(PCOL_ROW_VALID & p).formula = _
        "=IF(" & PCOL_ISIN & p & "="""",""""," & _
        "IF(" & PCOL_ROW_EXCLUSION_REASON & p & "="""",1,0))"


End Sub
Private Function BDHLastPointExprR1C1( _
    ByVal secExpr As String, _
    ByVal fieldExpr As String, _
    ByVal startDateExpr As String, _
    ByVal endDateExpr As String) As String


    ' -------------------------------------------------------------------------
    ' Safer BDH single-point extractor.
    '
    ' OLD VERSION USED:
    '   "Days=W","Fill=P","UseDPDF=N","Dts=H","Cols=N"
    '
    ' That is dangerous because UseDPDF / Dts / Cols can be interpreted by
    ' Bloomberg as invalid override field IDs depending on the Excel API version.
    '
    ' NEW VERSION:
    '   Keeps only safer BDH options:
    '       "Days=W"
    '       "Fill=P"
    '       "Sort=A"
    '
    ' Then it handles either:
    '   - normal BDH 2-column output: Date | Value
    '   - compressed/single-column output, if Bloomberg returns that
    '
    ' It returns the last available value in the BDH result.
    ' -------------------------------------------------------------------------


    BDHLastPointExprR1C1 = _
        "IFERROR(" & _
            "LET(" & _
                "_bdh,BDH(" & secExpr & "," & fieldExpr & "," & _
                    startDateExpr & "," & endDateExpr & "," & _
                    """Days=W"",""Fill=P"",""Sort=A"")," & _
                "_r,ROWS(_bdh)," & _
                "_c,COLUMNS(_bdh)," & _
                "IF(_c>=2,INDEX(_bdh,_r,2),INDEX(_bdh,_r,1))" & _
            ")," & _
        """"")"


End Function




' =============================================================================
' BQL T0 SINGLE-POINT HELPERS
'
' Replaces BDHLastPointExprR1C1 / BDHStgPointR1C1 for T0 only.
'
' Design:
'   - Retrieve one value for one security / one field / one T0 date.
'   - Use the BQL field/expression argument explicitly.
'   - Use DATES=<T0> and FILL=PREV to replicate the old BDH behavior:
'       BDH start = T0 - 3 days
'       BDH end   = T0
'       Fill=P
'       last available observation
'
' Important:
'   Different Bloomberg Excel Add-in versions can return BQL as either a scalar
'   or a small result table. The LET/INDEX wrapper extracts the bottom-right cell,
'   which should be the actual value when BQL returns a compact table.
' =============================================================================


Private Function FutFirstBDPMultiFieldFormulaR1C1( _
    ByVal secRef As String, _
    ByVal fields As Variant, _
    Optional ByVal fallbackExpr As String = "") As String


    Dim expr As String
    Dim i As Long


    If fallbackExpr = "" Then fallbackExpr = XlBlank()


    expr = fallbackExpr


    For i = UBound(fields) To LBound(fields) Step -1
        expr = GuardedBDPExprR1C1(secRef, XlText(CStr(fields(i))), expr)
    Next i


    FutFirstBDPMultiFieldFormulaR1C1 = WrapIfPresent(RC(FCOL_CONTRACTCODE), expr)
End Function




Private Function BQLT0ParamsExprR1C1(ByVal cfgT0 As String) As String
    Dim t0Value As Variant
    Dim t0Date As Date
    Dim t0Iso As String


    ' Read the actual T0 date from Config directly in VBA.
    ' This avoids Excel TEXT(...,"yyyy-mm-dd") regional-format problems.
    t0Value = ThisWorkbook.Worksheets(SH_CONFIG).Range(CFG_T0_DATE).value


    If Not IsDate(t0Value) Then
        Err.Raise 9801, , _
            "Invalid T0 date in Config " & CFG_T0_DATE & ". Value found: " & CStr(t0Value)
    End If


    t0Date = DateValue(CDate(t0Value))


    ' VBA Format$ is stable here and will produce e.g. 2026-06-15.
    t0Iso = Format$(t0Date, "yyyy-mm-dd")


    ' Return this as an Excel string literal:
    '   "DATES=2026-06-15;FILL=PREV"
    BQLT0ParamsExprR1C1 = XlText("DATES=" & t0Iso & ";FILL=PREV")
End Function


Private Function BQLPointExprR1C1( _
    ByVal secExpr As String, _
    ByVal fieldExpr As String, _
    ByVal cfgT0 As String _
) As String


    Dim paramsExpr As String
    paramsExpr = BQLT0ParamsExprR1C1(cfgT0)


    BQLPointExprR1C1 = _
        "LET(" & _
            "_q,BQL(" & secExpr & "," & fieldExpr & "," & paramsExpr & ")," & _
            "_r,ROWS(_q)," & _
            "_c,COLUMNS(_q)," & _
            "INDEX(_q,_r,_c)" & _
        ")"


End Function


Private Function BQLDateParamsExprR1C1(ByVal configDateCell As String) As String
    Dim v As Variant
    Dim d As Date
    Dim isoDate As String


    v = ThisWorkbook.Worksheets(SH_CONFIG).Range(configDateCell).value


    If Not IsDate(v) Then
        Err.Raise 9802, , "Invalid date in Config " & configDateCell & ". Value found: " & CStr(v)
    End If


    d = DateValue(CDate(v))
    isoDate = Format$(d, "yyyy-mm-dd")


    BQLDateParamsExprR1C1 = XlText("DATES=" & isoDate & ";FILL=PREV")
End Function


Private Function BQLPointExprR1C1_ByConfigDate(ByVal secExpr As String, ByVal fieldExpr As String, ByVal configDateCell As String) As String
    Dim paramsExpr As String


    paramsExpr = BQLDateParamsExprR1C1(configDateCell)


    BQLPointExprR1C1_ByConfigDate = _
        "IFERROR(" & _
            "LET(" & _
                "_q,BQL(" & secExpr & "," & fieldExpr & "," & paramsExpr & ")," & _
                "_r,ROWS(_q)," & _
                "_c,COLUMNS(_q)," & _
                "INDEX(_q,_r,_c)" & _
            ")," & _
        """"")"
End Function


Private Function BQLPointFormulaR1C1_ByConfigDate(ByVal secExpr As String, ByVal fieldExpr As String, ByVal configDateCell As String) As String
    BQLPointFormulaR1C1_ByConfigDate = _
        "=IF(" & secExpr & "="""",""""," & _
            BQLPointExprR1C1_ByConfigDate(secExpr, fieldExpr, configDateCell) & _
        ")"
End Function


Private Function BQLPointFormulaR1C1( _
    ByVal secExpr As String, _
    ByVal fieldExpr As String, _
    ByVal cfgT0 As String _
) As String


    BQLPointFormulaR1C1 = _
        "=IF(" & secExpr & "="""",""""," & _
            BQLPointExprR1C1(secExpr, fieldExpr, cfgT0) & _
        ")"
End Function




Private Sub AddToUnion(ByRef baseRng As Range, ByVal addRng As Range)
    ' Small helper used to build one combined range from multiple non-contiguous
    ' BQL output ranges.


    If baseRng Is Nothing Then
        Set baseRng = addRng
    Else
        Set baseRng = Application.Union(baseRng, addRng)
    End If
End Sub




' Progress-aware readiness wait.
'
'   The old version used a FLAT deadline: if any cell was still "#N/A Requesting Data"
'   when CFG_BBG_TIMEOUT elapsed it failed - even though the data was still streaming in
'   and finished moments later.  That false-aborted Step 4 on large BDH batches (117
'   curve cells, and 4,800 bond/future staging cells).
'
'   This version waits as long as Bloomberg keeps making PROGRESS (the count of pending
'   cells keeps dropping) and only fails on:
'     * a hard session error (Connection / Authorization / Invalid override), or
'     * a true STALL (no progress for CFG_BBG_TIMEOUT seconds), or
'     * an absolute cap (safety, for genuinely stuck batches).
'   A final settle + authoritative re-scan closes the race where the last cells resolve
'   just as the loop ends.
Private Function WaitForBloombergReadyMany(ByVal wsCfg As Worksheet, ParamArray rngs() As Variant) As Boolean
    If KeepFormulasMode() Then       ' offline/inspection: nothing to wait for, keep formulas
        WaitForBloombergReadyMany = True
        Exit Function
    End If
    Dim stallSec As Long
    stallSec = val(CleanText(wsCfg.Range(CFG_BBG_TIMEOUT).value))
    If stallSec <= 0 Then stallSec = 120
    If stallSec < 30 Then stallSec = 30


    ' Absolute backstop, independent of the progress-based stall reset below.  Lowered from
    ' 6x/600s to 4x/240s as part of the Step 4 OOM fix: with the smaller 25-row staging
    ' batches a wedged batch now releases its open Bloomberg subscriptions far sooner instead
    ' of sitting up to 10 minutes holding memory.  The stall timer still extends genuinely
    ' progressing batches.
    Dim capSec As Long
    capSec = stallSec * 4
    If capSec < 240 Then capSec = 240          ' absolute safety cap (>= 4 min)


    Dim tStart As Single:    tStart = Timer
    Dim tProgress As Single: tProgress = Timer
    Dim prevPending As Long: prevPending = -1
    Dim curPending As Long


    Dim rngArr As Variant
    rngArr = rngs


    Do
        DoEvents
        CalculateRangesMany rngArr


        ' Genuine session failure -> stop now (data is not actually coming).
        If RangeHasHardBBGError(rngArr) Then
            WaitForBloombergReadyMany = False
            Exit Function
        End If


        curPending = CountPendingBBG(rngArr)
        If curPending = 0 Then Exit Do                       ' everything settled


        ' Progress resets the stall timer; large batches keep getting time as long as
        ' cells keep clearing.
        If prevPending < 0 Or curPending < prevPending Then tProgress = Timer
        prevPending = curPending


        If (Timer - tProgress) > stallSec Then Exit Do       ' stalled (no progress)
        If (Timer - tStart) > capSec Then Exit Do            ' absolute cap


        WaitSeconds 1
    Loop


    ' Final settle, then make the authoritative decision from actual cell state.
    DoEvents
    CalculateRangesMany rngArr
    WaitSeconds 1
    DoEvents
    CalculateRangesMany rngArr


    If CountPendingBBG(rngArr) > 0 Then
        WaitForBloombergReadyMany = False
        Exit Function
    End If


    If RangeHasHardBBGError(rngArr) Then
        WaitForBloombergReadyMany = False
        Exit Function
    End If


    WaitForBloombergReadyMany = True
End Function


' Total number of cells still showing a Bloomberg "Requesting"/pending state.
Private Function CountPendingBBG(ByVal rngs As Variant) As Long
    Dim i As Long, n As Long
    Dim rg As Range
    n = 0
    For i = LBound(rngs) To UBound(rngs)
        If TypeName(rngs(i)) = "Range" Then
            Set rg = rngs(i)
            If RangeHasPendingBBG(rg) Then
                Dim c As Range
                For Each c In rg.Cells
                    If InStr(1, CStr(c.Text), "Requesting", vbTextCompare) > 0 Then n = n + 1
                Next c
            End If
        End If
    Next i
    CountPendingBBG = n
End Function


' True only for HARD Bloomberg errors that will not resolve by waiting (session/auth/
' invalid override).  Transient "Requesting" is deliberately excluded.
Private Function RangeHasHardBBGError(ByVal rngs As Variant) As Boolean
    Dim i As Long
    Dim rg As Range
    Dim c As Range
    Dim txt As String
    For i = LBound(rngs) To UBound(rngs)
        If TypeName(rngs(i)) = "Range" Then
            Set rg = rngs(i)
            For Each c In rg.Cells
                txt = CStr(c.Text)
                If InStr(1, txt, "#N/A Connection", vbTextCompare) > 0 _
                    Or InStr(1, txt, "Authorization", vbTextCompare) > 0 _
                    Or InStr(1, txt, "Invalid override field id", vbTextCompare) > 0 _
                    Or InStr(1, txt, "Invalid Field", vbTextCompare) > 0 _
                    Or InStr(1, txt, "Invalid Security", vbTextCompare) > 0 _
                    Or InStr(1, txt, "Field Not Applicable", vbTextCompare) > 0 Then
                    Debug.Print "Hard Bloomberg error at " & c.Worksheet.Name & "!" & c.Address(False, False) & " -> " & txt
                    RangeHasHardBBGError = True
                    Exit Function
                End If
            Next c
        End If
    Next i
    RangeHasHardBBGError = False
End Function


Private Function RangeHasCriticalBBGError(ByVal rngs As Variant) As Boolean


    Dim i As Long
    Dim rg As Range
    Dim c As Range
    Dim txt As String


    For i = LBound(rngs) To UBound(rngs)


        If TypeName(rngs(i)) = "Range" Then


            Set rg = rngs(i)


            For Each c In rg.Cells


                txt = CStr(c.Text)


                If InStr(1, txt, "#N/A Connection", vbTextCompare) > 0 _
                   Or InStr(1, txt, "Authorization", vbTextCompare) > 0 _
                   Or InStr(1, txt, "Invalid override field id", vbTextCompare) > 0 _
                   Or InStr(1, txt, "Invalid Field", vbTextCompare) > 0 _
                   Or InStr(1, txt, "Invalid Security", vbTextCompare) > 0 _
                   Or InStr(1, txt, "Field Not Applicable", vbTextCompare) > 0 Then


                    RangeHasCriticalBBGError = True
                    Exit Function


                End If


            Next c


        End If


    Next i


    RangeHasCriticalBBGError = False


End Function


Private Function RangeHasUnresolvedBBGProblem(ByVal rng As Range) As Boolean


    Dim c As Range
    Dim txt As String


    For Each c In rng.Cells


        txt = CStr(c.Text)


        If InStr(1, txt, "Requesting", vbTextCompare) > 0 Then
            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


        If InStr(1, txt, "#N/A Requesting", vbTextCompare) > 0 Then
            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


        If InStr(1, txt, "#N/A Connection", vbTextCompare) > 0 Then
            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


        If InStr(1, txt, "#N/A Invalid", vbTextCompare) > 0 Then
            Debug.Print "Bloomberg invalid formula at " & c.Worksheet.Name & "!" & c.Address(False, False)
            Debug.Print c.FormulaR1C1


            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


        If InStr(1, txt, "Invalid override field id", vbTextCompare) > 0 Then
            Debug.Print "Invalid override field ID at " & c.Worksheet.Name & "!" & c.Address(False, False)
            Debug.Print c.FormulaR1C1


            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


        If InStr(1, txt, "Authorization", vbTextCompare) > 0 Then
            RangeHasUnresolvedBBGProblem = True
            Exit Function
        End If


    Next c


    RangeHasUnresolvedBBGProblem = False


End Function




' =============================================================================
' CURVE PRIOR / T-1 BQL WRITER
'
' Legacy function name: WriteCurveT0BQLFormulas
'
' Writes BQL formulas directly into official T-1 curve input columns:
'
'   EUR:
'     D = OIS_T-1
'     G = Gov_T-1
'     J = Swap_T-1
'
'   USD:
'     U  = OIS_T-1
'     X  = Gov_T-1
'     AA = Swap_T-1
'
'   GBP:
'     AL = OIS_T-1
'     AO = Gov_T-1
'     AR = Swap_T-1
' =============================================================================


Private Function WriteCurveT0BQLFormulas(ByVal ws As Worksheet) As Range


    Dim cfgT0 As String
    cfgT0 = CfgR1C1(CFG_T0_DATE)


    Dim targetRng As Range


    ' -------------------------------------------------------------------------
    ' EUR curve T0 values
    ' -------------------------------------------------------------------------
    WriteBQLDown_T0 ws, "C", "D", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "F", "G", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "I", "J", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng


    ' -------------------------------------------------------------------------
    ' USD curve T0 values
    ' -------------------------------------------------------------------------
    WriteBQLDown_T0 ws, "T", "U", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "W", "X", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "Z", "AA", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng


    ' -------------------------------------------------------------------------
    ' GBP curve T0 values
    ' -------------------------------------------------------------------------
    WriteBQLDown_T0 ws, "AK", "AL", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "AN", "AO", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng
    WriteBQLDown_T0 ws, "AQ", "AR", CURVE_FIRST_ROW, CURVE_LAST_ROW, BBG_PX_LAST, cfgT0, targetRng


    Set WriteCurveT0BQLFormulas = targetRng


End Function


Private Sub WriteBQLDown_T0( _
    ByVal ws As Worksheet, _
    ByVal tickerCol As String, _
    ByVal outCol As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal fieldName As String, _
    ByVal cfgT0 As String, _
    ByRef targetRng As Range _
)


    Dim tickerColNum As Long
    tickerColNum = ws.Range(tickerCol & "1").Column


    Dim rg As Range
    Set rg = ws.Range(outCol & firstRow & ":" & outCol & lastRow)


    ' Example produced:
    '   =IF(RC3="","",IFERROR(LET(_q,BQL(RC3,"PX_LAST","DATES="&TEXT(INT(Config!B4),"yyyy-mm-dd")&";FILL=PREV"),...), ""))
    '
    ' Security comes from tickerCol.
    ' Output goes into outCol.


    rg.FormulaR1C1 = _
        "=IF(RC" & tickerColNum & "="""",""""," & _
            BQLPointExprR1C1("RC" & tickerColNum, XlText(fieldName), cfgT0) & _
        ")"


    AddToUnion targetRng, rg
End Sub




Private Function WriteSwapsT0BQLFormulas(ByVal ws As Worksheet, ByVal lastRow As Long) As Range


    If lastRow < DATA_ROW Then Exit Function


    Dim f As Long
    Dim targetRng As Range


    f = DATA_ROW


    ' -------------------------------------------------------------------------
    ' AU:AW = prior / T-1 market value from direct / fixed / float securities.
    '
    ' Current convention:
    '   CFG_T1_DATE = Config!B5 = user-facing T0 / current date
    '   CFG_T0_DATE = Config!B4 = user-facing T-1 / prior date
    '
    ' Previous version used:
    '   SW_MARKET_VAL_PRIOR at CFG_T1_DATE
    '
    ' New clearer version uses:
    '   SW_MARKET_VAL at CFG_T0_DATE
    '
    ' AD = RC30 = direct swap ID
    ' AE = RC31 = fixed leg ID
    ' AF = RC32 = float leg ID
    ' -------------------------------------------------------------------------


    ws.Range(WCOL_BQL_NPV_DIRECT_TM1 & f & ":" & WCOL_BQL_NPV_DIRECT_TM1 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_SWAP_DIRECT_ID), XlText(BQL_SWAP_MV), CFG_T0_DATE)
    AddToUnion targetRng, ws.Range(WCOL_BQL_NPV_DIRECT_TM1 & f & ":" & WCOL_BQL_NPV_DIRECT_TM1 & lastRow)


    ws.Range(WCOL_BQL_NPV_FIXED_TM1 & f & ":" & WCOL_BQL_NPV_FIXED_TM1 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_FIXED_LEG_ID), XlText(BQL_SWAP_MV), CFG_T0_DATE)
    AddToUnion targetRng, ws.Range(WCOL_BQL_NPV_FIXED_TM1 & f & ":" & WCOL_BQL_NPV_FIXED_TM1 & lastRow)


    ws.Range(WCOL_BQL_NPV_FLOAT_TM1 & f & ":" & WCOL_BQL_NPV_FLOAT_TM1 & lastRow).FormulaR1C1 = _
        BQLPointFormulaR1C1_ByConfigDate(RC(WCOL_BBG_FLOAT_LEG_ID), XlText(BQL_SWAP_MV), CFG_T0_DATE)
    AddToUnion targetRng, ws.Range(WCOL_BQL_NPV_FLOAT_TM1 & f & ":" & WCOL_BQL_NPV_FLOAT_TM1 & lastRow)


    ' AX = total prior / T-1 NPV.
    ws.Range(WCOL_BQL_NPV_TOTAL_TM1 & f & ":" & WCOL_BQL_NPV_TOTAL_TM1 & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), _
        "IF(ISNUMBER(" & RC(WCOL_BQL_NPV_DIRECT_TM1) & ")," & RC(WCOL_BQL_NPV_DIRECT_TM1) & ",IF(OR(ISNUMBER(" & RC(WCOL_BQL_NPV_FIXED_TM1) & "),ISNUMBER(" & RC(WCOL_BQL_NPV_FLOAT_TM1) & ")),SUM(" & RC(WCOL_BQL_NPV_FIXED_TM1) & ":" & _
        RC(WCOL_BQL_NPV_FLOAT_TM1) & "),""""))")
    AddToUnion targetRng, ws.Range(WCOL_BQL_NPV_TOTAL_TM1 & f & ":" & WCOL_BQL_NPV_TOTAL_TM1 & lastRow)


    ' AY = BQL swap PnL = current/T0 total NPV - prior/T-1 total NPV.
    '
    ' AT = RC46 = current/T0 total NPV
    ' AX = RC50 = prior/T-1 total NPV
    ws.Range(WCOL_BQL_SWAP_PNL & f & ":" & WCOL_BQL_SWAP_PNL & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), _
        "IF(AND(ISNUMBER(" & RC(WCOL_BQL_NPV_TOTAL_T0) & "),ISNUMBER(" & RC(WCOL_BQL_NPV_TOTAL_TM1) & "))," & RC(WCOL_BQL_NPV_TOTAL_T0) & "-" & RC(WCOL_BQL_NPV_TOTAL_TM1) & ","""")")
    AddToUnion targetRng, ws.Range(WCOL_BQL_SWAP_PNL & f & ":" & WCOL_BQL_SWAP_PNL & lastRow)


    ' AB = final swap PnL used by PNL_Attribution.
    '
    ' AY = RC51 = BQL swap PnL
    ws.Range(WCOL_PNL & f & ":" & WCOL_PNL & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), "IF(ISNUMBER(" & RC(WCOL_BQL_SWAP_PNL) & ")," & RC(WCOL_BQL_SWAP_PNL) & ","""")")
    AddToUnion targetRng, ws.Range(WCOL_PNL & f & ":" & WCOL_PNL & lastRow)


    ' AZ = BQL status.
    ws.Range(WCOL_BQL_SWAP_STATUS & f & ":" & WCOL_BQL_SWAP_STATUS & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), _
        "IF(" & RC(WCOL_BBG_SWAP_DIRECT_ID) & "&" & RC(WCOL_BBG_FIXED_LEG_ID) & "&" & RC(WCOL_BBG_FLOAT_LEG_ID) & "="""",""Missing BBG swap IDs"",IF(NOT(ISNUMBER(" & RC(WCOL_BQL_NPV_TOTAL_T0) & _
        ")),""Missing current/T0 NPV"",IF(NOT(ISNUMBER(" & RC(WCOL_BQL_NPV_TOTAL_TM1) & ")),""Missing prior/T-1 NPV"",""OK_BQL"")))")
    AddToUnion targetRng, ws.Range(WCOL_BQL_SWAP_STATUS & f & ":" & WCOL_BQL_SWAP_STATUS & lastRow)


    ' AC = final swap status mirror.
    ws.Range(WCOL_AF_STATUS & f & ":" & WCOL_AF_STATUS & lastRow).FormulaR1C1 = WrapIfPresent(RC(WCOL_DEALID), RC(WCOL_BQL_SWAP_STATUS))
    AddToUnion targetRng, ws.Range(WCOL_AF_STATUS & f & ":" & WCOL_AF_STATUS & lastRow)


    Set WriteSwapsT0BQLFormulas = targetRng


End Function






' =============================================================================
' SAFE RANGE HELPERS
'
' Purpose:
'   Avoid fragile Worksheet.Range("A1:B2,C1:D2") calls in Step 4.
'   Track the last attempted range so Error 1004 tells us exactly what failed.
' =============================================================================


Private Function GetFirstListObject(ByVal ws As Worksheet) As ListObject


    If ws.ListObjects.Count = 0 Then
        Set GetFirstListObject = Nothing
    Else
        Set GetFirstListObject = ws.ListObjects(1)
    End If


End Function


Private Sub RefreshListObject(ByVal lo As ListObject)


    On Error GoTo FallbackRefresh


    If Not lo.QueryTable Is Nothing Then
        lo.QueryTable.Refresh BackgroundQuery:=False
        Exit Sub
    End If


FallbackRefresh:
    Err.Clear
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone


End Sub


Private Function NormalizeRateForModel(ByVal v As Variant) As Variant


    If IsNumeric(v) Then
        If Abs(CDbl(v)) > 1 Then
            NormalizeRateForModel = CDbl(v) / 100
        Else
            NormalizeRateForModel = CDbl(v)
        End If
    Else
        NormalizeRateForModel = v
    End If


End Function


Private Function colNum(ByVal colLetters As String) As Long
    colNum = ThisWorkbook.Worksheets(1).Range(colLetters & "1").Column
End Function


Private Function SafeBlock( _
    ByVal ws As Worksheet, _
    ByVal firstCol As String, _
    ByVal firstRow As Long, _
    ByVal lastCol As String, _
    ByVal lastRow As Long _
) As Range


    If firstRow <= 0 Or lastRow <= 0 Then
        Err.Raise 9701, , "Invalid row number in SafeBlock: " & firstRow & ":" & lastRow
    End If


    If lastRow < firstRow Then
        Err.Raise 9702, , "Invalid row order in SafeBlock: " & firstRow & ":" & lastRow
    End If


    mLastRangeTarget = ws.Name & "!" & firstCol & firstRow & ":" & lastCol & lastRow


    Set SafeBlock = ws.Range( _
        ws.Cells(firstRow, colNum(firstCol)), _
        ws.Cells(lastRow, colNum(lastCol)) _
    )
End Function


Private Function SafeColRange( _
    ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long _
) As Range


    Set SafeColRange = SafeBlock(ws, colLetter, firstRow, colLetter, lastRow)
End Function


Private Sub SafeClearBlock( _
    ByVal ws As Worksheet, _
    ByVal firstCol As String, _
    ByVal firstRow As Long, _
    ByVal lastCol As String, _
    ByVal lastRow As Long _
)


    SafeBlock(ws, firstCol, firstRow, lastCol, lastRow).ClearContents
End Sub


Private Sub SafeCopyValues( _
    ByVal srcWs As Worksheet, _
    ByVal srcCol As String, _
    ByVal dstWs As Worksheet, _
    ByVal dstCol As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long _
)


    SafeColRange(dstWs, dstCol, firstRow, lastRow).value = _
        SafeColRange(srcWs, srcCol, firstRow, lastRow).value
End Sub










' =============================================================================
' BUTTON 5 - RECALCULATE PNL
' =============================================================================
' =============================================================================
' WRITE ALL FORMULAS  (no Bloomberg, no OPICS, no freeze)  -- button *
'
' Writes every derived formula AND every Bloomberg-formula SHELL into all sheets
' by calling the existing formula writers, WITHOUT triggering/waiting on Bloomberg
' and without freezing.  Use it to guarantee the formulas are present and auditable
' in the workbook even when a data fetch (OPICS or Bloomberg) fails: the Bloomberg
' cells display #N/A until a successful refresh, but the formula is always in place.
'
' Needs the input rows already loaded (Buttons 1 & 2).  It never fetches data, so
' it is safe to run offline / with Bloomberg disconnected.  Each sheet is written
' inside its own error trap so one section failing does not block the others.
' =============================================================================
' BLOOMBERG OUTPUT RANGES
'
' One place that answers "which cells on this sheet carry a Bloomberg formula",
' covering BOTH snapshots:
'   current / T0   -> BDP (live) columns written by the Write*_T1_BDP_* writers
'   prior   / T-1  -> BQL point-in-time columns written by the Write*T0BQL*
'                     writers, dated from Config!B4
'
' RefreshMarketData waits on exactly these ranges, so it never has to guess and
' never has to rewrite a formula just to find out where the formulas are.
'
' These lists MIRROR the writers.  tools/check_bbg_ranges.py fails the build if
' a writer touches a Bloomberg column that its range builder does not list, so
' the mirror cannot rot silently.
' =============================================================================

Private Function CurveBbgRange(ByVal ws As Worksheet) As Range

    Dim rg As Range
    Dim r1 As String
    Dim r2 As String

    r1 = CStr(CURVE_FIRST_ROW)
    r2 = CStr(CURVE_LAST_ROW)

    ' EUR: ticker / prior (BQL) / current (BDP) triplets
    AddToUnion rg, ws.Range(CVCOL_EUR_ESTR_OIS_TM1 & r1 & ":" & CVCOL_EUR_ESTR_OIS_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_EUR_GOV_TM1 & r1 & ":" & CVCOL_EUR_GOV_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_EUR_EURIBOR_SWAP_TM1 & r1 & ":" & CVCOL_EUR_EURIBOR_SWAP_T0 & r2)

    ' USD
    AddToUnion rg, ws.Range(CVCOL_USD_OIS_TM1 & r1 & ":" & CVCOL_USD_OIS_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_USD_GOV_TM1 & r1 & ":" & CVCOL_USD_GOV_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_USD_SWAP_TM1 & r1 & ":" & CVCOL_USD_SWAP_T0 & r2)

    ' GBP
    AddToUnion rg, ws.Range(CVCOL_GBP_OIS_TM1 & r1 & ":" & CVCOL_GBP_OIS_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_GBP_GOV_TM1 & r1 & ":" & CVCOL_GBP_GOV_T0 & r2)
    AddToUnion rg, ws.Range(CVCOL_GBP_SWAP_TM1 & r1 & ":" & CVCOL_GBP_SWAP_T0 & r2)

    Set CurveBbgRange = rg

End Function


Private Function BondsBbgRange( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long) As Range

    Dim rg As Range
    Dim f As String

    If lastRow < BOND_DATA_ROW Then Exit Function
    f = CStr(BOND_DATA_ROW)

    ' resolved security ticker (BDP-driven parse)
    AddToUnion rg, ws.Range(BCOL_BBG_TICKER & f & ":" & BCOL_BBG_TICKER & lastRow)

    ' current / T0  (BDP)
    AddToUnion rg, ws.Range(BCOL_FX_T0 & f & ":" & BCOL_FX_T0 & lastRow)
    AddToUnion rg, ws.Range(BCOL_CLEANPX_T0 & f & ":" & BCOL_ASW_T0 & lastRow)
    AddToUnion rg, ws.Range(BCOL_OAS_T0 & f & ":" & BCOL_OAS_T0 & lastRow)
    AddToUnion rg, ws.Range(BCOL_OAS_MODDURATION_RAW & f & ":" & BCOL_OAS_CONVEXITY & lastRow)
    AddToUnion rg, ws.Range(BCOL_PRICING_SOURCE & f & ":" & BCOL_BENCHMARK_NAME & lastRow)
    AddToUnion rg, ws.Range(BCOL_FUNDRATE_T0 & f & ":" & BCOL_FUNDRATE_T0 & lastRow)
    AddToUnion rg, ws.Range(BCOL_DAY_CNT_DES & f & ":" & BCOL_DV01_OPENING_EUR & lastRow)

    ' prior / T-1  (BQL at Config!B4)
    AddToUnion rg, ws.Range(BCOL_FX_TM1 & f & ":" & BCOL_FX_TM1 & lastRow)
    AddToUnion rg, ws.Range(BCOL_CLEANPX_TM1 & f & ":" & BCOL_ASW_TM1 & lastRow)
    AddToUnion rg, ws.Range(BCOL_OAS_TM1 & f & ":" & BCOL_OAS_TM1 & lastRow)
    AddToUnion rg, ws.Range(BCOL_FUNDRATE_TM1 & f & ":" & BCOL_FUNDRATE_TM1 & lastRow)

    Set BondsBbgRange = rg

End Function


Private Function FuturesBbgRange( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long) As Range

    Dim rg As Range
    Dim f As String

    If lastRow < DATA_ROW Then Exit Function
    f = CStr(DATA_ROW)

    ' current / T0  (BDP)
    AddToUnion rg, ws.Range(FCOL_DELIVDATE & f & ":" & FCOL_DELIVDATE & lastRow)
    AddToUnion rg, ws.Range(FCOL_FX_T0 & f & ":" & FCOL_FUTPX_T0 & lastRow)
    AddToUnion rg, ws.Range(FCOL_CTD_DIRTYPX_T0 & f & ":" & FCOL_CTD_TICKER_TM1 & lastRow)
    AddToUnion rg, ws.Range(FCOL_FUT_VAL_PT & f & ":" & FCOL_HEDGEUNITDV01 & lastRow)
    AddToUnion rg, ws.Range(FCOL_IMPLIEDREPO_BBG & f & ":" & FCOL_GROSSBASIS_BBG & lastRow)

    ' prior / T-1  (BQL at Config!B4)
    AddToUnion rg, ws.Range(FCOL_FUTPX_TM1 & f & ":" & FCOL_FUTPX_TM1 & lastRow)

    Set FuturesBbgRange = rg

End Function


Private Function SwapsBbgRange( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long) As Range

    Dim rg As Range
    Dim f As String

    If lastRow < DATA_ROW Then Exit Function
    f = CStr(DATA_ROW)

    ' terms pulled from Bloomberg
    AddToUnion rg, ws.Range(WCOL_FIXEDRATE & f & ":" & WCOL_FLOATSPREAD & lastRow)
    AddToUnion rg, ws.Range(WCOL_PAY_FLT_RATE_IDX & f & ":" & WCOL_PAY_FLT_RATE_IDX & lastRow)
    AddToUnion rg, ws.Range(WCOL_DV01_BBG & f & ":" & WCOL_DV01_BBG & lastRow)

    ' current / T0 and prior / T-1 NPV legs, PnL and status
    AddToUnion rg, ws.Range(WCOL_BQL_NPV_DIRECT_T0 & f & ":" & WCOL_BQL_SWAP_STATUS & lastRow)
    AddToUnion rg, ws.Range(WCOL_PNL & f & ":" & WCOL_AF_STATUS & lastRow)

    Set SwapsBbgRange = rg

End Function


' =============================================================================
' BUTTON - REFRESH MARKET DATA  (both snapshots, no freezing)
'
' The workbook holds LIVE formulas for both snapshots:
'   current / T0  = BDP, reads whatever Bloomberg has now
'   prior   / T-1 = BQL dated from Config!B4, so it reproduces exactly
'
' Because the prior snapshot is a dated query rather than a captured value,
' there is nothing to freeze: re-running this button reproduces the same T-1
' figures.  That is the whole reason the old "Freeze T0 snapshot" step is gone -
' it converted reproducible formulas into static values, which then had to be
' guarded, warned about, and un-frozen whenever the formulas were rewritten.
'
' This button does NOT write formulas.  If a sheet has no formulas in it yet,
' run WriteAllModelFormulas first.
' =============================================================================

' Ask Bloomberg for both snapshots, wait for each section to settle, then
' rebuild.  Returns a per-section summary; raises nothing.
'
' This was the "refresh market data" button.  It is not a button any more - it
' is the tail of Buttons 1 and 2, because there is nothing here a user should
' have to decide.  Two parts of it are load-bearing and neither is obvious:
'
'   THE WAIT.  BDP and BQL resolve asynchronously.  Recalculating straight after
'   asking for them computes the whole book against "#N/A Requesting Data", and
'   the result is not an error - it is blanks, which read as a small PnL rather
'   than a missing one.  RefreshWaitStatus blocks per section so a timeout names
'   the section that did not answer.
'
'   THE FULL REBUILD.  InterpOIS, InterpGov, InterpSwap and BondPullToParPrice
'   read OIS_Curves through the object model, not through cell references, so
'   Excel has no dependency edge from a curve cell to the bonds that use it.
'   Application.Calculate leaves every one of them holding the previous run's
'   number.  CalculateFullRebuild is the only thing that re-evaluates them.
Private Function RefreshMarketDataCore() As String

    Dim wb As Workbook
    Dim wsCfg As Worksheet
    Dim wsOIS As Worksheet
    Dim wsBnd As Worksheet
    Dim wsFut As Worksheet
    Dim wsSw As Worksheet

    Dim curveRng As Range
    Dim bondRng As Range
    Dim futRng As Range
    Dim swapRng As Range

    Dim lastBondRow As Long
    Dim lastFutRow As Long
    Dim lastSwapRow As Long

    Dim summary As String

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldStatusBar As Variant

    Set wb = ThisWorkbook
    Set wsCfg = wb.Worksheets(SH_CONFIG)
    Set wsOIS = wb.Worksheets(SH_OIS)
    Set wsBnd = wb.Worksheets(SH_BONDS)
    Set wsFut = wb.Worksheets(SH_FUTURES)
    Set wsSw = wb.Worksheets(SH_SWAPS)

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Row counts come from the sheets, every time.  The number of bonds, futures
    ' and swaps changes with the book; nothing here may assume a fixed count.
    lastBondRow = LastBondDataRow(wsBnd)
    lastFutRow = LastFutureDataRow(wsFut)
    lastSwapRow = LastSwapDataRow(wsSw)

    Set curveRng = CurveBbgRange(wsOIS)
    Set bondRng = BondsBbgRange(wsBnd, lastBondRow)
    Set futRng = FuturesBbgRange(wsFut, lastFutRow)
    Set swapRng = SwapsBbgRange(wsSw, lastSwapRow)

    ' -------------------------------------------------------------------------
    ' 1) Ask Bloomberg to re-pull everything, once.
    ' -------------------------------------------------------------------------
    Application.StatusBar = "Refresh: asking Bloomberg for both snapshots..."
    TriggerBloombergRefreshOnly

    ' -------------------------------------------------------------------------
    ' 2) Wait per section, so a failure names the section that failed rather
    '    than the whole book.
    ' -------------------------------------------------------------------------
    Application.StatusBar = "Refresh: waiting for curves..."
    summary = summary & RefreshWaitStatus( _
        "Curves", wsCfg, curveRng)

    Application.StatusBar = "Refresh: waiting for bonds..."
    summary = summary & RefreshWaitStatus( _
        "Bonds  (" & CStr(BondRowCount(lastBondRow)) & " rows)", wsCfg, bondRng)

    Application.StatusBar = "Refresh: waiting for futures..."
    summary = summary & RefreshWaitStatus( _
        "Futures (" & CStr(HedgeRowCount(lastFutRow)) & " rows)", wsCfg, futRng)

    Application.StatusBar = "Refresh: waiting for swaps..."
    summary = summary & RefreshWaitStatus( _
        "Swaps  (" & CStr(HedgeRowCount(lastSwapRow)) & " rows)", wsCfg, swapRng)

    ' -------------------------------------------------------------------------
    ' 3) Recalculate everything that depends on the new values.
    '
    ' CalculateFullRebuild, not Calculate: several model quantities come from
    ' UDFs that read the curve sheet directly (InterpOIS / InterpGov / InterpSwap
    ' and BondPullToParPrice).  Excel cannot see those reads in its dependency
    ' graph, so an ordinary recalculation leaves them holding the previous
    ' refresh's numbers.  This is the single most important line in the
    ' procedure and the reason the old "recalculate PnL" button existed at all.
    ' -------------------------------------------------------------------------
    Application.StatusBar = "Refresh: recalculating..."
    Application.Calculation = oldCalc
    Application.CalculateFullRebuild

    ' -------------------------------------------------------------------------
    ' 4) Post-refresh checks.
    '
    ' A curve cell that came back empty or errored while its ticker is present
    ' is the failure that silently poisons everything downstream: the bond
    ' spread columns just go blank and the PnL looks "small" rather than broken.
    ' Say so here rather than leaving it to be noticed on the Dashboard.
    ' -------------------------------------------------------------------------
    If CurrentCurveResultHasProblem(wsOIS) Then
        summary = summary & _
            "  [CHECK]   One or more curve points did not return a value " & _
            "despite having a ticker - see " & SH_OIS & "." & vbCrLf
    End If

    LogDiagnosticsForRanges "RefreshMarketData", _
        curveRng, bondRng, futRng, swapRng

    wsCfg.Range(CFG_LAST_BBG).value = Now()

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar

    RefreshMarketDataCore = summary

    Exit Function

CleanFail:

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar

    RefreshMarketDataCore = "  - Market data: " & CaptureErrInfo() & vbCrLf

End Function


' Wait on one section and return a one-line status for the summary box.
' An empty range means the section has no rows, which is normal, not a failure.
Private Function RefreshWaitStatus( _
    ByVal sectionName As String, _
    ByVal wsCfg As Worksheet, _
    ByVal rng As Range) As String

    If rng Is Nothing Then
        RefreshWaitStatus = "  [SKIP]    " & sectionName & "  -  no rows" & vbCrLf
        Exit Function
    End If

    If WaitForBloombergReadyMany(wsCfg, rng) Then
        RefreshWaitStatus = "  [OK]      " & sectionName & vbCrLf
    Else
        RefreshWaitStatus = "  [PARTIAL] " & sectionName & _
            "  -  some cells did not resolve; they show #N/A" & vbCrLf
    End If

End Function


' Row counts for the status box.  Both return 0 rather than a negative count
' when the sheet is empty.
Private Function BondRowCount(ByVal lastRow As Long) As Long
    If lastRow < BOND_DATA_ROW Then
        BondRowCount = 0
    Else
        BondRowCount = lastRow - BOND_DATA_ROW + 1
    End If
End Function

Private Function HedgeRowCount(ByVal lastRow As Long) As Long
    If lastRow < DATA_ROW Then
        HedgeRowCount = 0
    Else
        HedgeRowCount = lastRow - DATA_ROW + 1
    End If
End Function


' -----------------------------------------------------------------------------
' Bond Bloomberg ticker + its status column.
'
' Extracted from the old Step 3B so that WriteAllModelFormulas installs it like
' every other formula.  Previously the ticker column was written ONLY by the
' Bloomberg refresh step, so a workbook that had had all its formulas written
' still had an empty BBG_Ticker column and every bond pull returned blank.
' -----------------------------------------------------------------------------
Private Sub WriteBondTickerFormulas( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long)

    If lastRow < BOND_DATA_ROW Then Exit Sub

    ws.Range(BCOL_BBG_TICKER & BOND_DATA_ROW & ":" & _
             BCOL_BBG_TICKER & lastRow).FormulaR1C1 = BBGParseFormulaR1C1()

    ws.Range(BCOL_TICKER_STATUS & BOND_DATA_ROW & ":" & _
             BCOL_TICKER_STATUS & lastRow).FormulaR1C1 = _
        WrapIfPresent(RC(BCOL_ISIN), _
            "IF(OR(" & RC(BCOL_BBG_TICKER) & "=""""," & _
            RC(BCOL_BBG_TICKER) & "=""UNKNOWN""),""UNKNOWN"",""OK"")")

End Sub


' =============================================================================
' WriteAllModelFormulas is gone, and so is the button that ran it.
'
' It wrote every formula on every sheet in one pass, which sounds like the safe
' thing to have and was in fact the thing that broke workbooks.  A formula is
' only ever missing because a row appeared without one, and only the load knows
' how many rows there now are - so a separate "write formulas" step created a
' state (rows loaded, formulas not written, or worse: formulas written for last
' week's row count) that is always wrong and that the sheet could not report.
'
' Its four sections now live where the rows they fill are created:
'
'   Curves + Bonds  ->  Button1_LoadBonds
'   Futures + Swaps ->  Button2_LoadHedgesAndAttribute
'   PNL_Attribution ->  Button2_LoadHedgesAndAttribute
'
' Not one writer changed.  WriteBondsCalculatedFormulas_Efficient and the rest
' are the same procedures called in the same order; only the thing calling them
' moved.  See docs/BUTTONS.md.
' =============================================================================




' =============================================================================
' PNL_ATTRIBUTION SECTION WRITER  (owns its error handler)
'
' Returns "" on success, or a one-line problem string naming the stage that
' failed, so the caller never has to guess which part of the section died.
' =============================================================================


Private Function WritePNLSectionSafe( _
    ByVal wsPnl As Worksheet, _
    ByVal wsBnd As Worksheet, _
    ByVal lastBondRow As Long) As String


    Dim stageText As String


    On Error GoTo PnlFail


    stageText = "writing row " & CStr(PNL_HEADER_ROW) & " headers"
    SetupPNLAttributionHeaders wsPnl


    stageText = "validating row " & CStr(PNL_HEADER_ROW) & " headers"
    AssertPNLDashboardContract wsPnl


    If lastBondRow < BOND_DATA_ROW Then
        WritePNLSectionSafe = _
            "  - PNL_Attribution: no bond rows found on " & SH_BONDS & _
            "; headers written, no attribution rows." & vbCrLf
        Exit Function
    End If


    stageText = "writing attribution rows"
    WritePNLAttributionRows wsPnl, wsBnd


    WritePNLSectionSafe = ""
    Exit Function


PnlFail:


    WritePNLSectionSafe = _
        "  - PNL_Attribution (" & stageText & "): Error " & _
        CStr(Err.Number) & ": " & Err.Description & vbCrLf


End Function


' =============================================================================
' PNL_ATTRIBUTION -> DASHBOARD HEADER CONTRACT
'
' The Dashboard resolves every column it reads by header TEXT
' (DashPnlCol -> Rows(4).Find, LookAt:=xlWhole).  Drift between the strings
' written by SetupPNLAttributionHeaders and the strings the Dashboard looks up
' is invisible at write time and a hard error 9901 at Dashboard build time.
'
' This asserts the 31 header cells the Dashboard actually depends on, at the
' columns they must occupy.  Callable standalone at any time.
'
' It deliberately does NOT assert the other 54 headers: those are not part of
' the Dashboard contract and may be renamed freely.
'
' NOTE ON STYLE: the spec is loaded one statement per column rather than as a
' single Array(...) literal.  85 columns would need more than VBA's limit of
' 25 line-continuation characters per logical statement.
' =============================================================================




Public Sub AssertPNLDashboardContract(Optional ByVal wsIn As Worksheet = Nothing)

    ' Does the sheet in front of us still match PnlLayout?
    '
    ' Checked against the DISPLAY LABEL, because that is what the header writer
    ' puts in row 4 and therefore the only thing that can drift.  A mismatch
    ' means somebody typed over row 4 by hand, or the sheet predates a change to
    ' the table.  Either way the fix is to re-run the header writer, not to edit
    ' the cell back.
    '
    ' This no longer decides whether the Dashboard can find its columns - the
    ' published names do that, and they key off the column LETTER.  A wrong
    ' label is now a cosmetic problem the desk should know about rather than a
    ' hard stop, so this reports and does not raise.

    Dim ws As Worksheet
    Dim spec As Collection
    Dim item As Variant
    Dim colLetter As String
    Dim expected As String
    Dim actual As String
    Dim badList As String
    Dim badCount As Long


    If wsIn Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(SH_PNL)
    Else
        Set ws = wsIn
    End If


    Set spec = PnlLayout()


    For Each item In spec


        colLetter = CStr(item(0))
        expected = CStr(item(2))
        actual = CleanText(ws.Range(colLetter & PNL_HEADER_ROW).value)


        If StrComp(actual, expected, vbTextCompare) <> 0 Then
            badCount = badCount + 1
            badList = badList & "  " & colLetter & CStr(PNL_HEADER_ROW) & _
                      " = [" & actual & "]  expected [" & expected & "]" & vbCrLf
        End If


    Next item


    If badCount > 0 Then
        MsgBox CStr(badCount) & " PNL_Attribution header label(s) differ from " & _
               "PnlLayout:" & vbCrLf & vbCrLf & badList & vbCrLf & _
               "The Dashboard still works - it reads the published column " & _
               "names, not these labels." & vbCrLf & _
               "Re-run Button 2 to write them back.", vbExclamation
    ElseIf wsIn Is Nothing Then
        MsgBox "PNL_Attribution header labels match PnlLayout (" & _
               CStr(spec.Count) & " columns).", vbInformation
    End If


End Sub




' =============================================================================
' SQL QUERY BUILDERS  (verbatim from original; unchanged - OPICS schema is correct)
' =============================================================================


Private Function GetBondSelectQuery( _
    ByVal allInvTypes As String, _
    ByVal isinType As String, _
    ByVal asOf As String, _
    Optional ByVal branchCode As String = "") As String
    Dim s As String
    s = s & "SELECT isin.SECALTID AS ISIN,"
    s = s & " sm.DESCR AS Name,"
    s = s & " sp.CCY AS Currency,"
    s = s & " sm.COUPRATE_8 AS Coupon_Rate,"
    s = s & " sm.INTPAYCYCLE AS CouponFreq,"
    s = s & " sm.MDATE AS MaturityDate,"
    s = s & " sp.QTY AS Notional,"
    s = s & " CASE"
    s = s & " WHEN ABS(ISNULL(sp.QTY,0))>0 AND ABS(ISNULL(tp.TDBOOKVALUE,0))>0 THEN ABS(tp.TDBOOKVALUE)/NULLIF(ABS(sp.QTY),0)*100"
    s = s & " WHEN ABS(ISNULL(sp.QTY,0))>0 AND ABS(ISNULL(tp.ADJTDBOOKVAL,0))>0 THEN ABS(tp.ADJTDBOOKVAL)/NULLIF(ABS(sp.QTY),0)*100"
    s = s & " ELSE ISNULL(sp.PURCHAVGCOST,0) END AS AvgCost,"
    s = s & " sp.SETTDATE AS EntryDate,"
    s = s & " sp.PORT AS Portfolio,"
    s = s & " LTRIM(RTRIM(sp.INVTYPE)) AS AcctgCat,"
    s = s & " sp.BR AS Branch,"
    s = s & " sm.PRODUCT AS Product,"
    s = s & " ISNULL(rp.REPORATE_8,0) AS FundingRate_Dec,"
    s = s & " ISNULL(tp.TDBOOKVALUE,0) AS BookVal_EUR"
    s = s & " FROM dbo.SPOS sp"
    s = s & " INNER JOIN dbo.SECM sm ON sm.SECID=sp.SECID"
    s = s & " LEFT JOIN dbo.ASID isin ON isin.SECID=sp.SECID AND isin.SECIDTYPE='" & EscapeSQL(isinType) & "'"
    s = s & " LEFT JOIN dbo.TPOS tp ON tp.BR=sp.BR AND tp.COST=sp.COST AND tp.PORT=sp.PORT AND tp.SECID=sp.SECID AND tp.INVTYPE=sp.INVTYPE"
    s = s & " LEFT JOIN (SELECT r2.SECID,r2.PORT,AVG(r2.REPORATE_8) AS REPORATE_8 FROM dbo.RPDT r2"
    s = s & " WHERE r2.VDATE<='" & asOf & "' AND r2.MDATE>='" & asOf & "' GROUP BY r2.SECID,r2.PORT) rp"
    s = s & " ON rp.SECID=sp.SECID AND rp.PORT=sp.PORT"
    s = s & " WHERE sp.QTY<>0"
    s = s & " AND sp.PORT='PORT'"


    If CleanText(branchCode) <> "" Then
        s = s & " AND LTRIM(RTRIM(sp.BR))='" & EscapeSQL(branchCode) & "'"
    End If


    s = s & " AND LTRIM(RTRIM(sp.INVTYPE)) IN (" & allInvTypes & ")"
    s = s & " AND LTRIM(RTRIM(sp.INVTYPE)) <> 'T'"
    s = s & " AND LTRIM(RTRIM(sm.PRODUCT)) IN ('OBR','SECUR')"
    s = s & " ORDER BY sp.CCY,LTRIM(RTRIM(sp.INVTYPE)),sp.COST,sm.MDATE"
    GetBondSelectQuery = s
End Function


Private Function GetSwapSelectQuery(allPorts As String, isinList As String, isinType As String, asOf As String) As String
    Dim portFilter As String, isinFilter As String, whereClause As String, s As String
    If allPorts <> "" Then portFilter = "LTRIM(RTRIM(sw.PORT)) IN (" & allPorts & ")"
    If isinList <> "" Then isinFilter = "sw.UNDSECID IN (SELECT SECID FROM dbo.ASID WHERE SECALTID IN (" & isinList & ") AND SECIDTYPE='" & EscapeSQL(isinType) & "')"


    If portFilter <> "" And isinFilter <> "" Then
        whereClause = "(" & portFilter & " OR " & isinFilter & ")"
    ElseIf portFilter <> "" Then
        whereClause = portFilter
    ElseIf isinFilter <> "" Then
        whereClause = isinFilter
    Else
        whereClause = "1=1"
    End If


    s = s & "SELECT sw.DEALNO AS DealID,"
    s = s & " NULL AS Ccy,NULL AS Notional,NULL AS FixedRate,NULL AS FloatIndex,NULL AS FloatSpread,"
    s = s & " sw.STARTDATE AS StartDate,sw.MATDATE AS EndDate,sw.NETPAYIND AS PayFixed,sw.PORT AS Portfolio,"
    s = s & " (SELECT TOP 1 a2.SECALTID FROM dbo.ASID a2 WHERE a2.SECID=sw.UNDSECID AND a2.SECIDTYPE='" & EscapeSQL(isinType) & "') AS LinkedISIN,"
    s = s & " c.SN AS CptyName"
    s = s & " FROM dbo.SWDH sw"
    s = s & " LEFT JOIN dbo.CUST c ON c.CNO=sw.CNO"
    s = s & " WHERE LTRIM(RTRIM(CONVERT(varchar(10),sw.VERIND)))='1'"
    s = s & " AND LTRIM(RTRIM(sw.PRODUCT))='SWAP'"
    s = s & " AND sw.MATDATE>'" & asOf & "'"
    s = s & " AND " & whereClause
    s = s & " ORDER BY sw.PORT,sw.MATDATE"
    GetSwapSelectQuery = s
End Function


Private Function GetFuturesSelectQuery(allPorts As String, isinType As String, asOf As String) As String
    Dim s As String
    s = s & "WITH CTD AS (SELECT BR,CONTCODE,DELVDATE,CONVFACTOR_8,UNDSECID,"
    s = s & " ROW_NUMBER() OVER(PARTITION BY BR,CONTCODE ORDER BY DELVDATE) AS RN"
    s = s & " FROM dbo.FDEL WHERE DELVDATE>='" & asOf & "'),"
    s = s & " AE AS (SELECT BR,PORT,COST,CONTCODE,"
    s = s & " SUM(NUMCONT*CONTPRICE_8)/NULLIF(SUM(ABS(NUMCONT)),0) AS AvgPx"
    s = s & " FROM dbo.FFDH WHERE VERIND='Y' AND (REVREASON IS NULL OR REVREASON='')"
    s = s & " GROUP BY BR,PORT,COST,CONTCODE)"
    s = s & " SELECT fp.CONTCODE AS ContractCode,fc.EXCHANGE AS Exchange,fc.CCY AS Currency,"
    s = s & " fp.NUMCONT AS Contracts,fc.FACEVALUE AS FaceValue,ctd.DELVDATE AS DelivDate,"
    s = s & " isin.SECALTID AS CTD_ISIN,ctd.CONVFACTOR_8 AS CTD_CF,fp.PORT AS Portfolio,"
    s = s & " (SELECT TOP 1 a3.SECALTID FROM dbo.ASID a3 WHERE a3.SECID=("
    s = s & " SELECT TOP 1 fh2.UNDSECID FROM dbo.FFDH fh2 WHERE fh2.BR=fp.BR AND fh2.CONTCODE=fp.CONTCODE"
    s = s & " AND fh2.UNDSECID IS NOT NULL AND fh2.UNDSECID<>'' ORDER BY fh2.DEALDATE DESC)"
    s = s & " AND a3.SECIDTYPE='" & EscapeSQL(isinType) & "') AS LinkedISIN,"
    s = s & " ae.AvgPx AS AvgEntryPx"
    s = s & " FROM dbo.FPOS fp"
    s = s & " INNER JOIN dbo.FCON fc ON fc.BR=fp.BR AND fc.CONTCODE=fp.CONTCODE AND fc.PRODUCT=fp.PRODUCT"
    s = s & " LEFT JOIN CTD ctd ON ctd.BR=fp.BR AND ctd.CONTCODE=fp.CONTCODE AND ctd.RN=1"
    s = s & " LEFT JOIN dbo.ASID isin ON isin.SECID=ctd.UNDSECID AND isin.SECIDTYPE='" & EscapeSQL(isinType) & "'"
    s = s & " LEFT JOIN AE ae ON ae.BR=fp.BR AND ae.PORT=fp.PORT AND ae.CONTCODE=fp.CONTCODE"
    s = s & " WHERE fp.NUMCONT<>0 AND fp.PORT IN (" & allPorts & ")"
    s = s & " ORDER BY fc.CCY,fp.PORT,ctd.DELVDATE"
    GetFuturesSelectQuery = s
End Function




' =============================================================================
' INTERPOLATION UDFs  (completed - original was truncated mid-function)
'   OIS_Curves tenor/rate layout (T0 then T1 column per block):
'     EUR yrs=B  OIS D/E  Gov G/H  Swap J/K
'     USD yrs=S  OIS U/V  Gov X/Y  Swap AA/AB
'     GBP yrs=AJ OIS AL/AM Gov AO/AP Swap AR/AS
' =============================================================================


Public Function InterpOIS(ccy As String, years As Double, useT1 As Boolean) As Variant
    InterpOIS = InterpCurveValue(ccy, "OIS", years, useT1)
End Function
Public Function InterpGov(ccy As String, years As Double, useT1 As Boolean) As Variant
    InterpGov = InterpCurveValue(ccy, "GOV", years, useT1)
End Function
Public Function InterpSwap(ccy As String, years As Double, useT1 As Boolean) As Variant
    InterpSwap = InterpCurveValue(ccy, "SWAP", years, useT1)
End Function




Private Function InterpCurveValue(ccy As String, curveType As String, years As Double, useT1 As Boolean) As Variant
    On Error GoTo ErrHandler

    Dim yRange As Range, rRange As Range

    If Not CurveRanges(ccy, curveType, useT1, yRange, rRange) Then
        InterpCurveValue = CVErr(xlErrNA)
        Exit Function
    End If

    InterpCurveValue = LinInterp(yRange, rRange, years)
    Exit Function

ErrHandler:
    InterpCurveValue = CVErr(xlErrNA)
End Function


' =============================================================================
' CURVE RANGE RESOLVER  (single source of truth for "where does this curve live")
'
' Resolves (currency, curve type, which snapshot) to the tenor column and the
' rate column on OIS_Curves.  Both consumers go through here:
'   InterpCurveValue    - one point at a time, for worksheet formulas
'   BondPullCurveNodes  - the whole curve into arrays, for pull-to-par
'
' Keeping one copy matters because the three currencies each have their own
' tenor column (B / S / AJ); a curve added or moved used to need the same edit
' in two places, and the second one was easy to miss.
'
' NOTE ON THE useT1 FLAG: legacy naming.  useT1 = True means the CURRENT
' reporting snapshot (user-facing T0, column *_T0); False means the PRIOR
' snapshot (user-facing T-1, column *_TM1).  See the DATE CONVENTION WARNING at
' the top of this module.
' =============================================================================
Private Function CurveRanges( _
    ByVal ccy As String, _
    ByVal curveType As String, _
    ByVal useT1 As Boolean, _
    ByRef yRange As Range, _
    ByRef rRange As Range) As Boolean

    Dim ws As Worksheet
    Dim yearsCol As String
    Dim rateCol As String
    Dim ct As String

    On Error GoTo Fail

    Set ws = ThisWorkbook.Worksheets(SH_OIS)
    ct = UCase$(Trim$(curveType))

    Select Case UCase$(Trim$(ccy))

        Case "EUR"
            yearsCol = CVCOL_YEARS
            If ct = "OIS" Then
                rateCol = IIf(useT1, CVCOL_EUR_ESTR_OIS_T0, CVCOL_EUR_ESTR_OIS_TM1)
            ElseIf ct = "GOV" Then
                rateCol = IIf(useT1, CVCOL_EUR_GOV_T0, CVCOL_EUR_GOV_TM1)
            Else
                rateCol = IIf(useT1, CVCOL_EUR_EURIBOR_SWAP_T0, CVCOL_EUR_EURIBOR_SWAP_TM1)
            End If

        Case "USD"
            yearsCol = CVCOL_YEARS_S
            If ct = "OIS" Then
                rateCol = IIf(useT1, CVCOL_USD_OIS_T0, CVCOL_USD_OIS_TM1)
            ElseIf ct = "GOV" Then
                rateCol = IIf(useT1, CVCOL_USD_GOV_T0, CVCOL_USD_GOV_TM1)
            Else
                rateCol = IIf(useT1, CVCOL_USD_SWAP_T0, CVCOL_USD_SWAP_TM1)
            End If

        Case "GBP"
            yearsCol = CVCOL_YEARS_AJ
            If ct = "OIS" Then
                rateCol = IIf(useT1, CVCOL_GBP_OIS_T0, CVCOL_GBP_OIS_TM1)
            ElseIf ct = "GOV" Then
                rateCol = IIf(useT1, CVCOL_GBP_GOV_T0, CVCOL_GBP_GOV_TM1)
            Else
                rateCol = IIf(useT1, CVCOL_GBP_SWAP_T0, CVCOL_GBP_SWAP_TM1)
            End If

        Case Else
            GoTo Fail

    End Select

    Set yRange = ws.Range(yearsCol & CURVE_FIRST_ROW & ":" & yearsCol & CURVE_LAST_ROW)
    Set rRange = ws.Range(rateCol & CURVE_FIRST_ROW & ":" & rateCol & CURVE_LAST_ROW)

    CurveRanges = True
    Exit Function

Fail:
    Set yRange = Nothing
    Set rRange = Nothing
    CurveRanges = False

End Function


' Linear interpolation on (x=tenor years, y=rate) with flat extrapolation at ends.
Private Function LinInterp(xRange As Range, yRange As Range, x As Double) As Variant
    Dim xa() As Double, ya() As Double, n As Long, i As Long
    Dim cx As Range, cy As Range
    Dim xs As Variant, ys As Variant
    xs = xRange.value: ys = yRange.value


    ReDim xa(1 To UBound(xs, 1))
    ReDim ya(1 To UBound(ys, 1))
    n = 0
    For i = 1 To UBound(xs, 1)
        If Not IsError(xs(i, 1)) And _
            Not IsError(ys(i, 1)) Then
            If Len(CleanText(xs(i, 1))) > 0 And _
               Len(CleanText(ys(i, 1))) > 0 Then
                If IsNumeric(xs(i, 1)) And _
                   IsNumeric(ys(i, 1)) Then
                    n = n + 1
                    xa(n) = CDbl(xs(i, 1))
                    ya(n) = CDbl(ys(i, 1))
                End If
            End If
        End If
    Next i
    If n < 2 Then
        LinInterp = CVErr(xlErrNA)
        Exit Function
    End If


    ' simple insertion sort by x (nodes are few)
    Dim j As Long, tx As Double, ty As Double
    For i = 2 To n
        tx = xa(i): ty = ya(i): j = i - 1
        Do While j >= 1
            If xa(j) <= tx Then Exit Do
            xa(j + 1) = xa(j): ya(j + 1) = ya(j): j = j - 1
        Loop
        xa(j + 1) = tx: ya(j + 1) = ty
    Next i


    If x <= xa(1) Then LinInterp = ya(1): Exit Function
    If x >= xa(n) Then LinInterp = ya(n): Exit Function
    For i = 1 To n - 1
        If x >= xa(i) And x <= xa(i + 1) Then
            If xa(i + 1) = xa(i) Then
                LinInterp = ya(i)
            Else
                LinInterp = ya(i) + (ya(i + 1) - ya(i)) * (x - xa(i)) / (xa(i + 1) - xa(i))
            End If
            Exit Function
        End If
    Next i
    LinInterp = ya(n)
End Function




' =============================================================================
' OPICS CONNECTION
' =============================================================================


Private Function OpenOPICS(conn As Object, wsCfg As Worksheet) As Boolean
    On Error GoTo Fail
    Dim serverName As String, uid As String, pwd As String, connStr As String
    serverName = Replace(CleanText(wsCfg.Range(CFG_DSN).value), "/", "\")
    uid = CleanText(wsCfg.Range(CFG_UID).value)


    If serverName = "" Then
        MsgBox "Missing OPICS server/DSN in Config " & CFG_DSN & ".", vbExclamation
        OpenOPICS = False: Exit Function
    End If


    If uid = "" Then
        connStr = "Provider=MSOLEDBSQL;Data Source=" & serverName & _
                  ";Initial Catalog=OpicsMain;Integrated Security=SSPI;"
    Else
        pwd = CStr(wsCfg.Range(CFG_SQL_PWD).value)   ' optional SQL password cell
        connStr = "Provider=MSOLEDBSQL;Data Source=" & serverName & _
                  ";Initial Catalog=OpicsMain;User ID=" & uid & ";Password=" & pwd & ";"
    End If


    Dim t As Long: t = val(CleanText(wsCfg.Range(CFG_BBG_TIMEOUT).value))
    If t <= 0 Then t = 30
    conn.ConnectionTimeout = t
    conn.CommandTimeout = 120
    conn.Open connStr
    OpenOPICS = (conn.State = 1)
    Exit Function
Fail:
    MsgBox "OPICS connection failed:" & vbCrLf & Err.Description, vbCritical
    OpenOPICS = False
End Function


Private Function GetExcelBitness() As String
    #If Win64 Then
        GetExcelBitness = "64-bit"
    #Else
        GetExcelBitness = "32-bit"
    #End If
End Function




' =============================================================================
' RECORDSET / VALUE HELPERS
' =============================================================================


Private Function Fld(rs As Object, fieldName As String) As Variant
    On Error GoTo Miss
    Fld = rs.fields(fieldName).value
    Exit Function
Miss:
    Fld = Null
End Function


Private Function CDblSafe(v As Variant) As Double
    On Error GoTo Zero
    If IsNull(v) Or v = "" Then CDblSafe = 0 Else CDblSafe = CDbl(v)
    Exit Function
Zero:
    CDblSafe = 0
End Function


Private Function NullToStr(v As Variant) As String
    If IsNull(v) Then NullToStr = "" Else NullToStr = CStr(v)
End Function


Private Function NullToDate(v As Variant) As Variant
    On Error GoTo Blank
    If IsNull(v) Or v = "" Then NullToDate = "" Else NullToDate = CDate(v)
    Exit Function
Blank:
    NullToDate = ""
End Function


' Returns "" (blank) for NULL so the cell stays empty rather than 0.
Private Function NullToNumBlank(v As Variant) As Variant
    On Error GoTo Blank
    If IsNull(v) Or v = "" Then NullToNumBlank = "" Else NullToNumBlank = CDbl(v)
    Exit Function
Blank:
    NullToNumBlank = ""
End Function


Private Function CleanText(v As Variant) As String
    If IsNull(v) Then CleanText = "" Else CleanText = Trim$(CStr(v))
End Function


Private Function IdText(ByVal v As Variant) As String


    If IsError(v) Then
        IdText = ""
        Exit Function
    End If


    If IsNull(v) Then
        IdText = ""
        Exit Function
    End If


    If IsEmpty(v) Then
        IdText = ""
        Exit Function
    End If


    ' If the source is already text, preserve it exactly after trimming.
    If VarType(v) = vbString Then
        IdText = CleanText(v)
        Exit Function
    End If


    ' If Excel/ADO has already converted a long ID into a number, avoid scientific
    ' notation in the display. This cannot recover digits already lost by Excel,
    ' but it prevents further E+ formatting.
    If IsNumeric(v) Then
        IdText = Format$(CDbl(v), "0")
        Exit Function
    End If


    IdText = CleanText(v)


End Function


Private Function EscapeSQL(s As String) As String
    EscapeSQL = Replace(s, "'", "''")
End Function


Private Function TrimChar(s As String, ch As String) As String
    Dim r As String: r = s
    Do While Len(r) > 0 And Left$(r, 1) = ch
        r = Mid$(r, 2)
    Loop
    Do While Len(r) > 0 And Right$(r, 1) = ch
        r = Left$(r, Len(r) - 1)
    Loop
    TrimChar = r
End Function


' Builds a SQL IN-list fragment from a comma list, e.g. "H" -> "'H'" ; "A,T" -> "'A','T'"
Private Function BuildInClause(csv As String) As String
    Dim parts() As String, i As Long, out As String, t As String
    If CleanText(csv) = "" Then BuildInClause = "": Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        t = Trim$(parts(i))
        If t <> "" Then
            If out <> "" Then out = out & ","
            out = out & "'" & EscapeSQL(t) & "'"
        End If
    Next i
    BuildInClause = out
End Function


Private Function BuildLoadedBondISINList(wsBnd As Worksheet) As String
    Dim lastRow As Long: lastRow = LastBondDataRow(wsBnd)
    Dim r As Long, out As String, t As String
    For r = BOND_DATA_ROW To lastRow
        t = CleanText(wsBnd.Cells(r, colNum(BCOL_ISIN)).value)
        If t <> "" Then
            If out <> "" Then out = out & ","
            out = out & "'" & EscapeSQL(t) & "'"
        End If
    Next r
    BuildLoadedBondISINList = out
End Function


Public Function CouponFreqNum(ByVal v As Variant) As Variant
    Dim s As String
    s = UCase$(Trim$(CStr(v)))


    Select Case s
        Case "", "0"
            CouponFreqNum = CVErr(xlErrNA)


        Case "1", "A", "ANNUAL", "ANNUALLY", "12M", "Y", "YEARLY"
            CouponFreqNum = 1


        Case "2", "S", "SA", "SEMI", "SEMIANNUAL", "SEMI-ANNUAL", "6M"
            CouponFreqNum = 2


        Case "4", "Q", "QUARTERLY", "3M"
            CouponFreqNum = 4


        Case Else
            If IsNumeric(v) Then
                If CLng(v) = 1 Or CLng(v) = 2 Or CLng(v) = 4 Then
                    CouponFreqNum = CLng(v)
                Else
                    CouponFreqNum = CVErr(xlErrNA)
                End If
            Else
                CouponFreqNum = CVErr(xlErrNA)
            End If
    End Select
End Function






' =============================================================================
' LAST-ROW HELPERS
' =============================================================================


' Last row holding a bond.  Found from the sheet with End(xlUp), so a book that
' grows past any previous size is still seen in full.  The old version scanned a
' fixed 600 rows and silently ignored bond 601.
Private Function LastBondDataRow(ws As Worksheet) As Long
    LastBondDataRow = LastRowInColumn(ws, BCOL_ISIN, BOND_DATA_ROW)
End Function


' Last row holding a futures position.  Same reasoning as LastBondDataRow.
Private Function LastFutureDataRow(ws As Worksheet) As Long
    LastFutureDataRow = LastRowInColumn(ws, FCOL_CONTRACTCODE, DATA_ROW)
End Function


' Shared implementation: last non-empty row of `keyCol` at or below
' `firstDataRow`, or firstDataRow - 1 when the sheet holds no data.
Private Function LastRowInColumn( _
    ByVal ws As Worksheet, _
    ByVal keyCol As String, _
    ByVal firstDataRow As Long) As Long

    Dim r As Long

    On Error GoTo Empty_

    r = ws.Cells(ws.Rows.Count, colNum(keyCol)).End(xlUp).Row

    If r < firstDataRow Then GoTo Empty_

    LastRowInColumn = r
    Exit Function

Empty_:
    LastRowInColumn = firstDataRow - 1

End Function


' How far a clear or a scan must reach.
'
' Deliberately PAST the current data: if yesterday loaded 400 bonds and today
' loads 380, rows 381-400 still hold yesterday's formulas and must be cleared,
' or PNL_Attribution keeps reporting twenty positions the desk no longer has.
Private Function SheetClearLastRow( _
    ByVal ws As Worksheet, _
    ByVal firstDataRow As Long, _
    ByVal lastDataRow As Long) As Long

    Dim usedLast As Long

    On Error Resume Next
    usedLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0

    SheetClearLastRow = lastDataRow
    If usedLast > SheetClearLastRow Then SheetClearLastRow = usedLast
    If SheetClearLastRow < firstDataRow Then SheetClearLastRow = firstDataRow
    If SheetClearLastRow > firstDataRow + MAX_SHEET_ROWS - 1 Then
        SheetClearLastRow = firstDataRow + MAX_SHEET_ROWS - 1
    End If

End Function


Private Function LastSwapDataRow(ByVal ws As Worksheet) As Long
    Dim r As Long


    r = ws.Cells(ws.Rows.Count, WCOL_DEALID).End(xlUp).Row


    If r < DATA_ROW Then
        LastSwapDataRow = DATA_ROW - 1
    Else
        LastSwapDataRow = r
    End If
End Function


' =============================================================================
' BOND <-> HEDGE LINKING  (writes hedge id back onto the bond row)
'   Bond cols: BV=74 first swap link, BW=75 first future link (status/diagnostic)
' =============================================================================


Private Sub LinkSwapToBond(ByVal wsBnd As Worksheet, ByVal isin As String, ByVal dealId As String)


    Dim r As Long
    Dim cleanISIN As String
    Dim cleanDealId As String


    cleanISIN = UCase$(CleanText(isin))
    cleanDealId = IdText(dealId)


    If cleanISIN = "" Then Exit Sub
    If cleanDealId = "" Then Exit Sub


    For r = BOND_DATA_ROW To LastBondDataRow(wsBnd)


        If UCase$(CleanText(wsBnd.Cells(r, colNum(BCOL_ISIN)).value)) = cleanISIN Then


            ' BT = SwapLink.
            ' Force text format so long IDs are not converted to scientific notation.
            wsBnd.Cells(r, BCOL_SWAPLINK).numberFormat = "@"
            wsBnd.Cells(r, BCOL_SWAPLINK).value = AppendCsv(CStr(wsBnd.Cells(r, BCOL_SWAPLINK).value), cleanDealId)


        End If


    Next r


End Sub


Private Sub LinkFutureToBond(ByVal wsBnd As Worksheet, ByVal isin As String, ByVal contractCode As String)


    Dim r As Long
    Dim cleanISIN As String
    Dim cleanContractCode As String


    cleanISIN = UCase$(CleanText(isin))
    cleanContractCode = IdText(contractCode)


    If cleanISIN = "" Then Exit Sub
    If cleanContractCode = "" Then Exit Sub


    For r = BOND_DATA_ROW To LastBondDataRow(wsBnd)


        If UCase$(CleanText(wsBnd.Cells(r, colNum(BCOL_ISIN)).value)) = cleanISIN Then


            ' BU = FutLink.
            ' Force text format.
            wsBnd.Cells(r, BCOL_FUTLINK).numberFormat = "@"
            wsBnd.Cells(r, BCOL_FUTLINK).value = AppendCsv(CStr(wsBnd.Cells(r, BCOL_FUTLINK).value), cleanContractCode)


        End If


    Next r


End Sub


Private Function AppendCsv(ByVal existing As String, ByVal item As String) As String


    Dim e As String
    Dim x As String


    e = CleanText(existing)
    x = IdText(item)


    If x = "" Then
        AppendCsv = e
        Exit Function
    End If


    If e = "" Then
        AppendCsv = x
    ElseIf InStr(1, "," & e & ",", "," & x & ",", vbTextCompare) = 0 Then
        AppendCsv = e & "," & x
    Else
        AppendCsv = e
    End If


End Function




' =============================================================================
' BLOOMBERG REFRESH + WAIT
'   Triggers the BBG add-in recalc on the given ranges and waits for the
'   #N/A Requesting Data... cells to clear (or timeout).
' =============================================================================




Private Function RangeHasPendingBBG(ByVal rng As Range) As Boolean
    Dim c As Range, v As Variant
    For Each c In rng.Cells
        v = c.value
        If IsError(v) Then
            If InStr(1, c.Text, "Requesting", vbTextCompare) > 0 Then
                RangeHasPendingBBG = True
                Exit Function
            End If
        ElseIf VarType(v) = vbString Then
            If InStr(1, CStr(v), "Requesting Data", vbTextCompare) > 0 _
               Or InStr(1, CStr(v), "#N/A Requesting", vbTextCompare) > 0 Then
                RangeHasPendingBBG = True
                Exit Function
            End If
        End If
    Next c
    RangeHasPendingBBG = False
End Function


Private Sub WaitSeconds(ByVal s As Double)
    Dim t As Single: t = Timer
    Do While Timer - t < s
        DoEvents
    Loop
End Sub


' BDH writer used by T0 curve snapshot (single-point history -> "" on miss)
Private Sub WriteBDHDown(ws As Worksheet, tickerCol As String, outCol As String, _
                         firstRow As Long, lastRow As Long, fieldName As String)


    Dim n As Long
    Dim cfgT0 As String


    n = ws.Range(tickerCol & "1").Column
    cfgT0 = CfgR1C1(CFG_T0_DATE)


    ws.Range(outCol & firstRow & ":" & outCol & lastRow).FormulaR1C1 = _
        "=IF(RC" & n & "="""",""""," & _
            BDHLastPointExprR1C1( _
                "RC" & n, _
                XlText(fieldName), _
                "INT(" & cfgT0 & ")-3", _
                "INT(" & cfgT0 & ")" _
            ) & _
        ")"


End Sub




' =============================================================================
' DATE / CONFIG HELPERS
' =============================================================================


' GetAsOfDateString -> 'yyyy-mm-dd' from the T1 (as-of) config date.
Private Function GetAsOfDateString(wsCfg As Worksheet) As String
    Dim d As Variant: d = wsCfg.Range(CFG_T1_DATE).value
    If IsDate(d) Then
        GetAsOfDateString = Format$(CDate(d), "yyyy-mm-dd")
    Else
        GetAsOfDateString = Format$(Date, "yyyy-mm-dd")
    End If
End Function


' Sets T1 = last business day, T0 = prior business day, anchored to 5pm UTC cutoff.
Private Sub Update_Config_Dates_5pmUTC(wsCfg As Worksheet)
    ' Only auto-populate if the user has not pinned dates manually.
    If CleanText(wsCfg.Range(CFG_AUTO_DATE).value) = "AUTO" Then
        Dim t1 As Date, t0 As Date
        t1 = PrevBusinessDay(Date)
        t0 = PrevBusinessDay(t1)
        wsCfg.Range(CFG_T1_DATE).value = t1
        wsCfg.Range(CFG_T0_DATE).value = t0
    End If
End Sub


Private Function PrevBusinessDay(d As Date) As Date
    Dim x As Date: x = d - 1
    Do While Weekday(x, vbMonday) >= 6   ' Sat/Sun
        x = x - 1
    Loop
    PrevBusinessDay = x
End Function






Public Function SwapFloatFamily( _
    ByVal rateIndex As Variant, _
    Optional ByVal ccy As Variant = "") As String


    Dim s As String
    Dim cur As String


    s = UCase$(CleanText(rateIndex))
    cur = UCase$(CleanText(ccy))


    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")


    If Len(s) = 0 Then
        SwapFloatFamily = "UNKNOWN"
        Exit Function
    End If


    ' USD overnight-index swaps.
    If cur = "USD" Then
        If InStr(1, s, "SOFR", vbTextCompare) > 0 Or _
           InStr(1, s, "USSOFR", vbTextCompare) > 0 Or _
           InStr(1, s, "SOFRRATE", vbTextCompare) > 0 Or _
           InStr(1, s, "SOFRINDEX", vbTextCompare) > 0 Or _
           InStr(1, s, "OIS", vbTextCompare) > 0 Then


            SwapFloatFamily = "SOFR"
            Exit Function
        End If
    End If


    ' EUR overnight-index swaps.
    If cur = "EUR" Then
        If InStr(1, s, "ESTR", vbTextCompare) > 0 Or _
           InStr(1, s, "ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬STR", vbTextCompare) > 0 Or _
           InStr(1, s, "EONIA", vbTextCompare) > 0 Or _
           InStr(1, s, "OIS", vbTextCompare) > 0 Then


            SwapFloatFamily = "ESTR"
            Exit Function
        End If
    End If


    ' EURIBOR projection indices.
    If cur = "EUR" Then
        If InStr(1, s, "EURIBOR", vbTextCompare) > 0 Or _
           InStr(1, s, "EURIB", vbTextCompare) > 0 Or _
           InStr(1, s, "EUR003M", vbTextCompare) > 0 Or _
           InStr(1, s, "EUR006M", vbTextCompare) > 0 Or _
           InStr(1, s, "EUR012M", vbTextCompare) > 0 Then


            SwapFloatFamily = "EURIBOR"
            Exit Function
        End If
    End If


    SwapFloatFamily = "UNKNOWN"


End Function


Public Function SwapFloatFamilyStatus( _
    ByVal rateIndex As Variant, _
    ByVal ccy As Variant, _
    ByVal family As Variant) As String


    Dim s As String
    Dim cur As String
    Dim fam As String


    s = UCase$(CleanText(rateIndex))
    cur = UCase$(CleanText(ccy))
    fam = UCase$(CleanText(family))


    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")


    If Len(s) = 0 Then
        SwapFloatFamilyStatus = "UNKNOWN_SWAP_FAMILY"
        Exit Function
    End If


    If cur = "EUR" And _
       InStr(1, s, "SOFR", vbTextCompare) > 0 Then


        SwapFloatFamilyStatus = "SWAP_CCY_INDEX_MISMATCH"
        Exit Function
    End If


    If cur = "USD" And _
       (InStr(1, s, "EURIBOR", vbTextCompare) > 0 Or _
        InStr(1, s, "EURIB", vbTextCompare) > 0 Or _
        InStr(1, s, "EUR003M", vbTextCompare) > 0 Or _
        InStr(1, s, "EUR006M", vbTextCompare) > 0 Or _
        InStr(1, s, "EUR012M", vbTextCompare) > 0 Or _
        InStr(1, s, "ESTR", vbTextCompare) > 0 Or _
        InStr(1, s, "ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬STR", vbTextCompare) > 0 Or _
        InStr(1, s, "EONIA", vbTextCompare) > 0) Then


        SwapFloatFamilyStatus = "SWAP_CCY_INDEX_MISMATCH"
        Exit Function
    End If


    If fam = "UNKNOWN" Or Len(fam) = 0 Then
        SwapFloatFamilyStatus = "UNKNOWN_SWAP_FAMILY"
        Exit Function
    End If


    SwapFloatFamilyStatus = "OK"


End Function


Public Sub ValidateSwapFloatFamilies()


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dealId As String
    Dim ccy As String
    Dim family As String
    Dim statusText As String
    Dim issueCount As Long


    Set ws = ThisWorkbook.Worksheets(SH_SWAPS)
    lastRow = LastSwapDataRow(ws)


    If lastRow < DATA_ROW Then Exit Sub


    For r = DATA_ROW To lastRow


        dealId = CleanText( _
            ws.Cells(r, colNum(WCOL_DEALID)).value)


        If Len(dealId) > 0 Then


            ccy = UCase$(CleanText( _
                ws.Cells(r, colNum(WCOL_CCY)).value))


            family = UCase$(CleanText( _
                ws.Cells(r, _
                    colNum(WCOL_FLOATINDEX_FAMILY)).value))


            statusText = UCase$(CleanText( _
                ws.Cells(r, _
                    colNum(WCOL_FLOATFAMILY_STATUS)).value))


            If statusText <> "OK" Then


                issueCount = issueCount + 1


                Debug.Print _
                    "Swap row " & CStr(r) & _
                    "; DealID=" & dealId & _
                    "; CCY=" & ccy & _
                    "; Family=" & family & _
                    "; Status=" & statusText


            End If


        End If


    Next r


    MsgBox _
        CStr(issueCount) & _
        " swap family rows require review.", _
        IIf(issueCount = 0, _
            vbInformation, _
            vbExclamation), _
        "Swap Family Validation"


End Sub


Public Function SwapFloatTenor(ByVal rateIndex As Variant, Optional ByVal family As Variant = "") As String


    Dim s As String
    Dim fam As String


    s = UCase$(CleanText(rateIndex))
    fam = UCase$(CleanText(family))


    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")


    
    If fam = "ESTR" Or fam = "SOFR" Then
        SwapFloatTenor = "ON"
        Exit Function
    End If




    If InStr(1, s, "012M", vbTextCompare) > 0 Or _
       InStr(1, s, "12M", vbTextCompare) > 0 Or _
       InStr(1, s, "1Y", vbTextCompare) > 0 Then


        SwapFloatTenor = "12M"
        Exit Function


    End If


    If InStr(1, s, "006M", vbTextCompare) > 0 Or _
       InStr(1, s, "6M", vbTextCompare) > 0 Or _
       InStr(1, s, "6MO", vbTextCompare) > 0 Then


        SwapFloatTenor = "6M"
        Exit Function


    End If


    If InStr(1, s, "003M", vbTextCompare) > 0 Or _
       InStr(1, s, "3M", vbTextCompare) > 0 Or _
       InStr(1, s, "3MO", vbTextCompare) > 0 Then


        SwapFloatTenor = "3M"
        Exit Function


    End If


    SwapFloatTenor = "UNKNOWN"


End Function






' =============================================================================
' DIAGNOSTICS LOGGING
' =============================================================================


Private Sub SetupDiagnosticsSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DIAG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_DIAG
    End If
    If CleanText(ws.Range(CELL_DIAG_A1).value) = "" Then
        ws.Range(RNG_DIAG_A1_F1).value = Array("Timestamp", "Step", "Sheet", "Address", "ErrorType", "Sample")
        ws.Range(RNG_DIAG_A1_F1).Font.Bold = True
    End If
End Sub


' Column-signature dump for a sheet (diagnostic).  All args ByVal; first arg is a
' Workbook (never a string).  Safe call:  Call SigForSheet(ThisWorkbook, "Bonds", 4, 5)
Private Sub SigForSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal headerRow As Long, ByVal dataRow As Long)
    Dim ws As Worksheet, wsOut As Worksheet
    Dim sigHeaders As Variant
    Dim lastCol As Long, c As Long, outRow As Long


    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    Set wsOut = wb.Worksheets(SH_DIAG)
    On Error GoTo 0


    If ws Is Nothing Then Exit Sub
    If wsOut Is Nothing Then
        SetupDiagnosticsSheet
        Set wsOut = wb.Worksheets(SH_DIAG)
    End If
    If wsOut Is Nothing Then Exit Sub


    sigHeaders = Array("Col", "Header", "Address", "Formula", "Value", _
                       "NumberFormat", "Width", "Hidden")
    wsOut.Cells(headerRow, 1).Resize(1, UBound(sigHeaders) - LBound(sigHeaders) + 1).value = sigHeaders
    wsOut.Cells(headerRow, 1).Resize(1, UBound(sigHeaders) - LBound(sigHeaders) + 1).Font.Bold = True


    lastCol = ws.Cells(dataRow - 1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub


    outRow = headerRow + 1
    For c = 1 To lastCol
        wsOut.Cells(outRow, 1).value = c
        wsOut.Cells(outRow, 2).value = ws.Cells(dataRow - 1, c).value
        wsOut.Cells(outRow, 3).value = ws.Cells(dataRow, c).Address(False, False)
        wsOut.Cells(outRow, 4).value = ws.Cells(dataRow, c).formula
        wsOut.Cells(outRow, 5).value = ws.Cells(dataRow, c).Text
        wsOut.Cells(outRow, 6).value = ws.Cells(dataRow, c).numberFormat
        wsOut.Cells(outRow, 7).value = ws.Columns(c).ColumnWidth
        wsOut.Cells(outRow, 8).value = ws.Columns(c).Hidden
        outRow = outRow + 1
    Next c
End Sub


Private Sub LogDiagnosticsForRanges(ByVal stepName As String, ParamArray rngs() As Variant)
    If CleanText(ThisWorkbook.Worksheets(SH_CONFIG).Range(CFG_DIAGNOSTICS_ENABLED).value) = "FALSE" Then Exit Sub


    SetupDiagnosticsSheet


    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_DIAG)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2


    Dim i As Long, c As Range, rg As Range, v As Variant
    Dim etype As String, sampleText As String
    For i = LBound(rngs) To UBound(rngs)
        If TypeName(rngs(i)) = "Range" Then
            Set rg = rngs(i)
            For Each c In rg.Cells
                v = c.value
                etype = ""
                sampleText = ""
                If IsError(v) Then
                    etype = "ERROR"
                    sampleText = c.Text
                ElseIf VarType(v) = vbString Then
                    sampleText = CStr(v)
                    If InStr(1, sampleText, "#N/A", vbTextCompare) > 0 Or _
                       InStr(1, sampleText, "Invalid", vbTextCompare) > 0 Or _
                       InStr(1, sampleText, "Not Applicable", vbTextCompare) > 0 Or _
                       InStr(1, sampleText, "Authorization", vbTextCompare) > 0 Or _
                       InStr(1, sampleText, "Requesting Data", vbTextCompare) > 0 Then
                        etype = "BBG_NA"
                    End If
                End If
                If etype <> "" Then
                    ws.Cells(nextRow, 1).value = Now()
                    ws.Cells(nextRow, 2).value = stepName
                    ws.Cells(nextRow, 3).value = c.Worksheet.Name
                    ws.Cells(nextRow, 4).value = c.Address(False, False)
                    ws.Cells(nextRow, 5).value = etype
                    ws.Cells(nextRow, 6).value = Left$(sampleText, 60)
                    nextRow = nextRow + 1
                End If
            Next c
        End If
    Next i
End Sub




' =============================================================================
' SHEET LAYOUT SETUP  (idempotent header writers; data preserved)
' =============================================================================


' =============================================================================
' CONFIG SHEET LAYOUT
'
' Every label is written to the A-column cell IMMEDIATELY LEFT of the B-column
' cell the code actually reads.  This used to drift by ~10 rows (labels at
' A20/A37:A49 against constants at B18/B27:B39), so a user typing a value next
' to a label wrote a cell nothing ever read - e.g. "Keep formulas" was labelled
' at A48 while KeepFormulasMode() reads B39.
'
' If you move a CFG_* constant, move its label here in the same edit.
'
' Date labels: the legacy VBA names are inverted relative to the report labels
' (see the DATE CONVENTION WARNING at the top of this module).  B4 is the PRIOR
' snapshot and B5 is the CURRENT/as-of date, so they are labelled that way here.
' Filling them the other way round flips the sign of every delta on the sheet.
' =============================================================================
' BONDS SHEET - RESERVED COMMENT BAND
'
' Rows 1..BOND_COMMENT_ROWS on Bonds are reserved for the desk's own notes.
' Nothing this module writes may touch them: the header row is BOND_HEADER_ROW
' and data starts at BOND_DATA_ROW, both derived from BOND_COMMENT_ROWS.
'
' WHY THIS NEEDS CODE AND NOT JUST A CONSTANT
'
' Bonds!A:L is not written by VBA at all - it is the output of an Excel query
' against OPICS (see LoadOPICS_Bonds), and the query's table decides where its
' own header sits.  Changing BOND_HEADER_ROW alone would move every macro-owned
' column M:CN down while the query's ISIN/Name/Maturity stayed on row 1, and the
' two halves of every bond row would be two rows out of step - silently, because
' both halves still look populated.
'
' EnsureBondsCommentRows moves the query itself, by inserting whole rows above
' it.  Inserting rows is the one operation Excel propagates to a ListObject /
' QueryTable destination automatically, so the binding survives and the next
' refresh lands in the right place.  It is idempotent and it never deletes rows:
' if the header is already at or below BOND_HEADER_ROW it does nothing at all.
' =============================================================================

Private Sub EnsureBondsCommentRows(ByVal ws As Worksheet)

    Dim headerRow As Long
    Dim missingRows As Long

    On Error GoTo Fail

    headerRow = BondsQueryHeaderRow(ws)

    ' Nothing on the sheet yet: it will be built at the right rows anyway.
    If headerRow = 0 Then
        LabelBondsCommentRows ws
        Exit Sub
    End If

    If headerRow > BOND_HEADER_ROW Then
        ' Someone has deliberately pushed the table further down.  Leave it:
        ' more comment space than we reserve is their business, and shifting it
        ' back up would delete rows that may hold their notes.
        Exit Sub
    End If

    If headerRow = BOND_HEADER_ROW Then
        LabelBondsCommentRows ws
        Exit Sub
    End If

    missingRows = BOND_HEADER_ROW - headerRow

    ' Insert ABOVE the current header.  This carries the query table, its
    ' destination binding and every macro-owned column down together, which is
    ' exactly the invariant that matters: A:L and M:CN must stay on the same row
    ' as each other.
    ws.Rows(headerRow & ":" & CStr(headerRow + missingRows - 1)).Insert _
        Shift:=xlDown

    LabelBondsCommentRows ws

    Exit Sub

Fail:
    ' Never let a layout nicety abort a data load.  The caller checks the header
    ' row separately and reports if it is still wrong.
    Err.Clear

End Sub


' Where the bond table's header actually sits right now.
'
' Prefers the query table's own header row, because that is the authority.
' Falls back to searching column A for the ISIN header when the sheet is driven
' by a plain QueryTable rather than a ListObject.  Returns 0 for an empty sheet.
Private Function BondsQueryHeaderRow(ByVal ws As Worksheet) As Long

    Dim lo As ListObject
    Dim found As Range

    On Error GoTo Fail

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        BondsQueryHeaderRow = lo.Range.Row
        Exit Function
    End If

    Set found = ws.Columns(colNum(BCOL_ISIN)).Find( _
        What:="ISIN", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not found Is Nothing Then
        BondsQueryHeaderRow = found.Row
        Exit Function
    End If

    ' No table and no header: an empty or brand-new sheet.
    BondsQueryHeaderRow = 0
    Exit Function

Fail:
    BondsQueryHeaderRow = 0

End Function


' Mark the reserved band so nobody wonders what the blank rows are for.
' Only ever writes into column A, and only when that cell is empty, so a note
' already written there is never overwritten.
Private Sub LabelBondsCommentRows(ByVal ws As Worksheet)

    If BOND_COMMENT_ROWS < 1 Then Exit Sub

    On Error Resume Next

    If CleanText(ws.Cells(1, colNum(BCOL_ISIN)).value) = "" Then
        ws.Cells(1, colNum(BCOL_ISIN)).value = _
            "Comments (rows 1-" & CStr(BOND_COMMENT_ROWS) & _
            " are reserved; the bond table starts at row " & _
            CStr(BOND_HEADER_ROW) & ")"
        ws.Cells(1, colNum(BCOL_ISIN)).Font.Italic = True
    End If

    On Error GoTo 0

End Sub


' True when the bond table's header is where the module expects it.  Used by
' LoadOPICS_Bonds to refuse to interpret a sheet whose two halves are out of
' step, rather than loading a whole book of misaligned rows.
Private Function BondsLayoutIsSane(ByVal ws As Worksheet) As Boolean

    Dim headerRow As Long

    headerRow = BondsQueryHeaderRow(ws)

    If headerRow <> 0 And headerRow <> BOND_HEADER_ROW Then
        BondsLayoutIsSane = False
        Exit Function
    End If

    BondsLayoutIsSane = BondsQueryWidthIsSane(ws)

End Function


' The query must stop exactly at BONDS_QUERY_LAST_COL.
'
' Checked against the query table's own width where there is a ListObject, and
' otherwise against the header row: the first macro column must be blank before
' SetupBondsFinalHeaders writes it, and the last query column must not be.
' Returns True when there is no table yet - a sheet that has not been loaded is
' not misaligned, it is empty.
Private Function BondsQueryWidthIsSane(ByVal ws As Worksheet) As Boolean

    Dim lo As ListObject
    Dim expected As Long

    On Error GoTo Unknown_

    expected = colNum(BONDS_QUERY_LAST_COL)

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        BondsQueryWidthIsSane = _
            (lo.Range.Column + lo.Range.Columns.Count - 1 = expected)
        Exit Function
    End If

    ' No ListObject: fall back to the header text of the last query column.
    If CleanText(ws.Cells(BOND_HEADER_ROW, expected).value) = "" Then
        GoTo Unknown_
    End If

    BondsQueryWidthIsSane = True
    Exit Function

Unknown_:
    ' Cannot tell - do not block the load on a guess.
    BondsQueryWidthIsSane = True

End Function


' =============================================================================
Private Sub SetupExtendedConfig(ws As Worksheet)
    ws.Range(CELL_CFG_A1).value = "PNL Explainer - Config"
    SetIfBlank ws, "A4", "T-1 date (prior snapshot)":     SetIfBlank ws, CFG_T0_DATE, ""
    SetIfBlank ws, "A5", "T0 date (current / as-of)":     SetIfBlank ws, CFG_T1_DATE, ""
    SetIfBlank ws, "A6", "Date mode":         SetIfBlank ws, CFG_AUTO_DATE, "MANUAL"
    SetIfBlank ws, "A8", "HTC INVTYPE"
    SetIfBlank ws, "A9", "Sell INVTYPE":      SetIfBlank ws, CFG_SELL_PORTS, "A"
    SetIfBlank ws, "A13", "OPICS server/DSN": SetIfBlank ws, CFG_DSN, ""
    SetIfBlank ws, "A14", "OPICS UID":        SetIfBlank ws, CFG_UID, ""
    SetIfBlank ws, "A15", "ISIN SECIDTYPE":   SetIfBlank ws, CFG_ISIN_TYPE, "4"
    SetIfBlank ws, "A18", "Spread framework FALLBACK (blank = auto; per-bond overrides go on SpreadOverride)": SetIfBlank ws, CFG_SPREAD_FRAMEWORK, ""
    SetIfBlank ws, "A19", "Extra FX-hedge label tokens, semicolon separated (blank = built-in list)": SetIfBlank ws, CFG_FX_HEDGE_TOKENS, ""
    SetIfBlank ws, "A27", "Base CCY":         SetIfBlank ws, CFG_BASE_CCY, "EUR"
    SetIfBlank ws, "A28", "FX fix source":    SetIfBlank ws, CFG_FX_FIX_SOURCE, "BGN"
    SetIfBlank ws, "A29", "Fut px field":     SetIfBlank ws, CFG_FUT_PRICE_FIELD, "PX_LAST"
    SetIfBlank ws, "A30", "BBG stall timeout (s)": SetIfBlank ws, CFG_BBG_TIMEOUT, "120"
    SetIfBlank ws, "A31", "YearsLeft basis":  SetIfBlank ws, CFG_YEARSLEFT_BASIS, "T1"
    SetIfBlank ws, "A32", "Curve source":     SetIfBlank ws, CFG_CURVE_SOURCE, "BBG"
    SetIfBlank ws, "A33", "Funding source":   SetIfBlank ws, CFG_FUNDING_SOURCE, "BBG"
    SetIfBlank ws, "A34", "Actual PnL source (label only - does not gate the calculation)": SetIfBlank ws, CFG_ACTUAL_PNL_SOURCE, "PROXY"
    SetIfBlank ws, "A35", "Diagnostics":      SetIfBlank ws, CFG_DIAGNOSTICS_ENABLED, "TRUE"
    SetIfBlank ws, "A36", "Convexity bump bp": SetIfBlank ws, CFG_CONVEXITY_BUMP_BP, 100
    SetIfBlank ws, "A39", "Keep formulas (offline / no BBG)": SetIfBlank ws, CFG_KEEP_FORMULAS, "FALSE"
    SetIfBlank ws, "A49", "T0 snapshot frozen at (auto - do not edit)"
    SetIfBlank ws, "A50", "T-1 snapshot cut-off used (auto - do not edit)"

    ' B41:B45 belong to modAccess, so modAccess labels them.  Repeating them
    ' here would be two places to edit and one of them would go stale.
    Access_SetupConfigCells
End Sub




Private Sub SetIfBlank(ws As Worksheet, addr As String, val As Variant)
    If CleanText(ws.Range(addr).value) = "" Then ws.Range(addr).value = val
End Sub


Private Sub WriteBondConvexityBumpFormulas(ByVal ws As Worksheet, ByVal lastRow As Long)


    If lastRow < BOND_DATA_ROW Then Exit Sub


    Dim f As Long
    f = BOND_DATA_ROW


    Dim cfgT1 As String
    Dim cfgBump As String


    cfgT1 = CfgR1C1(CFG_T1_DATE)
    cfgBump = CfgR1C1(CFG_CONVEXITY_BUMP_BP)


    ' -------------------------------------------------------------------------
    ' CF = Convexity_Bump_bp
    '
    ' CF = column 84
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_CONVEXITY_BUMP_BP & f & ":" & BCOL_CONVEXITY_BUMP_BP & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), cfgBump)


    ' -------------------------------------------------------------------------
    ' CG = CouponFreq_Num
    '
    ' CG = column 85
    ' Coupon frequency source is Bonds!E = RC5
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_COUPONFREQ_NUM & f & ":" & BCOL_COUPONFREQ_NUM & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "CouponFreqNum(" & RC(BCOL_COUPON_FREQ) & ")")


    ' -------------------------------------------------------------------------
    ' CH = Price_Base
    '
    ' CH = column 86
    '
    ' Inputs:
    '   Settlement date = Config current/T0 date = legacy CFG_T1_DATE
    '   Maturity        = RC6
    '   Coupon          = RC4 / 100
    '   Yield           = RC23, normalized if Bloomberg returns percent
    '   Redemption      = 100
    '   Frequency       = RC85 = CouponFreq_Num
    '   Basis           = ExcelPriceBasisFromBondDCC(RC91)
    '
    ' Important:
    '   RC91 = CM = BondDCC_Code
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_PRICE_BASE & f & ":" & BCOL_PRICE_BASE & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
        "IFERROR(PRICE(" & _
            "INT(" & cfgT1 & ")," & _
            RC(BCOL_MATURITY) & "," & _
            RC(BCOL_COUPON) & "/100," & _
            "IF(ABS(" & RC(BCOL_YTM_T0) & ")>1," & RC(BCOL_YTM_T0) & "/100," & RC(BCOL_YTM_T0) & ")," & _
            "100," & _
            RC(BCOL_COUPONFREQ_NUM) & "," & _
            "ExcelPriceBasisFromBondDCC(" & RC(BCOL_BONDDCC_CODE) & ")" & _
        "),""""))"


    ' -------------------------------------------------------------------------
    ' CI = Price_Up
    '
    ' CI = column 87
    ' Yield up by bump.
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_PRICE_UP & f & ":" & BCOL_PRICE_UP & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
        "IFERROR(PRICE(" & _
            "INT(" & cfgT1 & ")," & _
            RC(BCOL_MATURITY) & "," & _
            RC(BCOL_COUPON) & "/100," & _
            "IF(ABS(" & RC(BCOL_YTM_T0) & ")>1," & RC(BCOL_YTM_T0) & "/100," & RC(BCOL_YTM_T0) & ")+" & RC(BCOL_CONVEXITY_BUMP_BP) & "/10000," & _
            "100," & _
            RC(BCOL_COUPONFREQ_NUM) & "," & _
            "ExcelPriceBasisFromBondDCC(" & RC(BCOL_BONDDCC_CODE) & ")" & _
        "),""""))"


    ' -------------------------------------------------------------------------
    ' CJ = Price_Down
    '
    ' CJ = column 88
    ' Yield down by bump.
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_PRICE_DOWN & f & ":" & BCOL_PRICE_DOWN & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
        "IFERROR(PRICE(" & _
            "INT(" & cfgT1 & ")," & _
            RC(BCOL_MATURITY) & "," & _
            RC(BCOL_COUPON) & "/100," & _
            "IF(ABS(" & RC(BCOL_YTM_T0) & ")>1," & RC(BCOL_YTM_T0) & "/100," & RC(BCOL_YTM_T0) & ")-" & RC(BCOL_CONVEXITY_BUMP_BP) & "/10000," & _
            "100," & _
            RC(BCOL_COUPONFREQ_NUM) & "," & _
            "ExcelPriceBasisFromBondDCC(" & RC(BCOL_BONDDCC_CODE) & ")" & _
        "),""""))"


    ' -------------------------------------------------------------------------
    ' CK = Convexity_Bump
    '
    ' CK = column 89
    '
    ' Formula:
    '   Convexity =
    '       (P_down + P_up - 2 * P_base)
    '       / (P_base * bump_decimal ^ 2)
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_CONVEXITY_BUMP & f & ":" & BCOL_CONVEXITY_BUMP & lastRow).FormulaR1C1 = _
        "=IF(" & RC(BCOL_ISIN) & "="""",""""," & _
        "IF(AND(" & _
            "ISNUMBER(" & RC(BCOL_PRICE_BASE) & ")," & _
            "ISNUMBER(" & RC(BCOL_PRICE_UP) & ")," & _
            "ISNUMBER(" & RC(BCOL_PRICE_DOWN) & ")," & _
            "ISNUMBER(" & RC(BCOL_CONVEXITY_BUMP_BP) & ")," & _
            RC(BCOL_PRICE_BASE) & "<>0," & _
            RC(BCOL_CONVEXITY_BUMP_BP) & ">0" & _
        ")," & _
        "(" & RC(BCOL_PRICE_DOWN) & "+" & RC(BCOL_PRICE_UP) & "-2*" & RC(BCOL_PRICE_BASE) & ")/(" & RC(BCOL_PRICE_BASE) & "*(" & RC(BCOL_CONVEXITY_BUMP_BP) & "/10000)^2)," & _
        """""))"


    ' -------------------------------------------------------------------------
    ' Y = Convexity used by model
    '
    ' Y = column 25
    ' CK = Convexity_Bump = RC89
    '
    ' Y should be owned by the bump-reprice method, not Bloomberg CONVEXITY.
    ' -------------------------------------------------------------------------
    ws.Range(BCOL_CONVEXITY & f & ":" & BCOL_CONVEXITY & lastRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IF(ISNUMBER(" & RC(BCOL_CONVEXITY_BUMP) & ")," & RC(BCOL_CONVEXITY_BUMP) & ","""")")


End Sub


' Curve sheet: tenor labels/years and ticker columns.  Tickers themselves should
' live in a CurveMap (left of each block); only header scaffolding is set here.


Private Sub SetupCurvesLayout(ws As Worksheet)


    ws.Range(CELL_CV_A1).value = "OIS / Gov / Swap Curves"


    ws.Range(CELL_CV_A4).value = "Tenor"
    ws.Range(CELL_CV_B4).value = "Years"


    ' -------------------------------------------------------------------------
    ' EUR curve block
    '
    ' C:D:E = EUR ESTR/OIS curve
    ' I:J:K = EUR EURIBOR swap curve
    ' -------------------------------------------------------------------------
    
    ws.Range(CVCOL_EUR_ESTR_OIS_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_EUR_ESTR_OIS_T0 & CURVE_HEADER_ROW).value = Array("EUR_ESTR_OIS_Ticker", "EUR_ESTR_OIS_T-1", "EUR_ESTR_OIS_T0")
    ws.Range(CVCOL_EUR_GOV_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_EUR_GOV_T0 & CURVE_HEADER_ROW).value = Array("EUR_Gov_Ticker", "EUR_Gov_T-1", "EUR_Gov_T0")
    ws.Range(CVCOL_EUR_EURIBOR_SWAP_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_EUR_EURIBOR_SWAP_T0 & CURVE_HEADER_ROW).value = Array("EUR_EURIBOR_Swap_Ticker", "EUR_EURIBOR_Swap_T-1", "EUR_EURIBOR_Swap_T0")




    ' -------------------------------------------------------------------------
    ' USD curve block
    ' -------------------------------------------------------------------------
    ws.Range(CELL_CV_S4).value = "Years"
    
    ws.Range(CVCOL_USD_OIS_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_USD_OIS_T0 & CURVE_HEADER_ROW).value = Array("USD_OIS_Ticker", "USD_OIS_T-1", "USD_OIS_T0")
    ws.Range(CVCOL_USD_GOV_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_USD_GOV_T0 & CURVE_HEADER_ROW).value = Array("USD_Gov_Ticker", "USD_Gov_T-1", "USD_Gov_T0")
    ws.Range(CVCOL_USD_SWAP_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_USD_SWAP_T0 & CURVE_HEADER_ROW).value = Array("USD_Swap_Ticker", "USD_Swap_T-1", "USD_Swap_T0")




    ' -------------------------------------------------------------------------
    ' GBP curve block
    ' -------------------------------------------------------------------------
    ws.Range(CELL_CV_AJ4).value = "Years"
    
    ws.Range(CVCOL_GBP_OIS_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_GBP_OIS_T0 & CURVE_HEADER_ROW).value = Array("GBP_OIS_Ticker", "GBP_OIS_T-1", "GBP_OIS_T0")
    ws.Range(CVCOL_GBP_GOV_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_GBP_GOV_T0 & CURVE_HEADER_ROW).value = Array("GBP_Gov_Ticker", "GBP_Gov_T-1", "GBP_Gov_T0")
    ws.Range(CVCOL_GBP_SWAP_TICKER & CURVE_HEADER_ROW & ":" & CVCOL_GBP_SWAP_T0 & CURVE_HEADER_ROW).value = Array("GBP_Swap_Ticker", "GBP_Swap_T-1", "GBP_Swap_T0")




    Dim yrs As Variant
    yrs = Array(0.0833, 0.25, 0.5, 1, 2, 3, 5, 7, 10, 15, 20, 30, 40)


    Dim estrTickers As Variant
    estrTickers = Array( _
        "ESTRON Index", _
        "EESWEA Curncy", _
        "EESWEC Curncy", _
        "EESWEF Curncy", _
        "EESWE1 Curncy", _
        "EESWE2 Curncy", _
        "EESWE3 Curncy", _
        "EESWE5 Curncy", _
        "EESWE7 Curncy", _
        "EESWE10 Curncy", _
        "EESWE15 Curncy", _
        "EESWE20 Curncy", _
        "EESWE30 Curncy")


    Dim i As Long


    For i = 0 To UBound(yrs)


        ' Existing year scaffolding.
        SetIfBlank ws, "B" & (CURVE_FIRST_ROW + i), yrs(i)
        SetIfBlank ws, "S" & (CURVE_FIRST_ROW + i), yrs(i)
        SetIfBlank ws, "AJ" & (CURVE_FIRST_ROW + i), yrs(i)


        ' Seed EUR ESTR/OIS tickers into existing EUR OIS block.
        ' This only fills blanks, so manual edits survive.
        SetIfBlank ws, "C" & (CURVE_FIRST_ROW + i), estrTickers(i)


    Next i


    ' EURIBOR swap curve tickers should be maintained in:
    '   I7:I19
    '
    ' Recommended:
    '   I7  = ESTRON Index
    '   I8  = EUR001M Index
    '   I9  = EUR003M Index
    '   I10 = EUR006M Index
    '   I11 = EUSA1 Curncy
    '   I12 = EUSA2 Curncy
    '   I13 = EUSA3 Curncy
    '   I14 = EUSA5 Curncy
    '   I15 = EUSA7 Curncy
    '   I16 = EUSA10 Curncy
    '   I17 = EUSA15 Curncy
    '   I18 = EUSA20 Curncy
    '   I19 = EUSA30 Curncy


End Sub


Private Sub SetupBondsFinalHeaders(ws As Worksheet)


    ' A:L are query-owned. Do not overwrite them.
    '
    ' De-formula the table's header cell only (it has been seen holding a stale
    ' formula from an earlier layout).  Deliberately addressed as
    ' CELL_BONDS_HEADER & BOND_HEADER_ROW rather than a literal "A1": A1 is now
    ' the desk's comment band.
    ws.Range(CELL_BONDS_HEADER & BOND_HEADER_ROW).value = _
        ws.Range(CELL_BONDS_HEADER & BOND_HEADER_ROW).value


    ws.Range(BCOL_DAYSLEFT & BOND_HEADER_ROW & ":" & BCOL_YTM_T0 & BOND_HEADER_ROW).value = Array( _
        "DaysLeft", "OIS_T0", "OIS_T-1", "DF_T0", "DF_T-1", _
        "FX_T0", "FX_T-1", "CleanPx_T0", "DirtyPx_T0", _
        "AccruedInterest_T0", "YTM_T0")


    ws.Range(BCOL_MODDUR_T0 & BOND_HEADER_ROW & ":" & BCOL_DIRTYMV_T0_EUR & BOND_HEADER_ROW).value = Array( _
        "ModDur_T0", "Convexity", "ZSprd_T0", "ASW_T0", _
        "CleanPx_T-1", "DirtyPx_T-1", "YTM_T-1", "ZSprd_T-1", _
        "ASW_T-1", "DV01_EUR", "DirtyMV_T0_EUR")


    ws.Range(BCOL_DIRTYMV_TM1_EUR & BOND_HEADER_ROW & ":" & BCOL_OAS_TM1 & BOND_HEADER_ROW).value = Array( _
        "DirtyMV_T-1_EUR", "BookVal_EUR2", "ISpread_T0", "ISpread_T-1", _
        "GSpread_T0", "GSpread_T-1", "ISpread_Status", "BBG_Ticker", _
        "Ticker_Status", "OAS_T0", "OAS_T-1")


    ws.Range(BCOL_DELTAOAS & BOND_HEADER_ROW & ":" & BCOL_GOV_TM1 & BOND_HEADER_ROW).value = Array( _
        "DeltaOAS", "DV01_Unit", "SpreadDuration_Used", _
        "OAS_ModDuration_Raw", "OAS_Convexity", "Pricing_Source", _
        "Last_Pricing_Date", "Benchmark_Bond", "Benchmark_Name", _
        "Gov_T0", "Gov_T-1")


    ' PositionSide3 used to close this block.  It had a header and nothing else
    ' - no writer, no reader - so it is gone with the other PositionSide copies.
    ws.Range(BCOL_SWAP_T0 & BOND_HEADER_ROW & ":" & BCOL_DELTA_Y_BP & BOND_HEADER_ROW).value = Array( _
        "Swap_T0", "Swap_T-1", "g_T0", "g_T-1", "Delta_g_bp", _
        "q_T0", "q_T-1", "Delta_q_bp", "Delta_i_bp", _
        "Delta_y_bp")


    ws.Range(BCOL_FUNDTKR & BOND_HEADER_ROW & ":" & BCOL_BBG_CAND_BGN_CORP & BOND_HEADER_ROW).value = Array( _
        "FundTkr", "FundRate_T0", "FundRate_T-1", "Bond_Status", _
        "SwapLink", "FutLink", "BBG_Cand_ISIN", "BBG_Cand_slashISIN", _
        "BBG_Cand_Corp", "BBG_Cand_BVAL_Corp", "BBG_Cand_BGN_Corp")


    ws.Range(BCOL_BBG_CAND_GOVT & BOND_HEADER_ROW & ":" & BCOL_CONVEXITY_BUMP & BOND_HEADER_ROW).value = Array( _
        "BBG_Cand_Govt", "BBG_Cand_BVAL_Govt", "BBG_Cand_BGN_Govt", _
        "BBG_Cand_Mtge", "BBG_Cand_MMkt", _
        "Convexity_Bump_bp", "CouponFreq_Num", "Price_Base", _
        "Price_Up", "Price_Down", "Convexity_Bump")


    ws.Range(BCOL_DAY_CNT_DES & BOND_HEADER_ROW & ":" & BCOL_DV01_OPENING_EUR & BOND_HEADER_ROW).value = Array( _
        "DAY_CNT_DES", "BondDCC_Code", "BondDCC_Name", "DV01_Opening_EUR")


    ws.Range(BCOL_DAYSLEFT & BOND_HEADER_ROW & ":" & BCOL_DV01_OPENING_EUR & BOND_HEADER_ROW).Font.Bold = True
    ws.Range(BCOL_DAYSLEFT & BOND_HEADER_ROW & ":" & BCOL_DV01_OPENING_EUR & BOND_HEADER_ROW).HorizontalAlignment = xlCenter


    ws.Columns(BCOL_BBG_CAND_ISIN & ":" & BCOL_BBG_CAND_MMKT).Hidden = True
    ws.Columns(BCOL_DAY_CNT_DES & ":" & BCOL_DV01_OPENING_EUR).Hidden = False


End Sub


Private Sub SetupFuturesFinalHeaders(ws As Worksheet)


    ws.Range(CELL_FUT_A1).value = "Futures"


    ws.Range(FCOL_CONTRACTCODE & FUT_HEADER_ROW & ":" & FCOL_AVGENTRYPX & FUT_HEADER_ROW).value = Array( _
        "ContractCode", _
        "Exchange", _
        "CCY", _
        "Contracts", _
        "FaceValue", _
        "DelivDate", _
        "CTD_ISIN", _
        "CTD_CF", _
        "Portfolio", _
        "LinkedISIN", _
        "AvgEntryPx")


    ws.Range(FCOL_DAYSTODELIV & FUT_HEADER_ROW & ":" & FCOL_FUTURESPNL_EUR & FUT_HEADER_ROW).value = Array( _
        "DaysToDeliv", _
        "FX_T0", _
        "FutPx_T0", _
        "CTD_DirtyPx_T0", _
        "CTD_Ticker_T-1", _
        "FutPx_T-1", _
        "ImpliedRepo_Calc", _
        "GrossBasis_Calc", _
        "NotionalValue_EUR", _
        "FuturesPnL_EUR")


    ws.Range(FCOL_BBG_TICKER & FUT_HEADER_ROW & ":" & FCOL_STATUS_TM1 & FUT_HEADER_ROW).value = Array( _
        "BBG_Ticker", _
        "Status_T0", _
        "Status_T-1")


    ws.Range(FCOL_FUT_VAL_PT & FUT_HEADER_ROW & ":" & FCOL_STATUS & FUT_HEADER_ROW).value = Array( _
        "FUT_VAL_PT", _
        "CTD_Ticker", _
        "ConvFactor", _
        "HedgeUnitDV01", _
        "Futures_DV01_EUR", _
        "ImpliedRepo_BBG", _
        "NetBasis_BBG", _
        "GrossBasis_BBG", _
        "Status")


    ws.Range(FCOL_COVERAGE_SOURCEROW & FUT_HEADER_ROW & ":" & FCOL_HEDGE_CLASS & FUT_HEADER_ROW).value = Array( _
        "Coverage_SourceRow", _
        "CoverageRelation", _
        "FutureLabel", _
        "HedgeType", _
        "Coverage_StartDate", _
        "CoverageInfo_D", _
        "Hedge_Source", _
        "Coverage_BPV", _
        "Hedge_Class")


    ws.Range(FCOL_CONTRACTCODE & FUT_HEADER_ROW & ":" & FCOL_HEDGE_CLASS & FUT_HEADER_ROW).Font.Bold = True
    ws.Range(FCOL_CONTRACTCODE & FUT_HEADER_ROW & ":" & FCOL_HEDGE_CLASS & FUT_HEADER_ROW).HorizontalAlignment = xlCenter


End Sub


Private Sub SetupSwapsFinalHeaders(ws As Worksheet)


    ws.Range(CELL_SWAP_A1).value = "Swaps"


    ' -------------------------------------------------------------------------
    ' A:M = front-table economics and explicit float-curve identifier
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_DEALID & SWAP_HEADER_ROW & ":" & WCOL_CPTY & SWAP_HEADER_ROW).value = Array( _
        "DealID", _
        "Ccy", _
        "Notional", _
        "FixedRate", _
        "FloatIndex", _
        "FloatSpread", _
        "StartDate", _
        "EndDate", _
        "PayFixed", _
        "Portfolio", _
        "FloatCurve_Type", _
        "LinkedISIN", _
        "Cpty")


    ' -------------------------------------------------------------------------
    ' N:AC = model fields, model PnL, final PnL/status
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_YEARFRAC & SWAP_HEADER_ROW & ":" & WCOL_AF_STATUS & SWAP_HEADER_ROW).value = Array( _
        "YearFrac", _
        "OIS_T0", _
        "OIS_T-1", _
        "DF_T0", _
        "DF_T-1", _
        "Annuity_T0", _
        "Annuity_T-1", _
        "FX", _
        "PV_T0_Model", _
        "PV_T-1_Model", _
        "Swap_DV01_EUR", _
        "Model_PnL", _
        "Status", _
        "PnL_Source", _
        "PnL", _
        "AF_Status")


    ' -------------------------------------------------------------------------
    ' AD:AK = Bloomberg IDs and mapping source
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_BBG_SWAP_DIRECT_ID & SWAP_HEADER_ROW & ":" & WCOL_SWAPMAP_STATUS & SWAP_HEADER_ROW).value = Array( _
        "BBG_Swap_Direct_ID", _
        "BBG_Fixed_Leg_ID", _
        "BBG_Float_Leg_ID", _
        "Swap_ID_Source", _
        "CoverageRelation", _
        "SwapMap_SourceRow", _
        "SwapMap_Class", _
        "SwapMap_Status")


    ' -------------------------------------------------------------------------
    ' AL:AP = mapped economics
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_MAP_NOTIONAL & SWAP_HEADER_ROW & ":" & WCOL_NOTIONAL_SOURCE & SWAP_HEADER_ROW).value = Array( _
        "Map_Notional", _
        "Map_CCY", _
        "Map_Counterparty", _
        "Notional_Final", _
        "Notional_Source")


    ' -------------------------------------------------------------------------
    ' AQ:AT = BQL T1 NPV block
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_BQL_NPV_DIRECT_T0 & SWAP_HEADER_ROW & ":" & WCOL_BQL_NPV_TOTAL_T0 & SWAP_HEADER_ROW).value = Array( _
        "BQL_NPV_Direct_T0", _
        "BQL_NPV_Fixed_T0", _
        "BQL_NPV_Float_T0", _
        "BQL_NPV_Total_T0")


    ' -------------------------------------------------------------------------
    ' AU:AX = BQL T0 NPV block
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_BQL_NPV_DIRECT_TM1 & SWAP_HEADER_ROW & ":" & WCOL_BQL_NPV_TOTAL_TM1 & SWAP_HEADER_ROW).value = Array( _
        "BQL_NPV_Direct_T-1", _
        "BQL_NPV_Fixed_T-1", _
        "BQL_NPV_Float_T-1", _
        "BQL_NPV_Total_T-1")


    ' -------------------------------------------------------------------------
    ' AY:AZ = BQL PnL/status
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_BQL_SWAP_PNL & SWAP_HEADER_ROW & ":" & WCOL_BQL_SWAP_STATUS & SWAP_HEADER_ROW).value = Array( _
        "BQL_Swap_PnL", _
        "BQL_Swap_Status")


    ' -------------------------------------------------------------------------
    ' BA:BI = float-index / curve-selection diagnostics
    ' -------------------------------------------------------------------------
    ws.Range(WCOL_PAY_FLT_RATE_IDX & SWAP_HEADER_ROW & ":" & WCOL_DELTA_MODELSPREAD_BP & SWAP_HEADER_ROW).value = Array( _
        "PAY_FLT_RATE_IDX", _
        "FloatIndex_Family", _
        "FloatIndex_Tenor", _
        "FloatCurve_T0", _
        "FloatCurve_T-1", _
        "ModelSpread_T0", _
        "ModelSpread_T-1", _
        "Delta_FloatCurve_bp", _
        "Delta_ModelSpread_bp")


    ' -------------------------------------------------------------------------
    ' BJ = BQL_Swap_DV01 (Bloomberg swap DV01); BK:BM = true reserved tail
    ' -------------------------------------------------------------------------
    
    ws.Range(WCOL_RESERVED_1 & SWAP_HEADER_ROW & ":" & WCOL_RESERVED_4 & SWAP_HEADER_ROW).value = Array( _
        "Reserved_1", _
        "Reserved_2", _
        "Reserved_3", _
        "Reserved_4")


    ws.Range(WCOL_DV01_BBG & SWAP_HEADER_ROW).value = "DV01_BBG"
    
    ws.Range(WCOL_FLOATFAMILY_STATUS & SWAP_HEADER_ROW).value = "FloatFamily_Status"




    ws.Range(WCOL_DEALID & SWAP_HEADER_ROW & ":" & _
         WCOL_FLOATFAMILY_STATUS & SWAP_HEADER_ROW).Font.Bold = True


    ws.Range(WCOL_DEALID & SWAP_HEADER_ROW & ":" & _
         WCOL_FLOATFAMILY_STATUS & SWAP_HEADER_ROW).HorizontalAlignment = xlCenter


End Sub






Public Sub Debug_Dump_T0_BDH_Formulas()


    Dim wsDiag As Worksheet
    Dim nextRow As Long


    SetupDiagnosticsSheet
    Set wsDiag = ThisWorkbook.Worksheets(SH_DIAG)


    wsDiag.Cells.ClearContents
    wsDiag.Range(RNG_DIAG_A1_E1).value = Array("Timestamp", "Sheet", "Cell", "Displayed_Text", "FormulaR1C1")
    wsDiag.Range(RNG_DIAG_A1_E1).Font.Bold = True


    nextRow = 2


    DumpFormulaRange ThisWorkbook.Worksheets(SH_OIS).Range(RNG_CV_A51_A112__D51_D112__G51_G112__R51_R112__U51_U112__Y51_Y112__AI51_AI112__AL51_AL112__AO51_AO112), wsDiag, nextRow


    DumpFormulaRange ThisWorkbook.Worksheets(SH_BONDS).Range(BCOL_DIRTYPX_T0 & DATA_ROW & ":" & BCOL_DIRTYPX_T0 & CStr(LastBondDataRow(ThisWorkbook.Worksheets(SH_BONDS))) & "," & BCOL_YTM_TM1 & DATA_ROW & ":" & BCOL_DIRTYMV_T0_EUR & CStr(LastBondDataRow(ThisWorkbook.Worksheets(SH_BONDS))) & "," & BCOL_DV01_UNIT & DATA_ROW & ":" & BCOL_DV01_UNIT & CStr(LastBondDataRow(ThisWorkbook.Worksheets(SH_BONDS))) & "," & BCOL_SWAPLINK & DATA_ROW & ":" & BCOL_SWAPLINK & CStr(LastBondDataRow(ThisWorkbook.Worksheets(SH_BONDS)))), _
        wsDiag, nextRow


    DumpFormulaRange ThisWorkbook.Worksheets(SH_FUTURES).Range(FCOL_FUTPX_TM1 & DATA_ROW & ":" & FCOL_FUTPX_TM1 & CStr(LastFutureDataRow(ThisWorkbook.Worksheets(SH_FUTURES)))), _
        wsDiag, nextRow


    MsgBox "BDH formulas dumped to Diagnostics.", vbInformation


End Sub




Private Sub DumpFormulaRange(ByVal rng As Range, ByVal wsDiag As Worksheet, ByRef nextRow As Long)


    Dim c As Range
    Dim a As Range


    For Each a In rng.Areas
        For Each c In a.Cells


            If c.HasFormula Then
                wsDiag.Cells(nextRow, 1).value = Now()
                wsDiag.Cells(nextRow, 2).value = c.Worksheet.Name
                wsDiag.Cells(nextRow, 3).value = c.Address(False, False)
                wsDiag.Cells(nextRow, 4).value = c.Text
                wsDiag.Cells(nextRow, 5).value = c.FormulaR1C1
                nextRow = nextRow + 1
            End If


        Next c
    Next a


End Sub


' Adding targeted range calculation helper


Private Sub CalculateRangesMany(ByVal rngs As Variant)
    Dim i As Long
    Dim rg As Range


    For i = LBound(rngs) To UBound(rngs)
        If TypeName(rngs(i)) = "Range" Then
            Set rg = rngs(i)
            rg.Calculate
        End If
    Next i
End Sub




Public Sub Refresh_CoverageSupport_And_BuildSwapMap()
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldAskLinks As Boolean


    Dim importedSwapRows As Long
    Dim importedFutureRows As Long


    Dim wsCfg As Worksheet


    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldAskLinks = Application.AskToUpdateLinks


    On Error GoTo CleanFail


    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing Coverage support from Hedge_Risco..."


    
    ClearCoverageSupportSheet


    
    RefreshCoverageSupportFromHedgeRisco


    
    Application.StatusBar = "Building SwapMap from Coverage support..."
    importedSwapRows = BuildSwapMapFromCoverageSupport()


    
    Application.StatusBar = "Building CoverageFutures from Coverage support..."
    importedFutureRows = BuildCoverageFuturesMapFromCoverageSupport(wsCfg)


    ClearCoverageSupportSheet


    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.AskToUpdateLinks = oldAskLinks
    Application.StatusBar = False


    MsgBox CStr(importedSwapRows) & " swap mapping rows imported into SwapMap." & vbCrLf & _
           CStr(importedFutureRows) & " futures mapping rows imported into CoverageFutures." & vbCrLf & vbCrLf & _
           "Coverage support was cleared after import.", _
           vbInformation, _
           "Coverage Refresh Complete"


    Exit Sub


CleanFail:
    Dim eInfo As String
    eInfo = CaptureErrInfo()


    On Error Resume Next
    ClearCoverageSupportSheet
    On Error GoTo 0


    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.AskToUpdateLinks = oldAskLinks
    Application.StatusBar = False


    MsgBox "Coverage support refresh failed." & vbCrLf & vbCrLf & eInfo, _
           vbCritical, _
           "Coverage Refresh Failed"
End Sub




Private Sub RefreshCoverageSupportFromHedgeRisco()
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim wasOpen As Boolean
    Dim arrValues As Variant
    Dim arrFormulaA As Variant
    Dim i As Long


    Set wsDst = EnsureCoverageSupportSheet()


    On Error Resume Next
    Set wbSrc = Workbooks(Dir$(SWAP_ID_BOOK_PATH))
    On Error GoTo 0


    If wbSrc Is Nothing Then
        Set wbSrc = Workbooks.Open(Filename:=SWAP_ID_BOOK_PATH, UpdateLinks:=0, ReadOnly:=True)
        wasOpen = False
    Else
        wasOpen = True
    End If


    Set wsSrc = wbSrc.Worksheets(SWAP_ID_BOOK_SHEET)


    ' Copy displayed/value data A:AI.
    arrValues = wsSrc.Range("A" & SWAP_ID_FIRST_ROW & ":" & HEDGE_RISCO_COPY_LAST_COL & SWAP_ID_LAST_ROW).Value2


    ' Separately store source formulas from column A in Coverage support column AJ.
    arrFormulaA = wsSrc.Range("A" & SWAP_ID_FIRST_ROW & ":A" & SWAP_ID_LAST_ROW).formula


    wsDst.Range("A" & SWAP_ID_FIRST_ROW & ":" & HEDGE_RISCO_COPY_LAST_COL & SWAP_ID_LAST_ROW).Value2 = arrValues


    wsDst.Range(HEDGE_RISCO_FORMULA_A_COL & SWAP_ID_FIRST_ROW & ":" & HEDGE_RISCO_FORMULA_A_COL & SWAP_ID_LAST_ROW).numberFormat = "@"


    For i = 1 To UBound(arrFormulaA, 1)
        wsDst.Cells(SWAP_ID_FIRST_ROW + i - 1, HEDGE_RISCO_FORMULA_A_COL).value = "'" & CStr(arrFormulaA(i, 1))
    Next i


    If Not wasOpen Then
        wbSrc.Close SaveChanges:=False
    End If
End Sub


' =============================================================================
' HEDGE RISCO TOTAL - the second coverage book
'
' Same three steps as Hedge Risco Tx Juro, against a different file with a
' different Resumo layout:
'
'   1) RefreshCoverageTotalSupportFromHedgeRiscoTotal   copy Resumo A8:W34
'   2) AppendCoverageFuturesMapFromCoverageTotal        filter and map
'   3) ClearCoverageTotalSupportSheet                   drop the temporary copy
'
' Step 2 APPENDS to the same CoverageFutures sheet the Tx Juro pass already
' filled, so the Futures sheet ends up with the Tx Juro positions first and the
' Total positions continuing below them - and every row carries the source tag
' that FuturesRTJ_DV01 / FuturesRT_DV01 split on.
' =============================================================================


Private Function EnsureCoverageTotalSupportSheet() As Worksheet
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_COV_TOTAL_SUPPORT)
    On Error GoTo 0


    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_COV_TOTAL_SUPPORT
    End If


    ws.Range(CELL_COV_A1).value = "Temporary copy of Hedge_Risco Total Resumo A8:W34"
    ws.Range(CELL_COV_A2).value = "This sheet is cleared after CoverageFutures is built."
    ws.Range(RNG_COV_A1_A2).Font.Bold = True


    Set EnsureCoverageTotalSupportSheet = ws
End Function


Private Sub ClearCoverageTotalSupportSheet()
    Dim ws As Worksheet


    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_COV_TOTAL_SUPPORT)
    On Error GoTo 0


    If ws Is Nothing Then Exit Sub


    ws.Range("A" & HEDGE_TOTAL_FIRST_ROW & ":" & HEDGE_TOTAL_FORMULA_A_COL & _
             HEDGE_TOTAL_LAST_ROW).ClearContents
    ws.Range("A" & HEDGE_TOTAL_FIRST_ROW & ":" & HEDGE_TOTAL_FORMULA_A_COL & _
             HEDGE_TOTAL_LAST_ROW).ClearFormats
End Sub


Private Sub RefreshCoverageTotalSupportFromHedgeRiscoTotal()
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim wasOpen As Boolean
    Dim arrValues As Variant
    Dim arrFormulaA As Variant
    Dim i As Long


    Set wsDst = EnsureCoverageTotalSupportSheet()


    On Error Resume Next
    Set wbSrc = Workbooks(Dir$(HEDGE_TOTAL_BOOK_PATH))
    On Error GoTo 0


    If wbSrc Is Nothing Then
        Set wbSrc = Workbooks.Open(Filename:=HEDGE_TOTAL_BOOK_PATH, UpdateLinks:=0, ReadOnly:=True)
        wasOpen = False
    Else
        wasOpen = True
    End If


    Set wsSrc = wbSrc.Worksheets(HEDGE_TOTAL_BOOK_SHEET)


    ' Values only for A:W.  Reading .Value2 rather than copying the range keeps
    ' the source book's own formulas and links out of this workbook entirely.
    arrValues = wsSrc.Range("A" & HEDGE_TOTAL_FIRST_ROW & ":" & _
                            HEDGE_TOTAL_COPY_LAST_COL & HEDGE_TOTAL_LAST_ROW).Value2


    ' Column A's FORMULA text is kept separately, exactly as the Tx Juro import
    ' does: the coverage relation is often a link into another sheet and the
    ' displayed value alone cannot say which relation a row belongs to.
    arrFormulaA = wsSrc.Range("A" & HEDGE_TOTAL_FIRST_ROW & ":A" & _
                              HEDGE_TOTAL_LAST_ROW).formula


    wsDst.Range("A" & HEDGE_TOTAL_FIRST_ROW & ":" & _
                HEDGE_TOTAL_COPY_LAST_COL & HEDGE_TOTAL_LAST_ROW).Value2 = arrValues


    wsDst.Range(HEDGE_TOTAL_FORMULA_A_COL & HEDGE_TOTAL_FIRST_ROW & ":" & _
                HEDGE_TOTAL_FORMULA_A_COL & HEDGE_TOTAL_LAST_ROW).numberFormat = "@"


    For i = 1 To UBound(arrFormulaA, 1)
        wsDst.Cells(HEDGE_TOTAL_FIRST_ROW + i - 1, HEDGE_TOTAL_FORMULA_A_COL).value = _
            "'" & CStr(arrFormulaA(i, 1))
    Next i


    If Not wasOpen Then
        wbSrc.Close SaveChanges:=False
    End If
End Sub


' Append the Hedge Risco Total futures rows underneath whatever the Tx Juro
' pass already wrote into CoverageFutures.  Returns the number of rows added.
Private Function AppendCoverageFuturesMapFromCoverageTotal() As Long
    Dim wsCov As Worksheet
    Dim wsFMap As Worksheet
    Dim r As Long
    Dim outRow As Long
    Dim importedCount As Long


    Dim portfolioType As String
    Dim coverageRelation As String
    Dim linkedISIN As String
    Dim typeLabel As String
    Dim futureCode As String
    Dim counterparty As String
    Dim contracts As Variant
    Dim startDate As Variant
    Dim ccy As String
    Dim bpv As Variant
    Dim statusText As String


    Set wsCov = EnsureCoverageTotalSupportSheet()
    Set wsFMap = EnsureCoverageFuturesSheet()


    ' Continue below the Tx Juro rows rather than clearing: this pass ADDS a
    ' second book, it does not replace the first.
    outRow = wsFMap.Cells(wsFMap.Rows.Count, CFCOL_SOURCEROW).End(xlUp).Row + 1
    If outRow < DATA_ROW Then outRow = DATA_ROW


    For r = HEDGE_TOTAL_FIRST_ROW To HEDGE_TOTAL_LAST_ROW


        portfolioType = CleanText(wsCov.Cells(r, HEDGE_TOTAL_PORTFOLIO_TYPE_COL).value)


        If IsHtCSCoveragePortfolio(portfolioType) Then


            coverageRelation = CleanText(wsCov.Cells(r, "A").value)
            If Len(coverageRelation) = 0 Then
                coverageRelation = CleanText(wsCov.Cells(r, HEDGE_TOTAL_FORMULA_A_COL).value)
            End If


            linkedISIN = UCase$(CleanText(wsCov.Cells(r, "B").value))
            typeLabel = CleanText(wsCov.Cells(r, HEDGE_TOTAL_TYPE_COL).value)
            counterparty = CleanText(wsCov.Cells(r, HEDGE_TOTAL_COUNTERPARTY_COL).value)
            contracts = CoverageNotionalValue(wsCov.Cells(r, HEDGE_TOTAL_NOTIONAL_COL).value)
            startDate = wsCov.Cells(r, HEDGE_TOTAL_STARTDATE_COL).value
            ccy = CleanText(wsCov.Cells(r, HEDGE_TOTAL_CCY_COL).value)
            bpv = CDblSafe(wsCov.Cells(r, HEDGE_TOTAL_BPV_COL).value)


            ' A swap in this book must not become a futures row: it would be
            ' sent to Bloomberg as a contract code and come back #N/A.  It is
            ' still recorded, with a status saying why it was not taken, so a
            ' row is never silently absent from both sheets.
            If IsCoverageSwapTypeLabel(typeLabel) Then
                futureCode = ""
                statusText = "Swap row in Hedge Risco Total - not loaded as a future"
            Else
                futureCode = FutureCodeFromCoverageLabel(typeLabel)


                If Len(linkedISIN) = 0 Then
                    statusText = "Missing LinkedISIN"
                ElseIf Len(futureCode) = 0 Then
                    statusText = "Missing FutureCode"
                ElseIf Not IsNumeric(contracts) Then
                    statusText = "Missing Contracts"
                Else
                    statusText = "OK"
                End If
            End If


            If Len(typeLabel) > 0 Or Len(linkedISIN) > 0 Then


                wsFMap.Cells(outRow, colNum(CFCOL_SOURCEROW)).value = r
                wsFMap.Cells(outRow, colNum(CFCOL_COVERAGERELATION)).value = coverageRelation
                wsFMap.Cells(outRow, colNum(CFCOL_LINKEDISIN)).value = linkedISIN
                wsFMap.Cells(outRow, colNum(CFCOL_COVERAGEINFO_D)).value = counterparty
                wsFMap.Cells(outRow, colNum(CFCOL_FUTURELABEL)).value = typeLabel
                wsFMap.Cells(outRow, colNum(CFCOL_FUTURECODE)).value = futureCode
                wsFMap.Cells(outRow, colNum(CFCOL_HEDGETYPE)).value = typeLabel
                wsFMap.Cells(outRow, colNum(CFCOL_CONTRACTS)).value = contracts
                wsFMap.Cells(outRow, colNum(CFCOL_STARTDATE)).value = startDate
                wsFMap.Cells(outRow, colNum(CFCOL_CCY)).value = ccy
                wsFMap.Cells(outRow, colNum(CFCOL_IMPORTSTATUS)).value = statusText
                wsFMap.Cells(outRow, colNum(CFCOL_HEDGE_SOURCE)).value = HEDGE_SOURCE_RT
                wsFMap.Cells(outRow, colNum(CFCOL_BPV)).value = bpv


                importedCount = importedCount + 1
                outRow = outRow + 1


            End If


        End If


    Next r


    AppendCoverageFuturesMapFromCoverageTotal = importedCount
End Function


' True when a coverage instrument label describes a swap rather than a future.
Private Function IsCoverageSwapTypeLabel(ByVal v As Variant) As Boolean
    Dim s As String


    s = UCase$(CleanText(v))


    IsCoverageSwapTypeLabel = _
        (InStr(1, s, "SWAP", vbTextCompare) > 0) Or _
        (InStr(1, s, "IRS", vbTextCompare) > 0)
End Function


Private Function SwapMatchText(ByVal v As Variant) As String


    SwapMatchText = UCase$(Trim$(CleanText(v)))


End Function


Private Function SwapMatchDate(ByVal v As Variant) As String


    If IsDate(v) Then
        SwapMatchDate = Format$(CDate(v), "yyyymmdd")
    Else
        SwapMatchDate = ""
    End If


End Function


Private Function SwapCoverageGroupKey( _
    ByVal sourceRow As Variant, _
    ByVal coverageRelation As Variant, _
    ByVal linkedISIN As Variant _
) As String


    SwapCoverageGroupKey = _
        SwapMatchText(sourceRow) & "|" & _
        SwapMatchText(coverageRelation) & "|" & _
        SwapMatchText(linkedISIN)


End Function


Private Sub AppendSwapEnrichmentStatus( _
    ByVal wsSw As Worksheet, _
    ByVal swapRow As Long, _
    ByVal statusText As String _
)


    Dim existingStatus As String


    existingStatus = CleanText( _
        wsSw.Cells( _
            swapRow, _
            colNum(WCOL_SWAPMAP_STATUS) _
        ).value _
    )


    If Len(existingStatus) > 0 Then


        wsSw.Cells( _
            swapRow, _
            colNum(WCOL_SWAPMAP_STATUS) _
        ).value = existingStatus & "; " & statusText


    Else


        wsSw.Cells( _
            swapRow, _
            colNum(WCOL_SWAPMAP_STATUS) _
        ).value = statusText


    End If


End Sub


Private Function SwapMatchCounterparty(ByVal v As Variant) As String


    Dim s As String


    s = SwapMatchText(v)
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    s = Replace(s, "-", "")
    s = Replace(s, " ", "")


    SwapMatchCounterparty = s


End Function


Private Function SwapPayFixedKeyExact( _
    ByVal linkedISIN As Variant, _
    ByVal startDate As Variant, _
    ByVal endDate As Variant, _
    ByVal counterparty As Variant _
) As String


    SwapPayFixedKeyExact = _
        SwapMatchText(linkedISIN) & "|" & _
        SwapMatchDate(startDate) & "|" & _
        SwapMatchDate(endDate) & "|" & _
        SwapMatchCounterparty(counterparty)


End Function


Private Function SwapPayFixedKeyDates( _
    ByVal linkedISIN As Variant, _
    ByVal startDate As Variant, _
    ByVal endDate As Variant _
) As String


    SwapPayFixedKeyDates = _
        SwapMatchText(linkedISIN) & "|" & _
        SwapMatchDate(startDate) & "|" & _
        SwapMatchDate(endDate)


End Function


Private Function SwapPayFixedKeyEndDate( _
    ByVal linkedISIN As Variant, _
    ByVal endDate As Variant _
) As String


    SwapPayFixedKeyEndDate = _
        SwapMatchText(linkedISIN) & "|" & _
        SwapMatchDate(endDate)


End Function


Private Sub AddPayFixedCandidate( _
    ByVal d As Object, _
    ByVal key As String, _
    ByVal payFixedYN As String _
)


    If Len(key) = 0 Then Exit Sub
    If Len(payFixedYN) = 0 Then Exit Sub


    If Not d.Exists(key) Then


        d.Add key, payFixedYN


    ElseIf d(key) <> payFixedYN Then


        d(key) = "AMBIGUOUS"


    End If


End Sub


Private Function BuildSwapMapFromCoverageSupport() As Long
    Dim wsCov As Worksheet
    Dim wsMap As Worksheet
    Dim r As Long
    Dim outRow As Long
    Dim importedCount As Long


    Dim formulaA As String
    Dim portfolioType As String
    Dim coverageRelation As String
    Dim linkedISIN As String
    Dim statusText As String


    Dim directId As String
    Dim fixedId As String
    Dim floatId As String


    Dim counterparty As String
    Dim ccy As String
    Dim notional As Variant
    Dim startDate As Variant
    Dim endDate As Variant


    Set wsCov = EnsureCoverageSupportSheet()
    Set wsMap = EnsureSwapMapSheet()


    wsMap.Range(SMCOL_SOURCEROW & DATA_ROW & ":" & SMCOL_NOTIONAL_SOURCE & 1000).ClearContents


    outRow = 5


    For r = SWAP_ID_FIRST_ROW To SWAP_ID_LAST_ROW
        portfolioType = CleanText(wsCov.Cells(r, HEDGE_RISCO_PORTFOLIO_TYPE_COL).value)


        If IsHtCSCoveragePortfolio(portfolioType) Then
            formulaA = CleanText(wsCov.Cells(r, HEDGE_RISCO_FORMULA_A_COL).value)


            If Len(formulaA) = 0 Then
                formulaA = CleanText(wsCov.Cells(r, "A").value)
            End If


            If IsResumoSwapRow(formulaA) Then
                coverageRelation = CleanText(wsCov.Cells(r, "B").value)
                linkedISIN = UCase$(CleanText(wsCov.Cells(r, "C").value))
                counterparty = CleanText(wsCov.Cells(r, "M").value)
                ccy = CleanText(wsCov.Cells(r, "T").value)
                startDate = wsCov.Cells(r, "P").value


                If Len(linkedISIN) = 0 Then
                    statusText = "Missing LinkedISIN"
                Else
                    statusText = "OK"
                End If


                ' Group 1: Plain Vanilla Swap
                ' IDs: I/J/K
                ' Counterparty: M
                ' Notional: O
                ' StartDate: P
                ' EndDate: Q
                ' CCY: T
                directId = SwapIdToBBGSecurity(wsCov.Cells(r, "I").value)
                fixedId = SwapIdToBBGSecurity(wsCov.Cells(r, "J").value)
                floatId = SwapIdToBBGSecurity(wsCov.Cells(r, "K").value)


                If HasAnySwapId(directId, fixedId, floatId) Then
                    notional = CoverageNotionalValue(wsCov.Cells(r, "O").value)
                    endDate = wsCov.Cells(r, "Q").value


                    AppendSwapMapRow wsMap, outRow, r, coverageRelation, linkedISIN, _
                        "PLAIN", directId, fixedId, floatId, statusText, _
                        counterparty, notional, startDate, endDate, ccy, "PLAIN_O"


                    importedCount = importedCount + 1
                End If


                ' Group 2: Synthetic Swap
                ' IDs: AC/AD/AE
                ' Counterparty: M
                ' Notional: AF
                ' StartDate: P
                ' EndDate: AI
                ' CCY: T
                directId = SwapIdToBBGSecurity(wsCov.Cells(r, "AC").value)
                fixedId = SwapIdToBBGSecurity(wsCov.Cells(r, "AD").value)
                floatId = SwapIdToBBGSecurity(wsCov.Cells(r, "AE").value)


                If HasAnySwapId(directId, fixedId, floatId) Then
                    notional = CoverageNotionalValue(wsCov.Cells(r, "AF").value)
                    endDate = wsCov.Cells(r, "AI").value


                    AppendSwapMapRow wsMap, outRow, r, coverageRelation, linkedISIN, _
                        "SYNTHETIC", directId, fixedId, floatId, statusText, _
                        counterparty, notional, startDate, endDate, ccy, "SYNTHETIC_AF"


                    importedCount = importedCount + 1
                End If
            End If
        End If
    Next r


    BuildSwapMapFromCoverageSupport = importedCount
End Function


Private Function BuildCoverageFuturesMapFromCoverageSupport(ByVal wsCfg As Worksheet) As Long
    Dim wsCov As Worksheet
    Dim wsFMap As Worksheet
    Dim r As Long
    Dim outRow As Long
    Dim importedCount As Long


    Dim valueA As String
    Dim formulaA As String
    Dim typeH As String
    Dim portfolioType As String
    Dim coverageRelation As String
    Dim linkedISIN As String
    Dim coverageInfoD As String
    Dim futureLabel As String
    Dim futureCode As String
    Dim contracts As Variant
    Dim startDate As Variant
    Dim ccy As String
    Dim statusText As String


    Set wsCov = EnsureCoverageSupportSheet()
    Set wsFMap = EnsureCoverageFuturesSheet()


    wsFMap.Range(CFCOL_SOURCEROW & DATA_ROW & ":" & CFCOL_BPV & 1000).ClearContents


    outRow = 5


    For r = SWAP_ID_FIRST_ROW To SWAP_ID_LAST_ROW
        portfolioType = CleanText(wsCov.Cells(r, HEDGE_RISCO_PORTFOLIO_TYPE_COL).value)


        If IsHtCSCoveragePortfolio(portfolioType) Then
            valueA = CleanText(wsCov.Cells(r, "A").value)
            formulaA = CleanText(wsCov.Cells(r, HEDGE_RISCO_FORMULA_A_COL).value)
            typeH = CleanText(wsCov.Cells(r, "H").value)


            If IsCoverageFutureRow(valueA, formulaA, typeH) Then
                coverageRelation = CleanText(wsCov.Cells(r, "B").value)
                linkedISIN = UCase$(CleanText(wsCov.Cells(r, "C").value))
                coverageInfoD = CleanText(wsCov.Cells(r, "D").value)
                futureLabel = valueA
                futureCode = FutureCodeFromCoverageLabel(futureLabel)
                contracts = CoverageNotionalValue(wsCov.Cells(r, "O").value)
                startDate = wsCov.Cells(r, "P").value
                ccy = CleanText(wsCov.Cells(r, "T").value)


                If Len(linkedISIN) = 0 Then
                    statusText = "Missing LinkedISIN"
                ElseIf Len(futureCode) = 0 Then
                    statusText = "Missing FutureCode"
                ElseIf Not IsNumeric(contracts) Then
                    statusText = "Missing Contracts"
                Else
                    statusText = "OK"
                End If


                wsFMap.Cells(outRow, colNum(CFCOL_SOURCEROW)).value = r
                wsFMap.Cells(outRow, colNum(CFCOL_COVERAGERELATION)).value = coverageRelation
                wsFMap.Cells(outRow, colNum(CFCOL_LINKEDISIN)).value = linkedISIN
                wsFMap.Cells(outRow, colNum(CFCOL_COVERAGEINFO_D)).value = coverageInfoD
                wsFMap.Cells(outRow, colNum(CFCOL_FUTURELABEL)).value = futureLabel
                wsFMap.Cells(outRow, colNum(CFCOL_FUTURECODE)).value = futureCode
                wsFMap.Cells(outRow, colNum(CFCOL_HEDGETYPE)).value = typeH
                wsFMap.Cells(outRow, colNum(CFCOL_CONTRACTS)).value = contracts
                wsFMap.Cells(outRow, colNum(CFCOL_STARTDATE)).value = startDate
                wsFMap.Cells(outRow, colNum(CFCOL_CCY)).value = ccy
                wsFMap.Cells(outRow, colNum(CFCOL_IMPORTSTATUS)).value = statusText
                wsFMap.Cells(outRow, colNum(CFCOL_HEDGE_SOURCE)).value = HEDGE_SOURCE_RTJ


                importedCount = importedCount + 1
                outRow = outRow + 1
            End If
        End If
    Next r


    BuildCoverageFuturesMapFromCoverageSupport = importedCount
End Function
Public Sub Debug_CheckSwapMapAndSwaps()
    Dim wsMap As Worksheet
    Dim wsSw As Worksheet
    Dim lastMapRow As Long
    Dim lastSwRow As Long


    Set wsMap = ThisWorkbook.Worksheets(SH_SWAPMAP)
    Set wsSw = ThisWorkbook.Worksheets(SH_SWAPS)


    lastMapRow = wsMap.Cells(wsMap.Rows.Count, SMCOL_SOURCEROW).End(xlUp).Row
    lastSwRow = wsSw.Cells(wsSw.Rows.Count, WCOL_DEALID).End(xlUp).Row


    MsgBox "SwapMap last row: " & lastMapRow & vbCrLf & _
           "SwapMap data rows: " & Application.Max(0, lastMapRow - 4) & vbCrLf & _
           "Swaps last row: " & lastSwRow & vbCrLf & _
           "Swaps data rows: " & Application.Max(0, lastSwRow - 4), vbInformation
End Sub


Public Sub RestoreBondMarketFormulas_Debug()


    Dim wsBnd As Worksheet
    Dim wsOIS As Worksheet
    Dim lastBondRow As Long


    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldStatusBar As Variant


    Set wsBnd = ThisWorkbook.Worksheets(SH_BONDS)
    Set wsOIS = ThisWorkbook.Worksheets(SH_OIS)


    lastBondRow = LastBondDataRow(wsBnd)


    If lastBondRow < BOND_DATA_ROW Then
        MsgBox "No bond rows found.", vbExclamation
        Exit Sub
    End If


    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldStatusBar = Application.StatusBar


    On Error GoTo CleanFail


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Restoring bond market/Bloomberg formulas for debugging..."


    mLastFormulaTarget = ""
    mLastRangeTarget = ""


    ' Recreate ticker candidates and resolved ticker.
    WriteBondSecurityCandidates wsBnd, lastBondRow


    wsBnd.Range(BCOL_BBG_TICKER & BOND_DATA_ROW & ":" & BCOL_BBG_TICKER & lastBondRow).FormulaR1C1 = _
        BBGParseFormulaR1C1()


    wsBnd.Range(BCOL_TICKER_STATUS & BOND_DATA_ROW & ":" & BCOL_TICKER_STATUS & lastBondRow).FormulaR1C1 = WrapIfPresent(RC(BCOL_ISIN), "IF(OR(" & RC(BCOL_BBG_TICKER) & "=""""," & RC(BCOL_BBG_TICKER) & "=""UNKNOWN""),""UNKNOWN"",""OK"")")


    ' Reinsert Bloomberg formulas.
    WriteBonds_T1_BDP_Efficient wsBnd, lastBondRow


    ' Reinsert derived formulas.
    WriteBondsCalculatedFormulas_Efficient wsBnd, lastBondRow
    WriteBondConvexityBumpFormulas wsBnd, lastBondRow


    wsOIS.Calculate
    wsBnd.Calculate


CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar


    mLastFormulaTarget = ""
    mLastRangeTarget = ""


    If Err.Number = 0 Then
        MsgBox "Bond market formulas restored for debugging. No freezing was applied.", vbInformation
    End If


    Exit Sub


CleanFail:
    Dim eInfo As String
    eInfo = CaptureErrInfo()


    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatusBar


    mLastFormulaTarget = ""
    mLastRangeTarget = ""


    MsgBox "RestoreBondMarketFormulas_Debug failed:" & vbCrLf & eInfo, vbCritical


End Sub





' =============================================================================
' AssertGeometryConstants  (debug-only: run once to confirm the letter<->number
' bridges still line up; catches silent drift between the legacy numeric COL_*
' constants and the new letter-string column constants).
' =============================================================================
Public Sub AssertGeometryConstants()

    ' Asserts only what CANNOT be checked from the source text.
    '
    ' tools/check_layout.py already proves the letter constants are unique,
    ' gapless and agree with their "' col N =" comments, so repeating that here
    ' would be duplication.  What it cannot see is the handful of places where a
    ' column is ALSO held as a raw number, because Find/Resize/Cells want an
    ' index rather than a letter.  Those pairs are what this checks.
    '
    ' The previous version additionally froze a snapshot of the old geometry
    ' (colNum(BCOL_BONDDCC_NAME) = 92, RC(BCOL_GOV_T0) = "RC55", and so on).
    ' Those literals proved a one-off refactor had not moved anything, years
    ' ago; kept afterwards they simply assert that no column may ever move, and
    ' they fired the moment PositionSide was removed.  A test that must be
    ' edited every time the thing it guards changes legitimately is not a test.

    Debug.Assert colNum(BCOL_ISIN) = 1
    Debug.Assert colNum(BCOL_BBG_TICKER) = COL_BBG_TICKER
    Debug.Assert colNum(BCOL_TICKER_STATUS) = COL_BBG_TICKER_STATUS
    Debug.Assert colNum(BCOL_BBG_CAND_ISIN) = COL_BBG_CAND_FIRST
    Debug.Assert colNum(BCOL_BBG_CAND_MMKT) = COL_BBG_CAND_LAST

    ' The candidate block must stay contiguous: BBGParseFormulaR1C1 walks it by
    ' index from first to last.
    Debug.Assert COL_BBG_CAND_LAST > COL_BBG_CAND_FIRST

    ' PNL_LAST_COL bounds the range every recalculation calculates, so it must
    ' be the last column the header writer actually fills.
    Debug.Assert colNum(PCOL_COUPON_PAID_EUR) = colNum(PNL_LAST_COL)
    Debug.Assert colNum(PCOL_RISK_TIMING_BIAS) = colNum(PNL_LAST_COL) - 1
    Debug.Assert colNum(PNL_CLEAR_LAST_COL) >= colNum(PNL_LAST_COL)

    ' The macro block must begin immediately after the OPICS query's last column.
    Debug.Assert colNum(BCOL_DAYSLEFT) = colNum(BONDS_QUERY_LAST_COL) + 1

    Debug.Assert colNum(WCOL_FLOATFAMILY_STATUS) = colNum(WCOL_DV01_BBG) + 1

    MsgBox "Geometry constants OK.", vbInformation

End Sub


' =============================================================================
' THE POSITION STORE BRIDGE
'
' modAccess keeps the runs.  It must not know where a column sits, and this
' module must not know what a database is - so the two meet here, on field
' NAMES.
'
' PnlPositionFieldMap is the whole contract: one entry per stored field, each
' "FieldName|ColumnLetter", the letter coming from the BCOL_/WCOL_/FCOL_
' constant rather than a literal.  Move a column by editing its constant, as
' always, and the store follows it with no other edit anywhere.  The field
' names are the Access column names, and tools/check_access_schema.py fails the
' build if the two lists ever disagree.
'
' Only INPUT columns appear here.  Everything a formula computes is left out on
' purpose: it is derived from these, so storing it would be storing the same
' fact twice, and restoring it would overwrite the formula that produces it.
' =============================================================================

' STORE_KIND_BOND / _SWAP / _FUTURE and STORE_ROWNUM_TOKEN are declared with
' the other module constants at the top of the file, because VBA requires it.


Public Function PnlPositionFieldMap(ByVal kind As String) As Variant

    Select Case UCase$(Trim$(kind))

    Case STORE_KIND_BOND
        ' Bonds!A:K, the block the OPICS query owns.  Nothing to its right is
        ' stored: every one of those columns is a formula.
        PnlPositionFieldMap = Array( _
            "ISIN|" & BCOL_ISIN, _
            "InstrumentName|" & BCOL_NAME, _
            "CCY|" & BCOL_CCY, _
            "Coupon|" & BCOL_COUPON, _
            "CouponFreqDesc|" & BCOL_COUPON_FREQ, _
            "Maturity|" & BCOL_MATURITY, _
            "Notional|" & BCOL_NOTIONAL, _
            "AcctgCat|" & BCOL_ACCTGCAT, _
            "Portfolio|" & BCOL_PORTFOLIO, _
            "BookVal|" & BCOL_BOOKVAL, _
            "SourceRow|" & STORE_ROWNUM_TOKEN)

    Case STORE_KIND_SWAP
        ' Everything AppendSwapMapRowsToSwaps writes, plus PayFixed, which the
        ' OPICS enrichment pass fills in afterwards.
        PnlPositionFieldMap = Array( _
            "PositionKey|" & WCOL_DEALID, _
            "CCY|" & WCOL_CCY, _
            "Notional|" & WCOL_NOTIONAL, _
            "StartDate|" & WCOL_STARTDATE, _
            "EndDate|" & WCOL_ENDDATE, _
            "PayFixed|" & WCOL_PAYFIXED, _
            "Portfolio|" & WCOL_PORTFOLIO, _
            "LinkedISIN|" & WCOL_LINKEDISIN, _
            "Cpty|" & WCOL_CPTY, _
            "BBGSwapDirectID|" & WCOL_BBG_SWAP_DIRECT_ID, _
            "BBGFixedLegID|" & WCOL_BBG_FIXED_LEG_ID, _
            "BBGFloatLegID|" & WCOL_BBG_FLOAT_LEG_ID, _
            "SwapIDSource|" & WCOL_SWAP_ID_SOURCE, _
            "CoverageRelation|" & WCOL_COVERAGERELATION, _
            "SwapMapSourceRow|" & WCOL_SWAPMAP_SOURCEROW, _
            "SwapMapClass|" & WCOL_SWAPMAP_CLASS, _
            "SwapMapStatus|" & WCOL_SWAPMAP_STATUS, _
            "MapNotional|" & WCOL_MAP_NOTIONAL, _
            "MapCCY|" & WCOL_MAP_CCY, _
            "MapCounterparty|" & WCOL_MAP_COUNTERPARTY, _
            "NotionalFinal|" & WCOL_NOTIONAL_FINAL, _
            "NotionalSource|" & WCOL_NOTIONAL_SOURCE, _
            "SourceRow|" & STORE_ROWNUM_TOKEN)

    Case STORE_KIND_FUTURE
        ' Everything AppendCoverageFuturesRowsToFutures writes, plus Exchange,
        ' which EnrichCoverageFuturesFromOPICS fills in afterwards.  Hedge_Class
        ' is NOT here: it is a formula over HedgeType.
        PnlPositionFieldMap = Array( _
            "ContractCode|" & FCOL_CONTRACTCODE, _
            "Exchange|" & FCOL_EXCHANGE, _
            "CCY|" & FCOL_CCY, _
            "Contracts|" & FCOL_CONTRACTS, _
            "Portfolio|" & FCOL_PORTFOLIO, _
            "LinkedISIN|" & FCOL_LINKEDISIN, _
            "ImportStatus|" & FCOL_STATUS, _
            "CoverageSourceRow|" & FCOL_COVERAGE_SOURCEROW, _
            "CoverageRelation|" & FCOL_COVERAGERELATION, _
            "FutureLabel|" & FCOL_FUTURELABEL, _
            "HedgeType|" & FCOL_HEDGETYPE, _
            "CoverageStartDate|" & FCOL_COVERAGE_STARTDATE, _
            "CoverageInfoD|" & FCOL_COVERAGEINFO_D, _
            "HedgeSource|" & FCOL_HEDGE_SOURCE, _
            "CoverageBPV|" & FCOL_COVERAGE_BPV, _
            "SourceRow|" & STORE_ROWNUM_TOKEN)

    Case Else
        PnlPositionFieldMap = Array()

    End Select

End Function


' The sheet, and the first row that holds a position on it.
Private Function StoreSheet(ByVal kind As String) As Worksheet

    Select Case UCase$(Trim$(kind))
    Case STORE_KIND_BOND
        Set StoreSheet = ThisWorkbook.Worksheets(SH_BONDS)
    Case STORE_KIND_SWAP
        Set StoreSheet = ThisWorkbook.Worksheets(SH_SWAPS)
    Case STORE_KIND_FUTURE
        Set StoreSheet = ThisWorkbook.Worksheets(SH_FUTURES)
    Case Else
        Set StoreSheet = Nothing
    End Select

End Function


Private Function StoreFirstRow(ByVal kind As String) As Long

    If UCase$(Trim$(kind)) = STORE_KIND_BOND Then
        StoreFirstRow = BOND_DATA_ROW
    Else
        StoreFirstRow = DATA_ROW
    End If

End Function


Private Function StoreLastRow(ByVal kind As String, ByVal ws As Worksheet) As Long

    Select Case UCase$(Trim$(kind))
    Case STORE_KIND_BOND
        StoreLastRow = LastBondDataRow(ws)
    Case STORE_KIND_SWAP
        StoreLastRow = LastSwapDataRow(ws)
    Case STORE_KIND_FUTURE
        StoreLastRow = LastFutureDataRow(ws)
    Case Else
        StoreLastRow = 0
    End Select

End Function


' Field names only - what modAccess needs to build a table without reading
' cells, and what the schema checker compares against access/schema.sql.
Public Function PnlPositionFieldNames(ByVal kind As String) As Variant

    Dim spec As Variant
    Dim names() As String
    Dim i As Long

    spec = PnlPositionFieldMap(kind)

    If UBound(spec) < LBound(spec) Then
        PnlPositionFieldNames = Array()
        Exit Function
    End If

    ReDim names(LBound(spec) To UBound(spec))

    For i = LBound(spec) To UBound(spec)
        names(i) = Split(CStr(spec(i)), "|")(0)
    Next i

    PnlPositionFieldNames = names

End Function


' How many positions are on the sheet right now.  Answered without reading the
' values, so modAccess can size a progress bar or refuse an empty save cheaply.
Public Function PnlPositionCount(ByVal kind As String) As Long

    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long

    Set ws = StoreSheet(kind)
    If ws Is Nothing Then Exit Function

    firstRow = StoreFirstRow(kind)
    lastRow = StoreLastRow(kind, ws)

    If lastRow < firstRow Then Exit Function

    PnlPositionCount = lastRow - firstRow + 1

End Function


' Every position of one kind, as a 2-D Variant:
'
'   v(0, j)  field name          (row 0 is the header, always)
'   v(i, j)  position i's value  (i = 1 .. count)
'
' One .Value read per COLUMN, not per cell: a 600-row book is eleven reads
' rather than six thousand, which is the difference between instant and a
' visible pause.  An empty sheet returns a one-row array holding only the
' header, so the caller never has to special-case Empty.
Public Function PnlPositionsForStore(ByVal kind As String) As Variant

    Dim ws As Worksheet
    Dim spec As Variant
    Dim outv() As Variant
    Dim colLetter As String
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rowCount As Long
    Dim i As Long
    Dim j As Long
    Dim block As Variant

    spec = PnlPositionFieldMap(kind)
    Set ws = StoreSheet(kind)

    If ws Is Nothing Or UBound(spec) < LBound(spec) Then
        PnlPositionsForStore = Empty
        Exit Function
    End If

    firstRow = StoreFirstRow(kind)
    lastRow = StoreLastRow(kind, ws)
    If lastRow < firstRow Then lastRow = firstRow - 1

    rowCount = lastRow - firstRow + 1
    If rowCount < 0 Then rowCount = 0

    ReDim outv(0 To rowCount, LBound(spec) To UBound(spec))

    For j = LBound(spec) To UBound(spec)

        outv(0, j) = Split(CStr(spec(j)), "|")(0)

        If rowCount > 0 Then

            colLetter = Split(CStr(spec(j)), "|")(1)

            If colLetter = STORE_ROWNUM_TOKEN Then
                For i = 1 To rowCount
                    outv(i, j) = firstRow + i - 1
                Next i
            Else
                block = ws.Range(colLetter & firstRow & ":" & _
                                 colLetter & lastRow).value

                If rowCount = 1 Then
                    outv(1, j) = block
                Else
                    For i = 1 To rowCount
                        outv(i, j) = block(i, 1)
                    Next i
                End If
            End If

        End If

    Next j

    PnlPositionsForStore = outv

End Function


' Put a stored run back on the sheet.
'
' Writes ONLY the mapped input columns - a formula column is never touched,
' because everything in it is derived from what is being written and will
' recompute.  Columns are cleared to the sheet's full previous extent first, so
' a smaller restored run cannot leave the tail of a larger one behind.
'
' The Bonds sheet is the awkward one: A:K belongs to the OPICS query's
' ListObject, and writing 380 rows into a table sized for 600 leaves 220 rows
' of a table that still claims to hold data.  So the table is RESIZED to the
' restored count first.  The next query refresh overwrites all of it, which is
' correct - a restored run is something you are looking at, not a new source of
' truth.
Public Function PnlRestorePositions( _
    ByVal kind As String, _
    ByVal data As Variant) As Long

    Dim ws As Worksheet
    Dim spec As Variant
    Dim colLetter As String
    Dim firstRow As Long
    Dim clearLast As Long
    Dim rowCount As Long
    Dim srcCol As Long
    Dim i As Long
    Dim j As Long
    Dim block() As Variant
    Dim fieldName As String

    Set ws = StoreSheet(kind)
    spec = PnlPositionFieldMap(kind)

    If ws Is Nothing Or UBound(spec) < LBound(spec) Then Exit Function
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    firstRow = StoreFirstRow(kind)
    rowCount = UBound(data, 1) - LBound(data, 1)
    If rowCount < 0 Then rowCount = 0
    If rowCount > MAX_SHEET_ROWS Then rowCount = MAX_SHEET_ROWS

    clearLast = SheetClearLastRow(ws, firstRow, StoreLastRow(kind, ws))

    If UCase$(Trim$(kind)) = STORE_KIND_BOND Then
        ResizeBondsTableForRestore ws, rowCount
    End If

    For j = LBound(spec) To UBound(spec)

        fieldName = Split(CStr(spec(j)), "|")(0)
        colLetter = Split(CStr(spec(j)), "|")(1)

        If colLetter <> STORE_ROWNUM_TOKEN Then

            ws.Range(colLetter & firstRow & ":" & _
                     colLetter & clearLast).ClearContents

            srcCol = StoreColumnOf(data, fieldName)

            If srcCol >= 0 And rowCount > 0 Then

                ReDim block(1 To rowCount, 1 To 1)

                For i = 1 To rowCount
                    block(i, 1) = data(LBound(data, 1) + i, srcCol)
                Next i

                ws.Range(colLetter & firstRow & ":" & _
                         colLetter & (firstRow + rowCount - 1)).value = block

            End If

        End If

    Next j

    PnlRestorePositions = rowCount

End Function


' Which column of the stored array holds a field, by name.  -1 when the stored
' run predates the field - restoring an older run then leaves that column blank
' rather than failing, which is the honest answer.
Private Function StoreColumnOf( _
    ByVal data As Variant, _
    ByVal fieldName As String) As Long

    Dim j As Long

    StoreColumnOf = -1

    For j = LBound(data, 2) To UBound(data, 2)
        If StrComp(CStr(data(LBound(data, 1), j)), fieldName, vbTextCompare) = 0 Then
            StoreColumnOf = j
            Exit Function
        End If
    Next j

End Function


' Make the query table exactly as tall as the run being restored.
'
' Silent when the sheet has no ListObject: a book set up without the query is
' unusual but not broken, and the values still land in A:K either way.
Private Sub ResizeBondsTableForRestore( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long)

    Dim lo As ListObject
    Dim bodyRows As Long

    If ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)

    bodyRows = rowCount
    If bodyRows < 1 Then bodyRows = 1

    On Error Resume Next
    lo.Resize ws.Range( _
        lo.Range.Cells(1, 1).Address & ":" & _
        BONDS_QUERY_LAST_COL & (BOND_DATA_ROW + bodyRows - 1))
    On Error GoTo 0

End Sub


' Everything Button 2 does EXCEPT reloading the hedge rows.
'
' A restored run already has its rows; what it does not have is formulas sized
' to them.  Calling Button 2 would refetch from Hedge Risco and throw the
' restored population away, so the formula half is available on its own.
Public Function PnlRewriteFormulasForRestoredRun() As String

    Dim wb As Workbook
    Dim wsBnd As Worksheet
    Dim wsFut As Worksheet
    Dim wsSw As Worksheet
    Dim wsPnl As Worksheet
    Dim wsOIS As Worksheet
    Dim lastBondRow As Long
    Dim problems As String
    Dim oldCalc As XlCalculation

    Set wb = ThisWorkbook
    Set wsBnd = wb.Worksheets(SH_BONDS)
    Set wsFut = wb.Worksheets(SH_FUTURES)
    Set wsSw = wb.Worksheets(SH_SWAPS)
    Set wsPnl = wb.Worksheets(SH_PNL)
    Set wsOIS = wb.Worksheets(SH_OIS)

    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    lastBondRow = LastBondDataRow(wsBnd)

    problems = problems & WriteCurveFormulasSafe(wsOIS)

    If lastBondRow >= BOND_DATA_ROW Then
        problems = problems & WriteBondFormulasSafe(wsBnd, lastBondRow)
    End If

    problems = problems & WriteFuturesFormulasSafe( _
        wsFut, LastFutureDataRow(wsFut))
    problems = problems & WriteSwapFormulasSafe( _
        wsSw, LastSwapDataRow(wsSw))

    If lastBondRow >= BOND_DATA_ROW Then
        problems = problems & WritePNLSectionSafe(wsPnl, wsBnd, lastBondRow)
        PublishPnlColumnNames DATA_ROW + PnlRowCount(lastBondRow) - 1
    End If

    Application.Calculation = oldCalc
    Application.CalculateFullRebuild

    PnlRewriteFormulasForRestoredRun = problems

End Function


' The reporting dates, as the DATABASE means them rather than as Config holds
' them.  Config!B4 is labelled "T-1 date" and Config!B5 "T0 date" - the two
' constants are named the other way round, which is a trap the store must not
' inherit.  Answering it once, here, is why modAccess never reads Config.
Public Function PnlAsOfDates() As Variant

    Dim wsCfg As Worksheet
    Dim t0 As Variant
    Dim tm1 As Variant

    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)

    ' CFG_T1_DATE holds T0 and CFG_T0_DATE holds T-1.  Not a typo.
    t0 = wsCfg.Range(CFG_T1_DATE).value
    tm1 = wsCfg.Range(CFG_T0_DATE).value

    PnlAsOfDates = Array(t0, tm1)

End Function
