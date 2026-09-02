-- =============================================================================
-- PNL_Data.accdb  -  schema
--
-- Access DDL dialect (ACE).  Run the statements in order through
-- tools/build_access_db.py, or paste them one at a time into an Access query
-- in SQL view.  Access executes ONE statement per Execute call - it has no
-- batch separator - which is why they are separated by ";" and a blank line
-- and why the builder splits on that.
--
-- Conventions:
--   *  every fact table carries RunID; that column IS the history mechanism
--   *  facts are immutable - a run inserts its own rows and never updates
--      another run's
--   *  dates are stored as the USER-FACING convention (T0 = current reporting
--      date, TM1 = prior), NOT the workbook's inverted CFG_T0_DATE/CFG_T1_DATE
--   *  text keys are TEXT(n) with an explicit length; Access defaults to 255
--      and indexes on 255-char text are wasteful
--
-- See docs/ACCESS_ARCHITECTURE.md.
-- =============================================================================


-- --- the run spine ----------------------------------------------------------
-- One row per execution.  Everything else hangs off this.

CREATE TABLE [Run] (
    RunID            AUTOINCREMENT    NOT NULL,
    RunStartedUtc    DATETIME         NOT NULL,
    RunFinishedUtc   DATETIME,
    AsOfT0           DATETIME         NOT NULL,
    AsOfTM1          DATETIME         NOT NULL,
    RunBy            TEXT(64),
    ModuleVersion    TEXT(32),
    ExcelBitness     TEXT(8),
    RunStatus        TEXT(16)         NOT NULL,
    RunFingerprint   TEXT(64),
    BondCount        LONG,
    SwapCount        LONG,
    FutureCount      LONG,
    SupersedesRunID  LONG,
    Notes            LONGTEXT,
    CONSTRAINT PK_Run PRIMARY KEY (RunID)
);

CREATE INDEX IX_Run_AsOfT0 ON [Run] (AsOfT0);

CREATE UNIQUE INDEX IX_Run_AsOf_Started ON [Run] (AsOfT0, RunStartedUtc);

-- The duplicate-run guard reads this.  RunFingerprint is a hash of every
-- position key the run holds, so two saves of an unchanged book collide here
-- and the second is refused instead of doubling the history.
CREATE INDEX IX_Run_Fingerprint ON [Run] (AsOfT0, RunFingerprint);


-- --- dimensions -------------------------------------------------------------
-- What a thing IS, independent of any one run.

CREATE TABLE Instrument (
    ISIN            TEXT(12)   NOT NULL,
    InstrumentName  TEXT(128),
    CCY             TEXT(3),
    Coupon          DOUBLE,
    CouponFreq      LONG,
    Maturity        DATETIME,
    DayCountCode    LONG,
    FirstSeenRunID  LONG,
    LastSeenRunID   LONG,
    CONSTRAINT PK_Instrument PRIMARY KEY (ISIN)
);

CREATE TABLE Portfolio (
    PortfolioCode   TEXT(32)   NOT NULL,
    AcctgCat        TEXT(32),
    Description     TEXT(128),
    CONSTRAINT PK_Portfolio PRIMARY KEY (PortfolioCode)
);

-- The #RC coverage relation, as a thing with an identity rather than a string
-- reconstructed by a forward-fill on every run.
--
-- SourceBook scopes RCKey: the two Hedge_Risco books number their relations
-- INDEPENDENTLY, so "163" in Tx Juro is not "163" in Total.  That is why the
-- unique index is on the pair and not on RCKey alone.

CREATE TABLE CoverageRelation (
    CoverageRelationID  AUTOINCREMENT  NOT NULL,
    SourceBook          TEXT(4)        NOT NULL,
    RCKey               TEXT(64)       NOT NULL,
    CoveredISIN         TEXT(12),
    FirstSeenRunID      LONG,
    LastSeenRunID       LONG,
    CONSTRAINT PK_CoverageRelation PRIMARY KEY (CoverageRelationID)
);

CREATE UNIQUE INDEX IX_CoverageRelation_Key
    ON CoverageRelation (SourceBook, RCKey);


-- --- the positions ----------------------------------------------------------
-- THE THREE TABLES THE WORKBOOK ACTUALLY WRITES.  modAccess owns them; every
-- other table in this file is design ahead of implementation.
--
-- Shape is the same for all three, because the three problems are the same:
--
--   RunID        separates the runs.  It is the ONLY history mechanism; a run
--                never edits another run's rows.
--   PositionKey  separates the positions WITHIN a run.  It is built in VBA
--                from the fields that make a position distinct, and being in
--                the primary key is what makes a duplicate impossible rather
--                than merely unlikely - Access refuses the second insert.
--   SourceRow    where the row sat on the sheet, so a stored position can be
--                traced back to what produced it.
--
-- The natural-key components are stored as their own columns as well, because
-- a key you can only match on as a whole string is not a key you can query.
--
-- These replace the Raw_OpicsBond / Raw_OpicsHedge sketch that used to sit
-- here.  That sketch keyed bonds on (RunID, ISIN, Portfolio), which silently
-- merges one ISIN held in two accounting categories, and keyed hedges on a
-- DealID the Hedge Risco books do not supply.

CREATE TABLE Pos_Bond (
    RunID           LONG       NOT NULL,
    PositionKey     TEXT(200)  NOT NULL,
    ISIN            TEXT(12),
    InstrumentName  TEXT(128),
    CCY             TEXT(8),
    Coupon          DOUBLE,
    CouponFreqDesc  TEXT(32),
    Maturity        DATETIME,
    Notional        DOUBLE,
    AcctgCat        TEXT(32),
    Portfolio       TEXT(32),
    BookVal         DOUBLE,
    SourceRow       LONG,
    CONSTRAINT PK_Pos_Bond PRIMARY KEY (RunID, PositionKey)
);

CREATE INDEX IX_Pos_Bond_ISIN ON Pos_Bond (RunID, ISIN);

-- One row per MAPPING row, not per swap: the same swap hedging two bonds is
-- two positions, and PositionKey (the sheet's own mapping-row id) already
-- carries LinkedISIN, so the two never collapse into one.

CREATE TABLE Pos_Swap (
    RunID              LONG       NOT NULL,
    PositionKey        TEXT(200)  NOT NULL,
    CCY                TEXT(8),
    Notional           DOUBLE,
    StartDate          DATETIME,
    EndDate            DATETIME,
    PayFixed           TEXT(4),
    Portfolio          TEXT(64),
    LinkedISIN         TEXT(12),
    Cpty               TEXT(128),
    BBGSwapDirectID    TEXT(64),
    BBGFixedLegID      TEXT(64),
    BBGFloatLegID      TEXT(64),
    SwapIDSource       TEXT(16),
    CoverageRelation   TEXT(64),
    SwapMapSourceRow   TEXT(32),
    SwapMapClass       TEXT(16),
    SwapMapStatus      TEXT(128),
    MapNotional        DOUBLE,
    MapCCY             TEXT(8),
    MapCounterparty    TEXT(128),
    NotionalFinal      DOUBLE,
    NotionalSource     TEXT(32),
    SourceRow          LONG,
    CONSTRAINT PK_Pos_Swap PRIMARY KEY (RunID, PositionKey)
);

CREATE INDEX IX_Pos_Swap_Bond ON Pos_Swap (RunID, LinkedISIN);

-- HedgeSource is part of the key, not decoration.  The two Hedge Risco books
-- number their coverage rows INDEPENDENTLY, so row 163 of Tx Juro and row 163
-- of Total are different positions; without the book in the key one of them
-- would be discarded as a duplicate of the other.

CREATE TABLE Pos_Future (
    RunID              LONG       NOT NULL,
    PositionKey        TEXT(200)  NOT NULL,
    ContractCode       TEXT(64),
    Exchange           TEXT(32),
    CCY                TEXT(8),
    Contracts          DOUBLE,
    Portfolio          TEXT(64),
    LinkedISIN         TEXT(12),
    ImportStatus       TEXT(128),
    CoverageSourceRow  TEXT(32),
    CoverageRelation   TEXT(64),
    FutureLabel        TEXT(128),
    HedgeType          TEXT(32),
    CoverageStartDate  DATETIME,
    CoverageInfoD      TEXT(128),
    HedgeSource        TEXT(4),
    CoverageBPV        DOUBLE,
    SourceRow          LONG,
    CONSTRAINT PK_Pos_Future PRIMARY KEY (RunID, PositionKey)
);

CREATE INDEX IX_Pos_Future_Bond ON Pos_Future (RunID, LinkedISIN);


-- --- what a save had to refuse ----------------------------------------------
-- A duplicate key, a row with no key at all, a value too long for its column.
-- Written by the same transaction that writes the positions, so "the run saved
-- clean" is a fact in the database rather than a message box nobody kept.

CREATE TABLE Run_Issue (
    RunID       LONG      NOT NULL,
    IssueSeq    LONG      NOT NULL,
    Severity    TEXT(8)   NOT NULL,
    TableName   TEXT(32),
    PositionKey TEXT(200),
    SourceRow   LONG,
    Detail      TEXT(255),
    CONSTRAINT PK_Run_Issue PRIMARY KEY (RunID, IssueSeq)
);


-- --- what version of the schema this file is --------------------------------
-- modAccess creates missing tables and adds missing columns on every connect,
-- then stamps this.  A database opened by an older module is therefore visible
-- as such instead of failing on a column that is not there yet.

CREATE TABLE Meta_Schema (
    SchemaVersion  LONG      NOT NULL,
    AppliedUtc     DATETIME  NOT NULL,
    AppliedBy      TEXT(64),
    CONSTRAINT PK_Meta_Schema PRIMARY KEY (SchemaVersion)
);


-- --- raw facts --------------------------------------------------------------
-- Exactly as retrieved.  Never edited, never recomputed.  Design ahead of
-- implementation - see docs/ACCESS_ARCHITECTURE.md for the migration phases.

-- A byte-for-byte copy of what the coverage books said on the morning of the
-- run.  Those files are overwritten in place by the desk, so this is the ONLY
-- way that information is ever kept.  Columns are text because the point is
-- fidelity, not interpretation - the interpretation is in Fact_HedgePosition.

CREATE TABLE Raw_CoverageRow (
    RunID       LONG      NOT NULL,
    SourceBook  TEXT(4)   NOT NULL,
    SourceRow   LONG      NOT NULL,
    RCKeyRaw    TEXT(255),
    FormulaA    TEXT(255),
    ColValues   LONGTEXT,
    CONSTRAINT PK_Raw_CoverageRow PRIMARY KEY (RunID, SourceBook, SourceRow)
);

-- TALL on purpose: one row per (security, field, date).  The vendor field set
-- is open-ended and sparse, so a wide table would be mostly nulls and would
-- need a schema change every time somebody wants a new field.

CREATE TABLE Raw_BloombergPoint (
    RunID         LONG      NOT NULL,
    SecurityID    TEXT(64)  NOT NULL,
    FieldName     TEXT(64)  NOT NULL,
    SnapshotDate  DATETIME  NOT NULL,
    ValueNum      DOUBLE,
    ValueText     TEXT(255),
    [Status]      TEXT(32)
);

CREATE UNIQUE INDEX IX_Raw_BloombergPoint_Key
    ON Raw_BloombergPoint (RunID, SecurityID, FieldName, SnapshotDate);


-- --- the curve, tall ---------------------------------------------------------
-- OIS_Curves lays three currencies out side by side across 46 columns, which is
-- one repeating shape written three times.  As a table it is one shape, and
-- adding a fourth currency is ROWS rather than a new 15-column block.
--
-- It also makes the interpolation a single indexed lookup instead of a
-- per-currency Select Case that resolves to a different tenor column each time.

CREATE TABLE Curve_Point (
    RunID         LONG      NOT NULL,
    CCY           TEXT(3)   NOT NULL,
    CurveType     TEXT(8)   NOT NULL,
    SnapshotDate  DATETIME  NOT NULL,
    TenorLabel    TEXT(16),
    TenorYears    DOUBLE    NOT NULL,
    Rate          DOUBLE,
    [Status]      TEXT(32),
    CONSTRAINT PK_Curve_Point
        PRIMARY KEY (RunID, CCY, CurveType, SnapshotDate, TenorYears)
);

CREATE INDEX IX_Curve_Point_Lookup ON Curve_Point (RunID, CCY, CurveType);


-- --- processed facts --------------------------------------------------------

CREATE TABLE Fact_BondRisk (
    RunID            LONG      NOT NULL,
    ISIN             TEXT(12)  NOT NULL,
    Portfolio        TEXT(32)  NOT NULL,
    ModDur           DOUBLE,
    Convexity        DOUBLE,
    SpreadDuration   DOUBLE,
    DV01_EUR         DOUBLE,
    DV01_Opening_EUR DOUBLE,
    DirtyMV_T0_EUR   DOUBLE,
    DirtyMV_TM1_EUR  DOUBLE,
    CONSTRAINT PK_Fact_BondRisk PRIMARY KEY (RunID, ISIN, Portfolio)
);

-- One row per hedge.  This is the table the aggregation reads, and it is why
-- the 15 whole-column SUMIFS per bond row become a single GROUP BY.

CREATE TABLE Fact_HedgePosition (
    RunID               LONG      NOT NULL,
    HedgeKind           TEXT(4)   NOT NULL,
    HedgeID             TEXT(128) NOT NULL,
    LinkedISIN          TEXT(12),
    CoverageRelationID  LONG,
    LinkSource          TEXT(16),
    HedgeSource         TEXT(4),
    HedgeClass          TEXT(8),
    SwapIdSource        TEXT(16),
    Notional            DOUBLE,
    DV01_EUR            DOUBLE,
    PnL_EUR             DOUBLE,
    CONSTRAINT PK_Fact_HedgePosition PRIMARY KEY (RunID, HedgeKind, HedgeID)
);

CREATE INDEX IX_Fact_HedgePosition_Bond
    ON Fact_HedgePosition (RunID, LinkedISIN);

CREATE INDEX IX_Fact_HedgePosition_Relation
    ON Fact_HedgePosition (RunID, CoverageRelationID);

-- WIDE on purpose - see "Wide or tall" in docs/ACCESS_ARCHITECTURE.md.
--
-- The 86 measure fields are NOT written out here by hand.  They are generated
-- from the column registry by tools/gen_column_registry.py --ddl, so the table
-- and Meta_Column cannot drift apart.  This is the shell; the generator emits
-- the ALTER TABLE statements that add the measures.

CREATE TABLE Fact_PnlAttribution (
    RunID       LONG      NOT NULL,
    ISIN        TEXT(12)  NOT NULL,
    Portfolio   TEXT(32)  NOT NULL,
    CONSTRAINT PK_Fact_PnlAttribution PRIMARY KEY (RunID, ISIN, Portfolio)
);


-- --- meta: the layout registry ---------------------------------------------
-- The single source of truth for where a column sits on a sheet.
--
-- Ordinals go 10, 20, 30 - gaps on purpose, so inserting a column between two
-- others needs no renumbering.  Moving a column is one UPDATE here and nothing
-- else anywhere.

CREATE TABLE Meta_Column (
    ColumnID      AUTOINCREMENT  NOT NULL,
    TargetSheet   TEXT(32)       NOT NULL,
    FieldName     TEXT(64)       NOT NULL,
    DisplayLabel  TEXT(128),
    Ordinal       LONG           NOT NULL,
    NumberFormat  TEXT(32),
    ColumnWidth   DOUBLE,
    IsVisible     YESNO          NOT NULL,
    Notes         LONGTEXT,
    CONSTRAINT PK_Meta_Column PRIMARY KEY (ColumnID)
);

CREATE UNIQUE INDEX IX_Meta_Column_Field
    ON Meta_Column (TargetSheet, FieldName);

CREATE UNIQUE INDEX IX_Meta_Column_Ordinal
    ON Meta_Column (TargetSheet, Ordinal);
