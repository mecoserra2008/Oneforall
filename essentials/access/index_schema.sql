-- The run index.
--
-- One row per run database.  Applied to PNL_Runs_Index.accdb, which sits in the
-- same folder as the run databases themselves.
--
--     python3 tools/build_access_db.py --check
--
-- WHY THIS EXISTS
--
-- Runs are stored one database per file, named for the retrieval timestamp.
-- That makes each run immutable and keeps any single file far away from the 2 GB
-- ceiling, but it means "every run in September" is an attach-and-union across
-- files rather than one query - and a folder listing is a poor substitute,
-- because the filename carries the retrieval time and nothing else.
--
-- This table carries the rest: which as-of dates the run covered, whether it
-- finished, how many positions it held and what its fingerprint was.  So the
-- question "which files do I need to open" is answered here, once, before any
-- of them are opened.
--
-- DbPath is stored as written, not resolved.  A run database that has been moved
-- or archived should show as missing rather than be silently re-pointed.

CREATE TABLE Run_Index (
    RunIndexID      AUTOINCREMENT  NOT NULL,
    RunStamp        TEXT(15)       NOT NULL,   -- yyyymmdd_hhnnss, the retrieval moment
    DbPath          TEXT(255)      NOT NULL,
    AsOfT0          DATETIME,
    AsOfT1          DATETIME,
    RunStartedUtc   DATETIME,
    RunFinishedUtc  DATETIME,
    RunStatus       TEXT(16),
    RunFingerprint  TEXT(32),
    BondCount       LONG,
    SwapCount       LONG,
    FutureCount     LONG,
    MachineName     TEXT(64),
    UserName        TEXT(64),
    ScopeBranch     TEXT(8),
    ScopePortfolio  TEXT(32),
    ScopeAcctgCat   TEXT(32),
    ScopeProducts   TEXT(64),
    Notes           LONGTEXT,
    CONSTRAINT PK_Run_Index PRIMARY KEY (RunIndexID)
);

-- One row per retrieval.  A re-save of the same pull updates its row rather than
-- adding a second - the same rule the run databases themselves follow.
CREATE UNIQUE INDEX IX_Run_Index_Stamp ON Run_Index (RunStamp);

CREATE INDEX IX_Run_Index_AsOf ON Run_Index (AsOfT0);

-- Find every run over one slice without opening a single run database.
CREATE INDEX IX_Run_Index_Scope ON Run_Index (ScopePortfolio, ScopeAcctgCat, AsOfT0);
