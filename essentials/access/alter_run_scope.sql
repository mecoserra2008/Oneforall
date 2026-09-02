-- Run scope.
--
-- Apply after schema.sql. One statement per Execute; Access has no batch separator.
--
-- WHY
--
-- The bond extraction does not fetch the book, it fetches one slice: branch 01,
-- portfolio PORT, accounting category A, products OBR and SECUR. Only the
-- category survives into a column (Bonds!I, constant 'A'); the rest exist solely
-- as a WHERE clause and are invisible in every sheet and every stored table.
--
-- So [Run] as it stands records dates, counts, user, version and status - and
-- nothing about WHAT THE RUN COVERS. Two runs over different slices look the
-- same. Widening the category filter next month would produce a run
-- indistinguishable from today's that means something else entirely.
--
-- These four columns make the slice a property of the run. Any comparison across
-- runs - a time series, a variance, a Dashboard trend - must match on
-- ScopePortfolio and ScopeAcctgCat, which is the pair the Portfolio dimension
-- table (schema.sql:71) is already keyed on.

ALTER TABLE [Run] ADD COLUMN ScopeBranch TEXT(8);
ALTER TABLE [Run] ADD COLUMN ScopePortfolio TEXT(32);
ALTER TABLE [Run] ADD COLUMN ScopeAcctgCat TEXT(32);
ALTER TABLE [Run] ADD COLUMN ScopeProducts TEXT(64);

CREATE INDEX IX_Run_Scope ON [Run] (ScopePortfolio, ScopeAcctgCat, AsOfT0);
