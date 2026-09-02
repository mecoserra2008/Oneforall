-- Meta_Column: the provenance columns.
--
-- Apply after schema.sql, before seed_meta_column.sql.
-- Access executes one statement per Execute call and has no batch separator,
-- so each ALTER is its own statement.
--
-- WHY THESE FOUR
--
-- A field is not an "input" or an "output".  Bond_DV01_Opening is produced by
-- one step and read by twelve others; Spread_Framework_Auto is produced by one
-- and read by six.  A single input/output flag is wrong for most of the sheet,
-- and a schema built on one encodes the mistake permanently.
--
-- So the registry carries the edges and the ordering instead:
--
--   SourceKind      how the value arrives - QUERY / BLOOMBERG / DERIVED / UDF /
--                   AGGREGATE / UNWRITTEN.  Only the first two cross a network.
--   ProducedBy      the procedure (today) or the query (after the move) that
--                   writes it.
--   ProducedInWave  topological rank.  Nothing in wave N reads wave N or later,
--                   so wave N may not start until wave N-1 has been checked.
--   ConsumedBy      the fields that read this one, comma-separated.
--
-- "Is this an input?" is then answered per consumer against ConsumedBy, not
-- guessed from a name.
--
-- GENERATED alongside the seed by tools/gen_column_registry.py, which takes the
-- wave and the edges from tools/field_graph.py.

ALTER TABLE Meta_Column ADD COLUMN SourceKind TEXT(16);
ALTER TABLE Meta_Column ADD COLUMN ProducedBy TEXT(64);
ALTER TABLE Meta_Column ADD COLUMN ProducedInWave LONG;
ALTER TABLE Meta_Column ADD COLUMN ConsumedBy LONGTEXT;

CREATE INDEX IX_Meta_Column_Wave ON Meta_Column (TargetSheet, ProducedInWave);
