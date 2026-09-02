"""Deliverables: CSV, JSON metadata and the Excel workbook.

Three outputs, three audiences:

    CSV     the daily PnL explained, one row per position, plus an appended
            history file that the snail trail reads back.  Machine-readable,
            diff-able, and the format anything downstream can consume.
    JSON    the metadata: what was run, on which dates, with which conventions,
            what it totalled, and what the data quality was.  Enough to
            reproduce or to audit a figure months later.
    Excel   the report a human opens.
"""

from .csv_out import write_csv_outputs, append_history
from .json_out import write_metadata, build_metadata
from .excel import write_workbook

__all__ = [
    "write_csv_outputs",
    "append_history",
    "write_metadata",
    "build_metadata",
    "write_workbook",
]
