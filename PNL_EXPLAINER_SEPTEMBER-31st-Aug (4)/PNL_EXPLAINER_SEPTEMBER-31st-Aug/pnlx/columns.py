"""Self-describing column frames.

Every result in this package is a set of parallel numpy arrays plus a SCHEMA
that says, for each column, what it is called, what unit it is in, what it means
economically, and the formula that produced it.

One schema, three consumers:

    the CSV writer      column order and header names
    the Excel writer    number formats, widths, conditional formatting
    the documentation   the "Formulas and Economics" sheet is generated from
                        these entries, so it cannot drift from the code

That last point is the reason this module exists rather than a plain dict of
arrays.  In the legacy workbook the economic meaning of a column lived in a VBA
comment, the number format lived in a formatting routine, and the header text
lived in a third place - and the Dashboard's own header contract had already
drifted out of step with the sheet it read.  Here there is one place to change.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable, Iterator, Mapping, Sequence

import numpy as np

__all__ = ["Column", "ColumnFrame", "Kind"]


class Kind:
    """Column unit tags.  The Excel writer maps these onto number formats."""

    TEXT = "text"
    DATE = "date"
    EUR = "eur"          # base-currency amount
    BP = "bp"            # basis points
    PCT = "pct"          # a fraction, displayed as a percentage
    RATIO = "ratio"      # a pure ratio, displayed as a number
    RATE = "rate"        # a rate in percentage points (3.25 means 3.25%)
    PRICE = "price"      # per 100 nominal
    NUM = "num"          # a plain number
    COUNT = "count"      # an integer count
    NOTIONAL = "notional"
    FLAG = "flag"        # boolean


@dataclass(frozen=True, slots=True)
class Column:
    """One output column, with its unit and its economic meaning."""

    name: str
    label: str
    kind: str
    meaning: str
    formula: str = ""
    group: str = ""

    @property
    def is_numeric(self) -> bool:
        return self.kind not in (Kind.TEXT, Kind.DATE, Kind.FLAG)


@dataclass(slots=True)
class ColumnFrame:
    """Parallel arrays addressed by name, described by a schema."""

    schema: tuple[Column, ...]
    data: dict[str, np.ndarray] = field(default_factory=dict)

    # -- construction ------------------------------------------------------- #

    @classmethod
    def build(
        cls, schema: Sequence[Column], data: Mapping[str, np.ndarray]
    ) -> "ColumnFrame":
        """Assemble, checking that the schema and the data agree.

        A column in the schema with no array, or an array whose length differs
        from the others, is a programming error and is raised as one.  A silent
        length mismatch would misalign every row of the report against every
        other, which is not detectable by eye.
        """
        schema = tuple(schema)
        missing = [c.name for c in schema if c.name not in data]
        if missing:
            raise KeyError(f"schema columns with no data: {missing}")

        lengths = {c.name: np.asarray(data[c.name]).shape[0] for c in schema}
        distinct = set(lengths.values())
        if len(distinct) > 1:
            raise ValueError(f"columns of differing length: {lengths}")

        return cls(schema=schema, data={c.name: np.asarray(data[c.name]) for c in schema})

    # -- access ------------------------------------------------------------- #

    @property
    def names(self) -> tuple[str, ...]:
        return tuple(c.name for c in self.schema)

    @property
    def labels(self) -> tuple[str, ...]:
        return tuple(c.label for c in self.schema)

    def column(self, name: str) -> Column:
        for c in self.schema:
            if c.name == name:
                return c
        raise KeyError(name)

    def __len__(self) -> int:
        if not self.data:
            return 0
        first = next(iter(self.data.values()))
        return int(np.asarray(first).shape[0])

    def __contains__(self, name: object) -> bool:
        return name in self.data

    def __getitem__(self, name: str) -> np.ndarray:
        return self.data[name]

    def get(self, name: str, default: np.ndarray | None = None) -> np.ndarray:
        if name in self.data:
            return self.data[name]
        if default is not None:
            return default
        return np.full(len(self), np.nan)

    def __iter__(self) -> Iterator[Column]:
        return iter(self.schema)

    # -- views -------------------------------------------------------------- #

    def to_records(self) -> list[dict[str, Any]]:
        """Row dicts, in schema order, with numpy scalars unwrapped."""
        n = len(self)
        cols = [(c.name, self.data[c.name]) for c in self.schema]
        out: list[dict[str, Any]] = []
        for i in range(n):
            row: dict[str, Any] = {}
            for name, arr in cols:
                row[name] = _scalar(arr[i])
            out.append(row)
        return out

    def rows(self) -> Iterable[list[Any]]:
        """Row lists, for writers that want positional values."""
        n = len(self)
        cols = [self.data[c.name] for c in self.schema]
        for i in range(n):
            yield [_scalar(arr[i]) for arr in cols]

    def total(self, name: str) -> float:
        """NaN-tolerant sum of one numeric column.

        NaN means "not measured", not zero, so it is skipped rather than
        poisoning the total.  How many rows were skipped is reported separately
        by the data-quality block - a total that quietly covers 80% of the book
        is the failure mode this guards against.
        """
        arr = np.asarray(self.get(name), dtype=np.float64)
        return float(np.nansum(arr)) if arr.size else 0.0

    def totals(self, names: Iterable[str]) -> dict[str, float]:
        return {name: self.total(name) for name in names}

    def count_finite(self, name: str) -> int:
        arr = np.asarray(self.get(name), dtype=np.float64)
        return int(np.isfinite(arr).sum())

    def filtered(self, mask: np.ndarray) -> "ColumnFrame":
        """A new frame holding only the rows where `mask` is True."""
        mask = np.asarray(mask, dtype=bool)
        return ColumnFrame(
            schema=self.schema,
            data={name: arr[mask] for name, arr in self.data.items()},
        )

    def sorted_by(self, name: str, *, descending: bool = True, limit: int | None = None) -> "ColumnFrame":
        """Rows ordered by one numeric column, NaNs last."""
        arr = np.asarray(self.get(name), dtype=np.float64)
        keys = np.where(np.isfinite(arr), arr, -np.inf if descending else np.inf)
        order = np.argsort(keys, kind="stable")
        if descending:
            order = order[::-1]
        if limit is not None:
            order = order[:limit]
        return ColumnFrame(
            schema=self.schema,
            data={n: a[order] for n, a in self.data.items()},
        )

    def describe_schema(self) -> list[dict[str, str]]:
        """The schema itself, as rows - this is the documentation sheet."""
        return [
            {
                "group": c.group,
                "column": c.name,
                "label": c.label,
                "unit": c.kind,
                "meaning": c.meaning,
                "formula": c.formula,
            }
            for c in self.schema
        ]


def _scalar(value: Any) -> Any:
    """Unwrap a numpy scalar into a plain Python value."""
    if isinstance(value, np.datetime64):
        if np.isnat(value):
            return None
        return value.astype("datetime64[D]").astype(object)
    if isinstance(value, (np.floating,)):
        f = float(value)
        return None if np.isnan(f) else f
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, float) and value != value:
        return None
    return value
