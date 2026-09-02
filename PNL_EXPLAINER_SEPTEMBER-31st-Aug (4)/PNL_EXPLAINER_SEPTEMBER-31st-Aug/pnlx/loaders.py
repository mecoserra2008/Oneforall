"""CSV readers that turn files into books.

The workbook pulled positions from OPICS over ADO and market data from the
Bloomberg add-in.  Neither is reachable from Python here, and neither should
be: a PnL engine that can only run next to a terminal cannot be tested, cannot
be re-run for a past date, and cannot be checked by anyone without a licence.

The boundary is therefore a set of flat files.  Whatever produces them - an
Excel export of the existing sheets, a scheduled query, a market-data job - is
somebody else's problem, and the engine gets a reproducible input it can be
tested against.  `docs/MIGRATION.md` maps each column back to the sheet cell it
came from.

Reading rules, applied everywhere:

  * headers are matched case-insensitively and ignore spaces and underscores,
    so `Clean_Px_T0`, `clean px t0` and `cleanpxt0` are the same column;
  * a missing OPTIONAL column is fine and becomes all-NaN;
  * a missing REQUIRED column is a hard error naming the file and the column;
  * an unparseable number becomes NaN, never zero.
"""

from __future__ import annotations

import csv
import datetime as _dt
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping, Sequence

from .config import AppConfig
from .curves import CurveSet
from .instruments import (
    BondBook,
    BondPosition,
    FutureBook,
    FuturePosition,
    SwapBook,
    SwapPosition,
)

__all__ = [
    "LoadError",
    "InputBundle",
    "read_csv_rows",
    "load_bonds",
    "load_futures",
    "load_swaps",
    "load_curves",
    "load_fx",
    "load_spread_overrides",
    "load_inputs",
]


class LoadError(ValueError):
    """Raised when an input file is missing, malformed, or short a column."""


def _norm(name: str) -> str:
    return "".join(ch for ch in str(name).lower() if ch.isalnum())


def read_csv_rows(path: Path) -> list[dict[str, Any]]:
    """Read a CSV into dicts keyed on NORMALISED header names."""
    if not path.is_file():
        raise LoadError(f"input file not found: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise LoadError(f"{path}: no header row")
        mapping = {name: _norm(name) for name in reader.fieldnames if name is not None}
        rows: list[dict[str, Any]] = []
        for raw in reader:
            row: dict[str, Any] = {}
            for original, key in mapping.items():
                row[key] = raw.get(original)
            rows.append(row)
    return rows


class _Row:
    """One CSV row, with aliased column access and a required-column check."""

    __slots__ = ("_data", "_path", "_line")

    def __init__(self, data: Mapping[str, Any], path: Path, line: int) -> None:
        self._data = data
        self._path = path
        self._line = line

    def get(self, *aliases: str, default: Any = None) -> Any:
        for alias in aliases:
            key = _norm(alias)
            if key in self._data:
                value = self._data[key]
                if value is not None and str(value).strip() != "":
                    return value
        return default

    def require(self, *aliases: str) -> Any:
        value = self.get(*aliases)
        if value is None:
            raise LoadError(
                f"{self._path} line {self._line}: required column "
                f"{aliases[0]!r} is missing or blank"
            )
        return value

    def date(self, *aliases: str) -> _dt.date | None:
        value = self.get(*aliases)
        return _parse_date(value)


def _parse_date(value: Any) -> _dt.date | None:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    text = text.split("T")[0].split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y%m%d", "%d.%m.%Y"):
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _rows(path: Path) -> Iterator[_Row]:
    for offset, data in enumerate(read_csv_rows(path), start=2):
        yield _Row(data, path, offset)


# --------------------------------------------------------------------------- #
# bonds
# --------------------------------------------------------------------------- #


def load_bonds(path: Path) -> BondBook:
    """Positions and both market snapshots for the bond book.

    Column aliases keep the legacy sheet headers working, so an export of the
    existing Bonds sheet loads without being renamed first.
    """
    positions: list[BondPosition] = []
    for row in _rows(path):
        isin = str(row.require("isin")).strip().upper()
        if not isin:
            continue
        positions.append(
            BondPosition(
                isin=isin,
                name=str(row.get("name", "description", default="") or ""),
                currency=str(row.get("currency", "ccy", default="EUR") or "EUR"),
                portfolio=str(row.get("portfolio", "port", default="") or ""),
                acctg_cat=str(row.get("acctg_cat", "acctgcat", "invtype", default="") or ""),
                notional=row.get("notional", "qty"),
                coupon_pct=row.get("coupon_pct", "coupon", "coupon_rate", "couprate"),
                coupon_freq=row.get("coupon_freq", "couponfreq", "intpaycycle", "frequency"),
                maturity=row.date("maturity", "maturity_date", "mdate"),
                book_value=row.get("book_value", "bookval", "bookval_eur"),
                clean_px_prior=row.get("clean_px_prior", "cleanpx_tm1", "cleanpx_t_1"),
                dirty_px_prior=row.get("dirty_px_prior", "dirtypx_tm1", "dirtypx_t_1"),
                ytm_prior=row.get("ytm_prior", "ytm_tm1", "ytm_t_1"),
                zspread_prior=row.get("zspread_prior", "zsprd_tm1", "zsprd_t_1"),
                asw_prior=row.get("asw_prior", "asw_tm1", "asw_t_1"),
                oas_prior=row.get("oas_prior", "oas_tm1", "oas_t_1"),
                mod_duration_prior=row.get("mod_duration_prior", "moddur_tm1", "moddur_t_1"),
                clean_px_current=row.get("clean_px_current", "cleanpx_t0"),
                dirty_px_current=row.get("dirty_px_current", "dirtypx_t0"),
                ytm_current=row.get("ytm_current", "ytm_t0"),
                zspread_current=row.get("zspread_current", "zsprd_t0"),
                asw_current=row.get("asw_current", "asw_t0"),
                oas_current=row.get("oas_current", "oas_t0"),
                mod_duration_current=row.get("mod_duration_current", "moddur_t0"),
                oas_mod_duration_current=row.get(
                    "oas_mod_duration_current", "oas_modduration_raw"
                ),
                convexity=row.get("convexity"),
                fx_prior=row.get("fx_prior", "fx_tm1", "fx_t_1"),
                fx_current=row.get("fx_current", "fx_t0"),
                funding_rate_prior=row.get("funding_rate_prior", "fundrate_tm1"),
                funding_rate_current=row.get("funding_rate_current", "fundrate_t0"),
                day_count_desc=str(row.get("day_count_desc", "daycntdes", default="") or ""),
                day_count_code=row.get("day_count_code", "bonddcc_code"),
                pricing_source=str(row.get("pricing_source", default="") or ""),
                benchmark_bond=str(row.get("benchmark_bond", default="") or ""),
            )
        )
    if not positions:
        raise LoadError(f"{path}: no bond positions found")
    return BondBook.from_positions(positions)


# --------------------------------------------------------------------------- #
# futures
# --------------------------------------------------------------------------- #


def load_futures(path: Path) -> FutureBook:
    if not path.is_file():
        return FutureBook.from_positions([])
    positions: list[FuturePosition] = []
    for row in _rows(path):
        code = str(row.get("contract_code", "contractcode", default="") or "").strip()
        if not code:
            continue
        positions.append(
            FuturePosition(
                contract_code=code,
                exchange=str(row.get("exchange", default="") or ""),
                currency=str(row.get("currency", "ccy", default="EUR") or "EUR"),
                portfolio=str(row.get("portfolio", default="") or ""),
                linked_isin=str(row.get("linked_isin", "linkedisin", default="") or ""),
                hedge_type=str(row.get("hedge_type", "hedgetype", default="") or ""),
                contracts=row.get("contracts", "numcont"),
                face_value=row.get("face_value", "facevalue"),
                deliv_date=row.date("deliv_date", "delivdate", "delvdate"),
                ctd_isin=str(row.get("ctd_isin", default="") or ""),
                ctd_cf=row.get("ctd_cf", "convfactor", "conversion_factor"),
                ctd_dirty_px_current=row.get("ctd_dirty_px_current", "ctd_dirtypx_t0"),
                avg_entry_px=row.get("avg_entry_px", "avgentrypx"),
                fut_px_prior=row.get("fut_px_prior", "futpx_tm1", "futpx_t_1"),
                fut_px_current=row.get("fut_px_current", "futpx_t0"),
                fut_val_pt=row.get("fut_val_pt", "futvalpt", "point_value"),
                fut_px_val_bp=row.get(
                    "fut_px_val_bp", "futpxvalbp", "hedgeunitdv01", "unit_dv01"
                ),
                fx_prior=row.get("fx_prior", "fx_tm1"),
                fx_current=row.get("fx_current", "fx_t0"),
                implied_repo_bbg=row.get("implied_repo_bbg", "impliedrepo_bbg"),
                net_basis_bbg=row.get("net_basis_bbg", "netbasis_bbg"),
                gross_basis_bbg=row.get("gross_basis_bbg", "grossbasis_bbg"),
            )
        )
    return FutureBook.from_positions(positions)


# --------------------------------------------------------------------------- #
# swaps
# --------------------------------------------------------------------------- #


def load_swaps(path: Path) -> SwapBook:
    if not path.is_file():
        return SwapBook.from_positions([])
    positions: list[SwapPosition] = []
    for row in _rows(path):
        deal = str(row.get("deal_id", "dealid", "dealno", default="") or "").strip()
        if not deal:
            continue
        positions.append(
            SwapPosition(
                deal_id=deal,
                currency=str(row.get("currency", "ccy", default="EUR") or "EUR"),
                portfolio=str(row.get("portfolio", default="") or ""),
                linked_isin=str(row.get("linked_isin", "linkedisin", default="") or ""),
                counterparty=str(row.get("counterparty", "cpty", default="") or ""),
                swap_id_source=str(
                    row.get("swap_id_source", "swapidsource", "class", default="PLAIN")
                    or "PLAIN"
                ),
                notional=row.get("notional"),
                fixed_rate=row.get("fixed_rate", "fixedrate"),
                float_index=str(row.get("float_index", "floatindex", "payfltrateidx", default="") or ""),
                float_spread=row.get("float_spread", "floatspread"),
                start_date=row.date("start_date", "startdate"),
                end_date=row.date("end_date", "enddate", "matdate"),
                pay_fixed=str(row.get("pay_fixed", "payfixed", "netpayind", default="") or ""),
                float_curve_type=str(
                    row.get("float_curve_type", "floatcurvetype", "floatindexfamily", default="")
                    or ""
                ),
                dv01_supplied=row.get("dv01_supplied", "dv01_bbg", "dv01", "swcnvbpv"),
                npv_prior=row.get("npv_prior", "npv_tm1", "bql_npv_total_tm1"),
                npv_current=row.get("npv_current", "npv_t0", "bql_npv_total_t0"),
                fx_prior=row.get("fx_prior", "fx_tm1"),
                fx_current=row.get("fx_current", "fx_t0", "fx"),
            )
        )
    return SwapBook.from_positions(positions)


# --------------------------------------------------------------------------- #
# curves, fx, overrides
# --------------------------------------------------------------------------- #


def load_curves(path: Path) -> CurveSet:
    """Long-format curve nodes -> a `CurveSet`."""
    rows = read_csv_rows(path)
    records = []
    for row in rows:
        records.append(
            {
                "currency": row.get(_norm("currency")) or row.get(_norm("ccy")) or "",
                "curve_type": row.get(_norm("curve_type")) or row.get(_norm("curvetype")) or "",
                "tenor_years": row.get(_norm("tenor_years")) or row.get(_norm("years")),
                "rate_prior": row.get(_norm("rate_prior")) or row.get(_norm("rate_tm1")),
                "rate_current": row.get(_norm("rate_current")) or row.get(_norm("rate_t0")),
            }
        )
    curves = CurveSet.from_records(records)
    if len(curves) == 0:
        raise LoadError(f"{path}: no usable curves (each needs at least 2 nodes)")
    return curves


def load_fx(path: Path) -> dict[str, tuple[float, float]]:
    """`currency -> (fx_prior, fx_current)` in base per unit of currency.

    The base currency is pinned to 1.0 on both dates whatever the file says.
    A base-currency rate that is not exactly 1 would translate the entire
    domestic book, which is the kind of error that shows up as a large,
    perfectly explicable-looking FX line.
    """
    out: dict[str, tuple[float, float]] = {}
    if not path.is_file():
        return out
    for row in _rows(path):
        ccy = str(row.get("currency", "ccy", default="") or "").strip().upper()
        if not ccy:
            continue
        prior = _to_float(row.get("fx_prior", "fx_tm1"))
        current = _to_float(row.get("fx_current", "fx_t0"))
        out[ccy] = (prior, current)
    return out


def load_spread_overrides(path: Path) -> tuple[dict[str, str], dict[str, str]]:
    """`(isin -> framework, isin -> note)` from the manual overrides file."""
    codes: dict[str, str] = {}
    notes: dict[str, str] = {}
    if not path.is_file():
        return codes, notes
    for row in _rows(path):
        isin = str(row.get("isin", default="") or "").strip().upper()
        framework = str(row.get("framework", "spread_framework", default="") or "").strip().upper()
        if not isin or not framework:
            continue
        codes[isin] = framework
        note = str(row.get("reason", "note", default="") or "").strip()
        if note:
            notes[isin] = note
    return codes, notes


def _to_float(value: Any) -> float:
    if value is None:
        return float("nan")
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return float("nan")


# --------------------------------------------------------------------------- #
# bundle
# --------------------------------------------------------------------------- #


@dataclass(slots=True)
class InputBundle:
    """Everything one run needs, already parsed."""

    bonds: BondBook
    futures: FutureBook
    swaps: SwapBook
    curves: CurveSet
    fx: dict[str, tuple[float, float]]
    override_codes: dict[str, str]
    override_notes: dict[str, str]

    def summary(self) -> dict[str, int]:
        return {
            "bonds": len(self.bonds),
            "futures": len(self.futures),
            "swaps": len(self.swaps),
            "swaps_plain": int(self.swaps.is_plain.sum()) if len(self.swaps) else 0,
            "swaps_synthetic": int(self.swaps.is_synthetic.sum()) if len(self.swaps) else 0,
            "curves": len(self.curves),
            "fx_pairs": len(self.fx),
            "spread_overrides": len(self.override_codes),
        }


def load_inputs(config: AppConfig) -> InputBundle:
    """Read every input file named in the configuration."""
    paths = config.paths
    bonds = load_bonds(paths.input(paths.bonds))
    futures = load_futures(paths.input(paths.futures))
    swaps = load_swaps(paths.input(paths.swaps))
    curves = load_curves(paths.input(paths.curves))
    fx = load_fx(paths.input(paths.fx))
    codes, notes = load_spread_overrides(paths.input(paths.spread_overrides))

    fx.setdefault(config.run.base_ccy, (1.0, 1.0))
    fx[config.run.base_ccy] = (1.0, 1.0)

    return InputBundle(
        bonds=bonds,
        futures=futures,
        swaps=swaps,
        curves=curves,
        fx=fx,
        override_codes=codes,
        override_notes=notes,
    )
