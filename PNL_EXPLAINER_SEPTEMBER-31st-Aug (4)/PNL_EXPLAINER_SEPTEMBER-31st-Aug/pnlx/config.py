"""Typed configuration.

Everything the model can be steered by lives in one YAML file and arrives here
as frozen dataclasses.  Nothing reads the raw dict after `load_config` returns:
a typo in the YAML fails at load time with the key that was wrong, rather than
silently taking a default three modules later.

The VBA equivalent was the Config sheet (cells B4..B50).  The mapping from those
cells to these fields is in docs/MIGRATION.md.
"""

from __future__ import annotations

import datetime as _dt
import os
from dataclasses import dataclass, field, fields, is_dataclass
from pathlib import Path
from typing import Any, Mapping

import yaml

__all__ = [
    "AppConfig",
    "RunConfig",
    "PathsConfig",
    "ModelConfig",
    "PullToParConfig",
    "FuturesConfig",
    "SwapsConfig",
    "Tolerances",
    "OutputConfig",
    "ConfigError",
    "load_config",
    "default_config_path",
]


class ConfigError(ValueError):
    """Raised when the YAML cannot be turned into a valid configuration."""


# --------------------------------------------------------------------------- #
# enumerated choices
# --------------------------------------------------------------------------- #

RISK_DATES = ("prior", "current")
COUPON_CARRY_METHODS = ("exact", "smooth")
CONVEXITY_SOURCES = ("bump", "analytic", "supplied")
SWAP_DV01_SOURCES = ("supplied", "annuity")
FRAMEWORK_CODES = ("G", "I", "ASW", "Z", "OAS", "OIS", "SOFR", "MIXED", "REVIEW")


# --------------------------------------------------------------------------- #
# sections
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class RunConfig:
    """Which two dates are being compared, and in what currency.

    Naming here is the REPORT convention, not the legacy VBA one:

        prior   T-1, the opening snapshot
        as_of   T0,  the closing / reporting snapshot

    The VBA carried these the other way round internally (`CFG_T0_DATE` = B4 was
    the prior date) which is the single easiest way to flip the sign of every
    delta in the book.  There is no such ambiguity left.
    """

    as_of: _dt.date
    prior: _dt.date
    base_ccy: str = "EUR"
    label: str = "daily"
    book: str = "FIXED_INCOME"

    def __post_init__(self) -> None:
        if self.as_of <= self.prior:
            raise ConfigError(
                f"run.as_of ({self.as_of}) must be strictly after run.prior "
                f"({self.prior}); a zero or negative period explains nothing."
            )
        if not self.base_ccy or len(self.base_ccy) != 3:
            raise ConfigError(f"run.base_ccy must be a 3-letter code, got {self.base_ccy!r}")

    @property
    def days(self) -> int:
        """Calendar days in the period being explained."""
        return (self.as_of - self.prior).days

    @property
    def year_fraction(self) -> float:
        """ACT/365F year fraction of the period - the carry accrual axis."""
        return self.days / 365.0


@dataclass(frozen=True, slots=True)
class PathsConfig:
    """Where the inputs are read from and the outputs written to."""

    input_dir: Path = Path("data/sample")
    output_dir: Path = Path("out")
    bonds: str = "positions_bonds.csv"
    futures: str = "positions_futures.csv"
    swaps: str = "positions_swaps.csv"
    curves: str = "curves.csv"
    fx: str = "fx.csv"
    spread_overrides: str = "spread_overrides.csv"
    history: str = "history/daily_pnl_history.csv"

    def input(self, name: str) -> Path:
        return self.input_dir / name

    def output(self, name: str) -> Path:
        return self.output_dir / name


@dataclass(frozen=True, slots=True)
class PullToParConfig:
    """Pull to par, a.k.a. roll to par.

    `enabled=False` books zero rather than blank, so the bridge still closes for
    a desk that reports carry as coupon only.
    """

    enabled: bool = True
    day_basis: float = 365.0
    max_cashflows: int = 1200

    def __post_init__(self) -> None:
        if self.day_basis <= 0:
            raise ConfigError("model.pull_to_par.day_basis must be positive")
        if self.max_cashflows < 1:
            raise ConfigError("model.pull_to_par.max_cashflows must be at least 1")


@dataclass(frozen=True, slots=True)
class ModelConfig:
    """The economic choices.  Each one changes a number on the report."""

    spread_framework_override: str = ""
    attribution_risk_date: str = "prior"
    coupon_carry: str = "exact"
    convexity_source: str = "bump"
    convexity_bump_bp: float = 100.0
    include_funding_in_bridge: bool = False
    fx_cross_term: bool = True
    pull_to_par: PullToParConfig = field(default_factory=PullToParConfig)

    def __post_init__(self) -> None:
        ovr = self.spread_framework_override.strip().upper()
        if ovr and ovr not in FRAMEWORK_CODES:
            raise ConfigError(
                f"model.spread_framework_override {ovr!r} is not one of {FRAMEWORK_CODES}"
            )
        object.__setattr__(self, "spread_framework_override", ovr)

        if self.attribution_risk_date not in RISK_DATES:
            raise ConfigError(
                f"model.attribution_risk_date must be one of {RISK_DATES}, "
                f"got {self.attribution_risk_date!r}"
            )
        if self.coupon_carry not in COUPON_CARRY_METHODS:
            raise ConfigError(
                f"model.coupon_carry must be one of {COUPON_CARRY_METHODS}, "
                f"got {self.coupon_carry!r}"
            )
        if self.convexity_source not in CONVEXITY_SOURCES:
            raise ConfigError(
                f"model.convexity_source must be one of {CONVEXITY_SOURCES}, "
                f"got {self.convexity_source!r}"
            )
        if self.convexity_bump_bp <= 0:
            raise ConfigError("model.convexity_bump_bp must be positive")


@dataclass(frozen=True, slots=True)
class FuturesConfig:
    """Futures conventions.

    `dv01_includes_point_value` is the one that has bitten this book before.
    Bloomberg's FUT_PX_VAL_BP is quoted in PRICE POINTS per basis point (about
    0.0605 for a Bund), so the cash DV01 per contract needs a further multiply
    by FUT_VAL_PT (EUR 1000 per full point) - roughly EUR 60 per contract per bp.

    Desk check: Futures_DV01_EUR / Contracts should land in the 60-90 range for
    a Bund.  A figure near 60,000 means the feed is already in cash and this
    flag should be True.
    """

    dv01_includes_point_value: bool = False
    treat_missing_ctd_as_error: bool = False


@dataclass(frozen=True, slots=True)
class SwapsConfig:
    """Swap conventions.

    `dv01_source='supplied'` uses the risk the swap system publishes
    (Bloomberg SW_CNV_BPV in the legacy sheet).  `'annuity'` falls back to the
    internal annuity model, which is only a sanity check - it assumes a single
    bullet annuity rather than the real payment schedule.
    """

    dv01_source: str = "supplied"
    unknown_family_is_error: bool = True

    def __post_init__(self) -> None:
        if self.dv01_source not in SWAP_DV01_SOURCES:
            raise ConfigError(
                f"swaps.dv01_source must be one of {SWAP_DV01_SOURCES}, "
                f"got {self.dv01_source!r}"
            )


@dataclass(frozen=True, slots=True)
class Tolerances:
    """Thresholds that turn a number into a status.

    hedge_ratio    a ratio of exactly 1.00 is a perfect hedge, so "over-hedged"
                   must not fire on rounding.  2% of the bond DV01.
    residual_pct   share of practical PnL above which the residual is worth a
                   human look.
    residual_eur   absolute floor for the above.  A percentage test on its own
                   flags a 24 EUR residual on a 65 EUR position and buries the
                   real breaks under it; a residual has to be both a large
                   share AND a real amount to be worth anyone's morning.
    identity_pct   share of the duration total above which the framework chain
                   is treated as not tying.
    identity_eur   absolute floor for the above, same reasoning.
    """

    hedge_ratio: float = 0.02
    residual_pct: float = 0.10
    residual_eur: float = 250.0
    identity_pct: float = 0.01
    identity_eur: float = 1.0

    def __post_init__(self) -> None:
        for f in fields(self):
            if getattr(self, f.name) < 0:
                raise ConfigError(f"tolerances.{f.name} must not be negative")


@dataclass(frozen=True, slots=True)
class OutputConfig:
    """Deliverables."""

    excel: str = "pnl_explainer.xlsx"
    csv: str = "daily_pnl_explained.csv"
    json: str = "daily_pnl_metadata.json"
    bonds_csv: str = "daily_pnl_bonds.csv"
    futures_csv: str = "daily_pnl_futures.csv"
    swaps_csv: str = "daily_pnl_swaps.csv"
    top_n: int = 15
    append_history: bool = True
    date_stamp_filenames: bool = False
    float_format: str = "%.6f"

    def __post_init__(self) -> None:
        if self.top_n < 1:
            raise ConfigError("output.top_n must be at least 1")


@dataclass(frozen=True, slots=True)
class AppConfig:
    """The whole configuration, as one object passed down the call chain."""

    run: RunConfig
    paths: PathsConfig = field(default_factory=PathsConfig)
    model: ModelConfig = field(default_factory=ModelConfig)
    futures: FuturesConfig = field(default_factory=FuturesConfig)
    swaps: SwapsConfig = field(default_factory=SwapsConfig)
    tolerances: Tolerances = field(default_factory=Tolerances)
    output: OutputConfig = field(default_factory=OutputConfig)
    source_path: Path | None = None

    def to_dict(self) -> dict[str, Any]:
        """Plain-data view, for the JSON metadata sidecar."""
        return _as_plain(self)


# --------------------------------------------------------------------------- #
# loading
# --------------------------------------------------------------------------- #


def default_config_path() -> Path:
    """Repository-relative default, so `python -m pnlx` works with no arguments."""
    return Path(__file__).resolve().parent.parent / "config" / "pnl_explainer.yaml"


def load_config(path: str | os.PathLike[str] | None = None) -> AppConfig:
    """Read the YAML file and build a validated `AppConfig`.

    Unknown keys are a hard error.  A configuration file whose typo is silently
    ignored is worse than one that will not load: the run succeeds and reports
    the wrong number.
    """
    cfg_path = Path(path) if path is not None else default_config_path()
    if not cfg_path.is_file():
        raise ConfigError(f"configuration file not found: {cfg_path}")

    with cfg_path.open("r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    if not isinstance(raw, Mapping):
        raise ConfigError(f"{cfg_path}: top level must be a mapping")

    known = {"run", "paths", "model", "futures", "swaps", "tolerances", "output"}
    _reject_unknown(raw, known, "<root>")

    run = _build(RunConfig, raw.get("run"), "run", required=True, coercers={
        "as_of": _as_date,
        "prior": _as_date,
        "base_ccy": lambda v: str(v).strip().upper(),
    })

    base_dir = cfg_path.resolve().parent.parent
    paths = _build(PathsConfig, raw.get("paths"), "paths", coercers={
        "input_dir": lambda v: _resolve(base_dir, v),
        "output_dir": lambda v: _resolve(base_dir, v),
    })

    model_raw = dict(raw.get("model") or {})
    ptp_raw = model_raw.pop("pull_to_par", None)
    model_kwargs: dict[str, Any] = {}
    if ptp_raw is not None:
        model_kwargs["pull_to_par"] = _build(
            PullToParConfig, ptp_raw, "model.pull_to_par"
        )
    model = _build(ModelConfig, model_raw, "model", extra=model_kwargs)

    futures = _build(FuturesConfig, raw.get("futures"), "futures")
    swaps = _build(SwapsConfig, raw.get("swaps"), "swaps")
    tolerances = _build(Tolerances, raw.get("tolerances"), "tolerances")
    output = _build(OutputConfig, raw.get("output"), "output")

    return AppConfig(
        run=run,
        paths=paths,
        model=model,
        futures=futures,
        swaps=swaps,
        tolerances=tolerances,
        output=output,
        source_path=cfg_path,
    )


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #


def _resolve(base: Path, value: Any) -> Path:
    p = Path(str(value)).expanduser()
    return p if p.is_absolute() else (base / p)


def _as_date(value: Any) -> _dt.date:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y%m%d"):
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ConfigError(f"cannot read {value!r} as a date (use YYYY-MM-DD)")


def _reject_unknown(mapping: Mapping[str, Any], known: set[str], where: str) -> None:
    unknown = sorted(set(mapping) - known)
    if unknown:
        raise ConfigError(
            f"{where}: unknown key(s) {unknown}; expected one of {sorted(known)}"
        )


def _build(
    cls: type,
    raw: Any,
    where: str,
    *,
    required: bool = False,
    coercers: Mapping[str, Any] | None = None,
    extra: Mapping[str, Any] | None = None,
):
    """Instantiate a section dataclass from a mapping, validating keys."""
    if raw is None:
        if required:
            raise ConfigError(f"{where}: section is required")
        raw = {}
    if not isinstance(raw, Mapping):
        raise ConfigError(f"{where}: must be a mapping")

    names = {f.name for f in fields(cls)}
    _reject_unknown(raw, names, where)

    kwargs: dict[str, Any] = {}
    for key, value in raw.items():
        if coercers and key in coercers:
            kwargs[key] = coercers[key](value)
        else:
            kwargs[key] = value
    if extra:
        kwargs.update(extra)

    try:
        return cls(**kwargs)
    except ConfigError:
        raise
    except TypeError as exc:  # pragma: no cover - defensive
        raise ConfigError(f"{where}: {exc}") from exc


def _as_plain(obj: Any) -> Any:
    if is_dataclass(obj) and not isinstance(obj, type):
        return {f.name: _as_plain(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, (_dt.date, _dt.datetime)):
        return obj.isoformat()
    if isinstance(obj, (list, tuple)):
        return [_as_plain(v) for v in obj]
    if isinstance(obj, Mapping):
        return {str(k): _as_plain(v) for k, v in obj.items()}
    return obj
