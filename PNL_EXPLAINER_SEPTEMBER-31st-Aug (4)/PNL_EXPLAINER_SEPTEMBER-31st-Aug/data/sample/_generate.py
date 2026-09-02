"""Build the synthetic sample book.

The sample exists so the pipeline can be run, tested and demonstrated without a
market-data terminal.  It is generated rather than hand-typed for one reason:
the prices have to be INTERNALLY CONSISTENT with the curves in `curves.csv`.

Each bond is priced off its own base curve plus a spread, on both dates, using
the same `bondmath` the engine uses.  A yield move therefore really is the curve
move plus the spread move, so the attribution bridge closes to model error and
nothing else - which is what makes the sample useful as a regression fixture.

Deliberate imperfections are seeded, because a book where everything ties tests
nothing.  See `_IMPERFECTIONS` at the bottom.

Run:  python data/sample/_generate.py
"""

from __future__ import annotations

import csv
import datetime as dt
import sys
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from pnlx import bondmath, daycount as dc  # noqa: E402
from pnlx.curves import CurveSet  # noqa: E402
from pnlx.loaders import load_curves  # noqa: E402

HERE = Path(__file__).resolve().parent
PRIOR = dt.date(2025, 9, 29)
CURRENT = dt.date(2025, 9, 30)

# isin, name, ccy, coupon %, freq, maturity, notional, portfolio, acctg, dcc,
# base curve, spread bp at T-1, spread move bp
BONDS = [
    ("DE0001102614", "BUNDESREPUB. DEUTSCHLAND 2.3 02/15/33", "EUR", 2.30, 1, "2033-02-15", 45_000_000, "GOVIES", "HTC", 0, "GOV", 2.0, -0.4),
    ("DE0001102390", "BUNDESREPUB. DEUTSCHLAND 0.0 08/15/30", "EUR", 0.00, 1, "2030-08-15", 30_000_000, "GOVIES", "HTC", 0, "GOV", 1.5, -0.2),
    ("FR0014009O62", "FRANCE O.A.T. 2.0 11/25/32", "EUR", 2.00, 1, "2032-11-25", 38_000_000, "GOVIES", "HTC", 0, "GOV", 52.0, 1.8),
    ("FR0013480613", "FRANCE O.A.T. 0.5 05/25/40", "EUR", 0.50, 1, "2040-05-25", 15_000_000, "GOVIES", "AFS", 0, "GOV", 71.0, 2.6),
    ("IT0005534141", "BUONI POLIENNALI DEL TES 4.35 11/01/33", "EUR", 4.35, 2, "2033-11-01", 42_000_000, "GOVIES", "AFS", 0, "GOV", 118.0, -3.5),
    ("IT0005390874", "BUONI POLIENNALI DEL TES 1.45 03/01/36", "EUR", 1.45, 2, "2036-03-01", 22_000_000, "GOVIES", "AFS", 0, "GOV", 141.0, -4.2),
    ("ES0000012L52", "SPANISH GOVERNMENT 3.15 04/30/33", "EUR", 3.15, 1, "2033-04-30", 33_000_000, "GOVIES", "HTC", 0, "GOV", 63.0, -1.1),
    ("ES0000012J07", "SPANISH GOVERNMENT 1.0 10/31/50", "EUR", 1.00, 1, "2050-10-31", 8_000_000, "GOVIES", "AFS", 0, "GOV", 96.0, 3.1),
    ("PTOTEQOE0024", "PORTUGUESE OTS 3.5 06/18/38", "EUR", 3.50, 1, "2038-06-18", 26_000_000, "GOVIES", "HTC", 0, "GOV", 74.0, -2.0),
    ("PTOTEUOE0021", "PORTUGUESE OTS 2.25 04/18/34", "EUR", 2.25, 1, "2034-04-18", 31_000_000, "GOVIES", "HTC", 0, "GOV", 58.0, -1.4),
    ("BE0000359333", "BELGIUM KINGDOM 3.0 06/22/33", "EUR", 3.00, 1, "2033-06-22", 18_000_000, "GOVIES", "AFS", 0, "GOV", 61.0, 0.5),
    ("NL0015001AH9", "NETHERLANDS GOVERNMENT 2.5 07/15/33", "EUR", 2.50, 1, "2033-07-15", 20_000_000, "GOVIES", "HTC", 0, "GOV", 15.0, -0.3),
    ("XS2434891219", "SANTANDER 1.75 02/17/32", "EUR", 1.75, 1, "2032-02-17", 12_000_000, "CREDIT", "AFS", 5, "SWAP", 128.0, -5.5),
    ("XS2364001078", "BNP PARIBAS 1.25 07/13/31", "EUR", 1.25, 1, "2031-07-13", 10_000_000, "CREDIT", "AFS", 5, "SWAP", 141.0, -6.2),
    ("XS2600310275", "UNICREDIT 4.8 02/23/32", "EUR", 4.80, 1, "2032-02-23", 9_000_000, "CREDIT", "AFS", 5, "SWAP", 186.0, -8.4),
    ("XS2542168951", "ING GROEP 4.5 01/09/30", "EUR", 4.50, 1, "2030-01-09", 11_000_000, "CREDIT", "AFS", 5, "SWAP", 112.0, -4.0),
    ("XS2333563817", "TELEFONICA 1.807 04/15/31", "EUR", 1.807, 1, "2031-04-15", 7_500_000, "CREDIT", "AFS", 5, "SWAP", 154.0, -3.3),
    ("XS2445369418", "ENEL FINANCE 3.875 07/09/34", "EUR", 3.875, 1, "2034-07-09", 8_500_000, "CREDIT", "AFS", 5, "SWAP", 132.0, -2.7),
    ("XS2010044269", "IBERDROLA 1.875 01/28/32", "EUR", 1.875, 1, "2032-01-28", 6_000_000, "CREDIT", "AFS", 5, "SWAP", 119.0, -2.2),
    ("XS2588103241", "EDP FINANCE 4.496 04/30/30", "EUR", 4.496, 1, "2030-04-30", 9_500_000, "CREDIT", "AFS", 5, "SWAP", 145.0, -5.1),
    ("XS2513282538", "VOLKSWAGEN INTL FIN 4.375 11/24/31", "EUR", 4.375, 1, "2031-11-24", 7_000_000, "CREDIT", "AFS", 5, "SWAP", 168.0, -6.8),
    ("XS2244278008", "DEUTSCHE TELEKOM 0.5 07/05/32", "EUR", 0.50, 1, "2032-07-05", 5_500_000, "CREDIT", "AFS", 5, "SWAP", 108.0, -1.9),
    ("XS2453880566", "REPSOL 3.75 02/03/33", "EUR", 3.75, 1, "2033-02-03", 6_500_000, "CREDIT", "AFS", 5, "SWAP", 151.0, -4.6),
    ("XS2439096747", "AIRBUS 3.625 05/22/34", "EUR", 3.625, 1, "2034-05-22", 5_000_000, "CREDIT", "AFS", 5, "SWAP", 97.0, -1.5),
    ("XS2679523674", "NESTLE FINANCE 3.5 09/12/33", "EUR", 3.50, 1, "2033-09-12", 4_500_000, "CREDIT", "AFS", 5, "SWAP", 64.0, -0.9),
    ("US91282CJL65", "US TREASURY N/B 4.25 11/15/34", "USD", 4.25, 2, "2034-11-15", 25_000_000, "USD_RATES", "AFS", 0, "GOV", 3.0, -0.6),
    ("US912810UB50", "US TREASURY N/B 4.625 05/15/54", "USD", 4.625, 2, "2054-05-15", 9_000_000, "USD_RATES", "AFS", 0, "GOV", 5.0, 1.2),
    ("US912828YS31", "US TREASURY N/B 1.75 11/15/29", "USD", 1.75, 2, "2029-11-15", 14_000_000, "USD_RATES", "HTC", 0, "GOV", 2.0, -0.4),
    ("US023135CF19", "AMAZON.COM INC 4.65 12/01/29", "USD", 4.65, 2, "2029-12-01", 6_000_000, "USD_CREDIT", "AFS", 5, "SWAP", 62.0, -2.4),
    ("US037833EK43", "APPLE INC 4.0 05/10/28", "USD", 4.00, 2, "2028-05-10", 5_000_000, "USD_CREDIT", "AFS", 5, "SWAP", 48.0, -1.6),
    # short-dated, so the front of the curve is exercised
    ("DE0001141851", "BUNDESSCHATZANWEISUNGEN 2.2 04/12/27", "EUR", 2.20, 1, "2027-04-12", 28_000_000, "GOVIES", "HTC", 0, "GOV", 3.0, -0.5),
    ("FR0127792055", "FRANCE BTF 0.0 03/18/26", "EUR", 0.00, 1, "2026-03-18", 20_000_000, "GOVIES", "HTC", 2, "GOV", 12.0, 0.3),
    # ex-coupon inside the period: pays 30/09, which is the closing date
    ("XS2701234567", "SAMPLE EX-COUPON 5.0 09/30/31", "EUR", 5.00, 1, "2031-09-30", 10_000_000, "CREDIT", "AFS", 5, "SWAP", 135.0, -3.0),
    # a short position
    ("DE0001102580", "BUNDESREPUB. DEUTSCHLAND 1.7 08/15/32", "EUR", 1.70, 1, "2032-08-15", -16_000_000, "GOVIES", "HTC", 0, "GOV", 2.0, -0.3),
    # a callable, so OAS is the natural framework
    ("XS2621234098", "SAMPLE CALLABLE 5.25 06/15/33", "EUR", 5.25, 1, "2033-06-15", 6_000_000, "CREDIT", "AFS", 5, "SWAP", 210.0, -9.0),
]

# contract, ccy, contracts, ctd isin, cf, point value, px/bp, deliv, linked isin
FUTURES = [
    ("RXZ5", "EUX", "EUR", -180, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "DE0001102614"),
    ("RXZ5", "EUX", "EUR", -95, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "DE0001102390"),
    ("OATZ5", "EUX", "EUR", -140, "FR0014009O62", 0.702118, 1000.0, 0.0648, "2025-12-08", "FR0014009O62"),
    ("OATZ5", "EUX", "EUR", -46, "FR0013480613", 0.702118, 1000.0, 0.0648, "2025-12-08", "FR0013480613"),
    ("IKZ5", "EUX", "EUR", -155, "IT0005534141", 0.764902, 1000.0, 0.0592, "2025-12-08", "IT0005534141"),
    ("IKZ5", "EUX", "EUR", -70, "IT0005390874", 0.764902, 1000.0, 0.0592, "2025-12-08", "IT0005390874"),
    ("RXZ5", "EUX", "EUR", -105, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "ES0000012L52"),
    ("RXZ5", "EUX", "EUR", -78, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "PTOTEQOE0024"),
    ("RXZ5", "EUX", "EUR", -84, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "PTOTEUOE0021"),
    ("DUZ5", "EUX", "EUR", -220, "DE0001141851", 0.912344, 1000.0, 0.0189, "2025-12-08", "DE0001141851"),
    ("RXZ5", "EUX", "EUR", 62, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", "DE0001102580"),
    ("TYZ5", "CBT", "USD", -210, "US91282CJL65", 0.812004, 1000.0, 0.0712, "2025-12-19", "US91282CJL65"),
    ("USZ5", "CBT", "USD", -58, "US912810UB50", 0.640551, 1000.0, 0.1584, "2025-12-19", "US912810UB50"),
    # unlinked on purpose: a macro overlay that reaches no bond row
    ("RXZ5", "EUX", "EUR", -40, "DE0001102614", 0.681234, 1000.0, 0.0605, "2025-12-08", ""),
    ("FVZ5", "CBT", "USD", -30, "", 0.0, 1000.0, 0.0431, "2025-12-19", ""),
]

# deal, ccy, hedge fraction, fixed %, index, start, end, pay fixed, linked isin, class, cpty
#
# `hedge fraction` is the share of the linked bond's DV01 the swap is meant to
# neutralise.  Notional is SOLVED for in the generator so the hedge really does
# offset the bond - a sample book whose hedge ratios are all 1.15 would report
# a page of "Over-hedged" and teach nothing about the model.
SWAPS = [
    ("SW-100234", "EUR", 1.00, 2.61, "EURIBOR 6M", "2022-02-17", "2032-02-17", "Y", "XS2434891219", "PLAIN", "JPM"),
    ("SW-100235", "EUR", 1.00, 2.44, "EURIBOR 6M", "2021-07-13", "2031-07-13", "Y", "XS2364001078", "PLAIN", "GS"),
    ("SW-100236", "EUR", 0.98, 3.05, "EURIBOR 6M", "2023-02-23", "2032-02-23", "Y", "XS2600310275", "PLAIN", "BARC"),
    ("SW-100237", "EUR", 1.00, 2.72, "ESTR", "2023-01-09", "2030-01-09", "Y", "XS2542168951", "PLAIN", "DB"),
    ("SW-100238", "EUR", 0.95, 2.35, "EURIBOR 6M", "2021-04-15", "2031-04-15", "Y", "XS2333563817", "PLAIN", "SG"),
    ("SW-100239", "EUR", 1.00, 2.88, "EURIBOR 6M", "2022-07-09", "2034-07-09", "Y", "XS2445369418", "PLAIN", "HSBC"),
    ("SW-100240", "EUR", 1.00, 2.41, "ESTR", "2022-01-28", "2032-01-28", "Y", "XS2010044269", "PLAIN", "UBS"),
    ("SW-100241", "EUR", 0.99, 2.96, "EURIBOR 6M", "2023-04-30", "2030-04-30", "Y", "XS2588103241", "PLAIN", "CITI"),
    ("SW-100242", "EUR", 1.00, 2.91, "EURIBOR 6M", "2022-11-24", "2031-11-24", "Y", "XS2513282538", "PLAIN", "JPM"),
    ("SW-100243", "EUR", 0.92, 2.28, "EURIBOR 6M", "2021-07-05", "2032-07-05", "Y", "XS2244278008", "PLAIN", "MS"),
    ("SW-100244", "EUR", 1.00, 2.83, "EURIBOR 6M", "2023-02-03", "2033-02-03", "Y", "XS2453880566", "PLAIN", "BNP"),
    ("SW-100245", "EUR", 1.01, 2.79, "EURIBOR 6M", "2023-05-22", "2034-05-22", "Y", "XS2439096747", "PLAIN", "GS"),
    ("SW-100246", "EUR", 1.00, 2.74, "EURIBOR 6M", "2023-09-12", "2033-09-12", "Y", "XS2679523674", "PLAIN", "DB"),
    ("SW-100247", "EUR", 1.00, 2.86, "EURIBOR 6M", "2023-09-30", "2031-09-30", "Y", "XS2701234567", "PLAIN", "BARC"),
    ("SW-100248", "EUR", 1.00, 3.11, "EURIBOR 6M", "2023-06-15", "2033-06-15", "Y", "XS2621234098", "PLAIN", "HSBC"),
    ("SW-100249", "USD", 1.00, 3.79, "SOFR", "2022-12-01", "2029-12-01", "Y", "US023135CF19", "PLAIN", "CITI"),
    ("SW-100250", "USD", 1.00, 3.71, "SOFR", "2021-05-10", "2028-05-10", "Y", "US037833EK43", "PLAIN", "JPM"),
    # deliberately wrong-way, so the status check has something real to catch
    ("SW-100253", "EUR", -0.35, 2.59, "EURIBOR 6M", "2024-01-15", "2033-04-30", "N", "NL0015001AH9", "PLAIN", "UBS"),
    # synthetic targets: the hedge the coverage relationship says SHOULD be on.
    # Sized against the bond, but not matched by any real swap, so hedge
    # efficiency measures replication rather than flatness.
    ("SY-200011", "EUR", 1.00, 2.55, "EURIBOR 6M", "2023-04-30", "2033-04-30", "Y", "ES0000012L52", "SYNTHETIC", ""),
    ("SY-200012", "EUR", 1.00, 2.61, "EURIBOR 6M", "2023-06-18", "2038-06-18", "Y", "PTOTEQOE0024", "SYNTHETIC", ""),
    ("SY-200013", "EUR", 1.00, 2.58, "EURIBOR 6M", "2023-06-22", "2033-06-22", "Y", "BE0000359333", "SYNTHETIC", ""),
    # unlinked plain swap: real risk that reaches no bond row
    ("SW-100251", "EUR", None, 2.64, "EURIBOR 6M", "2024-03-15", "2034-03-15", "Y", "", "PLAIN", "SG"),
    # a deliberately unclassifiable floating index, small enough to read as a stray
    ("SW-100252", "EUR", 0.08, 2.70, "MYSTERY IDX", "2023-03-01", "2033-03-01", "Y", "XS2679523674", "PLAIN", "MS"),
]

FX = {"EUR": (1.0, 1.0), "USD": (0.85210, 0.85476), "GBP": (1.19840, 1.19755)}

_IMPERFECTIONS = """
Seeded on purpose, because a book where everything ties tests nothing:

  * two hedges with a blank LinkedISIN  -> the unlinked-PnL reconciliation
  * one plain swap with an unclassifiable floating index
                                        -> the model swap leg is suppressed
                                           rather than measured against a
                                           guessed curve
  * a bond whose coupon pays ON the closing date
                                        -> exercises the coupon-cash term that
                                           the legacy sheet had no place for
  * one short bond position             -> proves nothing double-applies a sign
  * three synthetic swap targets        -> must never reach a PnL total
  * one bond with an OAS-framework override
  * one bond with no prior duration     -> the model duration fallback
"""


def _price_book(curves: CurveSet, snapshot: str, spread_shift: bool):
    rows = []
    settle = np.datetime64(CURRENT if spread_shift else PRIOR, "D")
    for (
        isin, name, ccy, coupon, freq, maturity, notional, port, acctg, dcc,
        base, spread_bp, spread_move,
    ) in BONDS:
        mat = np.datetime64(maturity, "D")
        years = max((mat.astype("int64") - settle.astype("int64")) / 365.0, 0.0)
        curve = curves.require(ccy, base, snapshot)
        base_rate = float(curve.rate(np.array([years]))[0])
        spread = spread_bp + (spread_move if spread_shift else 0.0)
        ytm = base_rate + spread / 100.0

        clean, dirty = bondmath.bond_price_from_yield(
            np.array([settle], dtype="datetime64[D]"),
            np.array([mat], dtype="datetime64[D]"),
            np.array([coupon / 100.0]),
            np.array([ytm / 100.0]),
            np.array([float(freq)]),
            np.array([float(dcc)]),
        )
        moddur, convexity = bondmath.yield_risk(
            np.array([settle], dtype="datetime64[D]"),
            np.array([mat], dtype="datetime64[D]"),
            np.array([coupon / 100.0]),
            np.array([ytm / 100.0]),
            np.array([float(freq)]),
            np.array([float(dcc)]),
            bump_bp=100.0,
        )

        # Curve rates the bond's quoted spreads are measured against.
        gov = float(curves.require(ccy, "GOV", snapshot).rate(np.array([years]))[0])
        swp = float(curves.require(ccy, "SWAP", snapshot).rate(np.array([years]))[0])

        rows.append(
            {
                "isin": isin, "name": name, "currency": ccy, "portfolio": port,
                "acctg_cat": acctg, "coupon_pct": coupon, "coupon_freq": freq,
                "maturity": maturity, "notional": notional,
                "day_count_code": dcc,
                "ytm": ytm,
                "clean": float(clean[0]), "dirty": float(dirty[0]),
                "moddur": float(moddur[0]), "convexity": float(convexity[0]),
                # Z, ASW and OAS are quoted on their own conventions, so they
                # are seeded near the I-spread but not identical to it - which
                # is exactly why their chains tie only approximately.
                "zspread": (ytm - swp) * 100.0 + 3.0,
                "asw": (ytm - swp) * 100.0 - 2.5,
                "oas": (ytm - swp) * 100.0 - 6.0,
                "gspread": (ytm - gov) * 100.0,
            }
        )
    return rows


def main() -> None:
    curves = load_curves(HERE / "curves.csv")
    prior_rows = _price_book(curves, "prior", spread_shift=False)
    current_rows = _price_book(curves, "current", spread_shift=True)

    # ---- bonds ------------------------------------------------------------- #
    fields = [
        "isin", "name", "currency", "portfolio", "acctg_cat", "coupon_pct",
        "coupon_freq", "maturity", "notional", "book_value", "day_count_code",
        "clean_px_prior", "dirty_px_prior", "ytm_prior", "zspread_prior",
        "asw_prior", "oas_prior", "mod_duration_prior",
        "clean_px_current", "dirty_px_current", "ytm_current",
        "zspread_current", "asw_current", "oas_current",
        "mod_duration_current", "oas_mod_duration_current", "convexity",
        "fx_prior", "fx_current", "funding_rate_prior", "funding_rate_current",
        "pricing_source",
    ]
    with (HERE / "positions_bonds.csv").open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        for i, (p, c) in enumerate(zip(prior_rows, current_rows)):
            fx_p, fx_c = FX[p["currency"]]
            funding = 2.91 if p["currency"] == "EUR" else 4.82
            # One bond arrives with no prior duration, to exercise the model
            # duration fallback rather than assuming it is never needed.
            mod_prior = "" if p["isin"] == "XS2439096747" else round(p["moddur"], 6)
            writer.writerow(
                {
                    "isin": p["isin"], "name": p["name"], "currency": p["currency"],
                    "portfolio": p["portfolio"], "acctg_cat": p["acctg_cat"],
                    "coupon_pct": p["coupon_pct"], "coupon_freq": p["coupon_freq"],
                    "maturity": p["maturity"], "notional": p["notional"],
                    "book_value": round(p["notional"] * p["dirty"] / 100.0 * fx_p, 2),
                    "day_count_code": p["day_count_code"],
                    "clean_px_prior": round(p["clean"], 6),
                    "dirty_px_prior": round(p["dirty"], 6),
                    "ytm_prior": round(p["ytm"], 6),
                    "zspread_prior": round(p["zspread"], 4),
                    "asw_prior": round(p["asw"], 4),
                    "oas_prior": round(p["oas"], 4),
                    "mod_duration_prior": mod_prior,
                    "clean_px_current": round(c["clean"], 6),
                    "dirty_px_current": round(c["dirty"], 6),
                    "ytm_current": round(c["ytm"], 6),
                    "zspread_current": round(c["zspread"], 4),
                    "asw_current": round(c["asw"], 4),
                    "oas_current": round(c["oas"], 4),
                    "mod_duration_current": round(c["moddur"], 6),
                    "oas_mod_duration_current": "",
                    "convexity": round(c["convexity"], 6),
                    "fx_prior": fx_p, "fx_current": fx_c,
                    "funding_rate_prior": funding, "funding_rate_current": funding,
                    "pricing_source": "BGN",
                }
            )

    # ---- futures ----------------------------------------------------------- #
    # Futures prices are derived from the CTD: a bond future tracks its CTD's
    # forward clean price divided by the conversion factor, so moving the CTD
    # moves the contract and the hedge really does offset the bond.
    by_isin_prior = {r["isin"]: r for r in prior_rows}
    by_isin_current = {r["isin"]: r for r in current_rows}

    fut_fields = [
        "contract_code", "exchange", "currency", "portfolio", "contracts",
        "face_value", "deliv_date", "ctd_isin", "ctd_cf", "linked_isin",
        "hedge_type", "avg_entry_px", "fut_px_prior", "fut_px_current",
        "fut_val_pt", "fut_px_val_bp", "ctd_dirty_px_current",
        "fx_prior", "fx_current", "implied_repo_bbg", "net_basis_bbg",
        "gross_basis_bbg",
    ]
    with (HERE / "positions_futures.csv").open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fut_fields)
        writer.writeheader()
        for code, exch, ccy, contracts, ctd, cf, point, pxbp, deliv, linked in FUTURES:
            fx_p, fx_c = FX[ccy]
            if ctd and ctd in by_isin_prior:
                p_clean = by_isin_prior[ctd]["clean"]
                c_clean = by_isin_current[ctd]["clean"]
                ctd_dirty = by_isin_current[ctd]["dirty"]
                # A small, stable gross basis so the futures basis line is real
                # but not noise.
                px_prior = (p_clean - 0.145) / cf
                px_current = (c_clean - 0.152) / cf
            else:
                px_prior, px_current, ctd_dirty = 108.42, 108.39, ""
            writer.writerow(
                {
                    "contract_code": code, "exchange": exch, "currency": ccy,
                    "portfolio": "HEDGE", "contracts": contracts,
                    "face_value": 100_000, "deliv_date": deliv,
                    "ctd_isin": ctd, "ctd_cf": cf, "linked_isin": linked,
                    "hedge_type": "DIRECT" if linked else "MACRO",
                    "avg_entry_px": round(px_prior, 4),
                    "fut_px_prior": round(px_prior, 6),
                    "fut_px_current": round(px_current, 6),
                    "fut_val_pt": point, "fut_px_val_bp": pxbp,
                    "ctd_dirty_px_current": round(ctd_dirty, 6) if ctd_dirty else "",
                    "fx_prior": fx_p, "fx_current": fx_c,
                    "implied_repo_bbg": 2.88 if ccy == "EUR" else 4.79,
                    "net_basis_bbg": 0.041, "gross_basis_bbg": 0.148,
                }
            )

    # ---- swaps -------------------------------------------------------------- #
    # NPVs are built from the swap's own DV01 against its own curve move, plus a
    # small spread mismatch, so the swap basis line is non-zero but small.
    sw_fields = [
        "deal_id", "currency", "portfolio", "counterparty", "linked_isin",
        "swap_id_source", "notional", "fixed_rate", "float_index",
        "float_spread", "start_date", "end_date", "pay_fixed",
        "float_curve_type", "dv01_supplied", "npv_prior", "npv_current",
        "fx_prior", "fx_current",
    ]
    # Bond DV01 at the close, so a swap notional can be solved for.
    bond_dv01 = {}
    for p, c in zip(prior_rows, current_rows):
        fx_c = FX[c["currency"]][1]
        bond_dv01[c["isin"]] = (
            abs(c["notional"]) * c["moddur"] * abs(c["dirty"]) / 100.0 * fx_c * 1e-4
        )

    with (HERE / "positions_swaps.csv").open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=sw_fields)
        writer.writeheader()
        for deal, ccy, fraction, fixed, index, start, end, payfix, linked, cls, cpty in SWAPS:
            fx_p, fx_c = FX[ccy]
            years = (np.datetime64(end, "D").astype("int64")
                     - np.datetime64(CURRENT, "D").astype("int64")) / 365.0
            family = "EURIBOR" if "EURIBOR" in index else ("ESTR" if "ESTR" in index else
                     ("SOFR" if "SOFR" in index else ""))
            curve_kind = "SWAP" if family == "EURIBOR" else "OIS"
            if family:
                r_p = float(curves.require(ccy, curve_kind, "prior").rate(np.array([years]))[0])
                r_c = float(curves.require(ccy, curve_kind, "current").rate(np.array([years]))[0])
            else:
                r_p = r_c = float("nan")

            sign = -1.0 if payfix.upper().startswith("Y") else 1.0
            z = r_c / 100.0 if r_c == r_c else 0.0
            df = (1.0 + z) ** (-years) if 1.0 + z > 0 else 0.0
            annuity = (1.0 - df) / z if abs(z) > 1e-9 else years

            # Solve the notional from the DV01 the swap is meant to carry:
            #     DV01 = |notional| * 1bp * annuity * fx
            target = bond_dv01.get(linked)
            if fraction is not None and target:
                notional = round(abs(fraction) * target / (1e-4 * annuity * fx_c), -3)
                if fraction < 0:
                    sign = -sign  # a wrong-way hedge, on purpose
            else:
                notional = 15_000_000

            # A 1.5% haircut, so the supplied risk is close to but not identical
            # with the internal annuity model and the two can be compared.
            dv01 = sign * abs(notional) * 1e-4 * annuity * fx_c * 0.985

            if cls == "SYNTHETIC" or r_p != r_p:
                npv_p = npv_c = ""
            else:
                move_bp = (r_c - r_p) * 100.0
                npv_p = 0.0
                # actual = model plus a small spread mismatch, which is what the
                # swap basis line is for
                npv_c = round(-dv01 * move_bp + abs(notional) * 1e-7, 2)
            writer.writerow(
                {
                    "deal_id": deal, "currency": ccy, "portfolio": "HEDGE",
                    "counterparty": cpty, "linked_isin": linked,
                    "swap_id_source": cls, "notional": notional,
                    "fixed_rate": fixed, "float_index": index,
                    "float_spread": 0.0, "start_date": start, "end_date": end,
                    "pay_fixed": payfix, "float_curve_type": family,
                    "dv01_supplied": round(dv01, 4) if cls == "PLAIN" else round(dv01, 4),
                    "npv_prior": npv_p, "npv_current": npv_c,
                    "fx_prior": fx_p, "fx_current": fx_c,
                }
            )

    # ---- fx and overrides --------------------------------------------------- #
    with (HERE / "fx.csv").open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["currency", "fx_prior", "fx_current"])
        for ccy, (p, c) in FX.items():
            writer.writerow([ccy, p, c])

    with (HERE / "spread_overrides.csv").open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["isin", "framework", "reason"])
        writer.writerow(["XS2621234098", "OAS", "callable - optionality must be stripped"])
        writer.writerow(["IT0005390874", "Z", "desk measures this line on Z-spread"])

    print("sample book written to", HERE)
    print(_IMPERFECTIONS)


if __name__ == "__main__":
    main()
