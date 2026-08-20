#!/usr/bin/env python3
"""
Parse a *model / target allocation* Excel (weights only) into parsed.json —
the same data file `build_proposal.py` / `build_line_items_tables.py` read.

This is the sibling of `parse_portfolio.py`. `parse_portfolio.py` consumes a
custodian export of a client's *current* holdings (market values, ISINs,
ratings, durations, yields). This script consumes a much thinner input: a
proposed model allocation whose only hard number per line is a **weight**.
The two emit the *same* parsed.json schema so every downstream slide-filler
is reused unchanged.

Input shape (one sheet), a header row `Instrument | Type | Ccy | Wgt`
followed by section-header rows and instrument rows:

    Instrument                         Type   Ccy   Wgt
    Cash                                            0.10      <- section header
    Fixed income via funds                          0.12      <- section header
    BNY Mellon Global Short-Dated HY   Fund   USD   0.035     <- instrument
    ...

A **section header** is a row whose Type and Ccy are both blank; its Wgt is
the section subtotal. An **instrument** row has a Type and a Ccy. The header
row (`Instrument/Type/Ccy/Wgt`) may repeat and is skipped; a trailing
grand-total row (blank name, Wgt = 1.0) is skipped.

Because the input carries no market values, all monetary figures in the deck
are *indicative*, computed off a nominal base amount (``--nominal``, default
CHF 1,000,000). Fields a weights-only file genuinely cannot support
(portfolio-wide geography/sector, running yield, fixed-income duration,
income by sleeve, an individual bond selection) are emitted as the schema's
"unavailable" sentinels (``None`` / ``{"Not available": 100.0}`` / empty
list) exactly as `parse_portfolio.py` does when the custodian file is missing
a column — the deck then shows "Not available"/"n/a" rather than a
fabricated number (SKILL.md's core rule).

Usage:
    python parse_model_allocation.py Prop_Marco.xlsx \
        --profile Balanced --base-currency CHF --nominal 1000000 \
        --client-name "[Client Name]" -o parsed.json
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

RISK_PROFILES = ["Fixed Income", "Conservative", "Moderate", "Balanced", "Growth", "Equity"]

# Section header (normalized) -> asset-class bucket used throughout
# parsed.json. "Alternatives" maps to "Private Assets" because that is the
# sleeve key build_proposal.py's SLEEVE_SLIDES uses for the Alternatives
# slide (slide_recipe.md's "'Alternatives' naming" note).
SECTION_TO_CLASS = {
    "cash": "Cash",
    "fixed income via funds": "Fixed Income",
    "fixed income": "Fixed Income",
    "equities": "Equities",
    "equity structured products (arc)": "Structured Products",
    "structured products": "Structured Products",
    "commodities": "Commodities",
    "alternatives": "Private Assets",
}

GROWTH_CLASSES = {"Equities", "Alternatives", "Commodities", "Structured Products", "Private Assets"}
DEFENSIVE_CLASSES = {"Cash", "Fixed Income"}

# Vehicle label per raw Type value (for the sleeve vehicle-breakdown donut).
VEHICLE_LABELS = {
    "fund": "Fund",
    "etf": "ETF",
    "sp": "Structured product (ARC)",
    "hf": "Hedge fund",
    "amc": "Actively-managed certificate",
    "private credit": "Private credit",
    "cash": "Cash",
}

# Liquidity bucket per raw Type value. The key "Illiquid (lock-up)" must
# match build_proposal.fill_liquidity()'s lookup verbatim.
LIQUIDITY_BUCKETS = {
    "cash": "Cash & equivalents",
    "fund": "Daily-liquid fund / ETF",
    "etf": "Daily-liquid fund / ETF",
    "sp": "Structured product (periodic liquidity)",
    "amc": "Structured product (periodic liquidity)",
    "hf": "Illiquid (lock-up)",
    "private credit": "Illiquid (lock-up)",
}
LIQUID_BUCKETS = {"Cash & equivalents", "Daily-liquid fund / ETF"}

# Equity-region order matches parse_portfolio.py / build_line_items_tables.py.
EQUITY_REGION_ORDER = ["US", "Eurozone", "UK", "Switzerland", "Japan", "EM", "Global / thematic"]
FI_BUCKET_ORDER = ["Govies 1-10 (local)", "Govies 10+ (local)", "High Yield (local or global hdg)",
                   "Corporate IG (local)", "EM Debt"]
TOP_CAT_ORDER = ["Cash", "Fixed Income", "Equities", "Alternatives", "Structured Products",
                 "Private Assets", "Commodities", "Other"]


def norm(s) -> str:
    return " ".join(str(s or "").split()).strip().lower()


def classify_equity_region(name: str) -> str:
    """Region of an equity fund/ETF, inferred from its (self-describing) name.
    These model lines name their mandate explicitly ('S&P500', 'Eurozone',
    'Swiss', 'Japan', 'Emerging Markets', 'China'), so this is a reading of
    the label, not a look-through guess."""
    n = norm(name)
    if any(k in n for k in ("s&p500", "s&p 500", "us small", "u.s.", " us ", "nasdaq", "russell")):
        return "US"
    if "eurozone" in n or "euro stoxx" in n or "eurostoxx" in n:
        return "Eurozone"
    if "swiss" in n or " smi" in n or "switzerland" in n:
        return "Switzerland"
    if "japan" in n:
        return "Japan"
    if "emerging" in n or " em " in n or "china" in n:
        return "EM"
    return "Global / thematic"


def classify_fi_bucket(name: str) -> str:
    """Fixed-income sub-bucket from a fund's (self-describing) name."""
    n = norm(name)
    if "em " in n or "emerging" in n or "em debt" in n:
        return "EM Debt"
    if "hy" in n or "high yield" in n or "high-yield" in n:
        return "High Yield (local or global hdg)"
    if "ig" in n or "investment grade" in n or "obligataire" in n or "crossover" in n:
        return "Corporate IG (local)"
    return "Corporate IG (local)"


def classify_private_asset_group(name: str, vtype: str) -> str:
    t = norm(vtype)
    if t == "hf":
        return "Hedge funds"
    if t == "private credit":
        return "Private credit"
    return "Alternatives (other)"


def pct(x: float) -> float:
    return round(x * 100, 6)


def load_lines(ws):
    """Yield (asset_class, section_name, instrument_dict) for every
    instrument row, tracking the current section. Cash sections with no
    instrument rows are materialized as a single synthetic 'Cash &
    equivalents' line carrying the section weight."""
    header_seen = False
    current_class = None
    current_section = None
    section_weight = {}
    lines = []
    cash_line_added = defaultdict(bool)

    for row in ws.iter_rows(values_only=True):
        vals = list(row) + [None] * (4 - len(row))
        name, vtype, ccy, wgt = vals[0], vals[1], vals[2], vals[3]
        if name is None and vtype is None and ccy is None and (wgt is None or wgt in (1, 1.0)):
            continue  # blank row or trailing grand-total
        if norm(name) == "instrument":
            header_seen = True
            continue
        is_section = (vtype in (None, "")) and (ccy in (None, ""))
        if is_section:
            key = norm(name)
            if key not in SECTION_TO_CLASS:
                raise ValueError(f"Unknown section header {name!r} — add it to SECTION_TO_CLASS")
            current_class = SECTION_TO_CLASS[key]
            current_section = " ".join(str(name).split()).strip()
            try:
                section_weight[current_class] = float(wgt or 0)
            except (TypeError, ValueError):
                section_weight[current_class] = 0.0
            continue
        if current_class is None:
            continue
        try:
            w = float(wgt or 0)
        except (TypeError, ValueError):
            w = 0.0
        if w <= 0:
            continue
        lines.append({
            "asset_class": current_class,
            "section": current_section,
            "name": " ".join(str(name).split()).strip(),
            "type": str(vtype).strip(),
            "currency": str(ccy).strip() if ccy else None,
            "weight": w,
        })

    # Materialize a Cash line if the Cash section carried a weight but no rows.
    classes_with_lines = {l["asset_class"] for l in lines}
    if "Cash" in section_weight and "Cash" not in classes_with_lines and section_weight["Cash"] > 0:
        lines.append({
            "asset_class": "Cash",
            "section": "Cash",
            "name": "Cash & equivalents",
            "type": "Cash",
            "currency": None,
            "weight": section_weight["Cash"],
        })
    return lines, section_weight


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel_path")
    ap.add_argument("--profile", required=True, choices=RISK_PROFILES)
    ap.add_argument("--client-name", default="[Client Name]")
    ap.add_argument("--base-currency", default="CHF")
    ap.add_argument("--nominal", type=float, default=1_000_000.0,
                    help="Nominal base amount for indicative money figures (default 1,000,000)")
    ap.add_argument("--valuation-date", default=None, help="e.g. '20 August 2026'; defaults to today")
    ap.add_argument("--sheet", default=None, help="sheet name (default: first sheet)")
    ap.add_argument("-o", "--output", default="parsed.json")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.excel_path, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.worksheets[0]
    lines, section_weight = load_lines(ws)
    if not lines:
        sys.exit("Error: no instrument lines found in the model-allocation sheet")

    base_ccy = args.base_currency
    nominal = args.nominal
    # Cash lines have no currency in the source; a CHF mandate's cash is CHF.
    for l in lines:
        if l["currency"] in (None, ""):
            l["currency"] = base_ccy

    total_w = sum(l["weight"] for l in lines)
    if abs(total_w - 1.0) > 0.02:
        print(f"Warning: instrument weights sum to {total_w:.4f}, not ~1.0", file=sys.stderr)

    # --- asset allocation ---
    class_weight = defaultdict(float)
    for l in lines:
        class_weight[l["asset_class"]] += l["weight"]
    asset_allocation = {k: pct(v) for k, v in sorted(class_weight.items(), key=lambda kv: -kv[1])}

    # --- currency exposure ---
    ccy_weight = defaultdict(float)
    for l in lines:
        ccy_weight[l["currency"] or "Other"] += l["weight"]
    currency_exposure = {k: pct(v) for k, v in sorted(ccy_weight.items(), key=lambda kv: -kv[1])}
    dominant_ccy = max(ccy_weight.items(), key=lambda kv: kv[1])[0] if ccy_weight else base_ccy

    # --- sleeves: top holdings + vehicle breakdown per asset class ---
    def hold(l):
        val = l["weight"] * nominal
        return {
            "name": l["name"], "isin": None,
            "value_eur": round(val, 2), "value_qc": round(val, 2),
            "currency": base_ccy, "weight_pct": pct(l["weight"]),
            "duration": None, "yield_pct": None, "price": None,
        }

    sleeves = {}
    for ac in sorted(set(l["asset_class"] for l in lines)):
        rows_ac = [l for l in lines if l["asset_class"] == ac]
        ac_w = sum(l["weight"] for l in rows_ac) or 1.0
        top = sorted(rows_ac, key=lambda l: -l["weight"])[:5]
        vehicle_w = defaultdict(float)
        for l in rows_ac:
            label = VEHICLE_LABELS.get(norm(l["type"]), l["type"] or "Other")
            vehicle_w[label] += l["weight"] / ac_w
        sleeves[ac] = {
            "total_weight_pct": pct(sum(l["weight"] for l in rows_ac)),
            "num_lines": len(rows_ac),
            "top_holdings": [hold(l) for l in top],
            "vehicle_breakdown_pct": {k: pct(v) for k, v in sorted(vehicle_w.items(), key=lambda kv: -kv[1])},
            # No duration data in a weights-only model — blanks the FI caveat.
            "duration_coverage": None,
        }

    # --- equity breakdown (renormalized to the equity sleeve) ---
    equity_rows = [l for l in lines if l["asset_class"] == "Equities"]
    eq_w = sum(l["weight"] for l in equity_rows) or 1.0
    eq_geo = defaultdict(float)
    for l in equity_rows:
        eq_geo[classify_equity_region(l["name"])] += l["weight"] / eq_w
    equity_breakdown = {
        "by_geography_pct": {k: pct(v) for k, v in sorted(
            eq_geo.items(), key=lambda kv: EQUITY_REGION_ORDER.index(kv[0]) if kv[0] in EQUITY_REGION_ORDER else 99)},
        # Sector and market cap need issuer-level look-through not present in a
        # model of funds/ETFs — reported unavailable, not fabricated.
        "by_sector_pct": None,
        "by_market_cap_pct": None,
    }

    # --- concentration & top holdings (weights are enough for this) ---
    by_weight = sorted(lines, key=lambda l: -l["weight"])
    cum = 0.0
    conc_rows = []
    for i, l in enumerate(by_weight[:20], 1):
        cum += l["weight"]
        conc_rows.append({
            "rank": i, "name": l["name"], "value_eur": round(l["weight"] * nominal, 2),
            "weight_pct": pct(l["weight"]), "cumulative_pct": pct(cum),
        })
    concentration = {
        "top_holdings": conc_rows,
        "top_5_pct": pct(sum(l["weight"] for l in by_weight[:5])),
        "top_10_pct": pct(sum(l["weight"] for l in by_weight[:10])),
        "top_20_pct": pct(sum(l["weight"] for l in by_weight[:20])),
        "positions_above_5pct": sum(1 for l in lines if l["weight"] > 0.05),
    }

    # --- liquidity profile ---
    liq_w = defaultdict(float)
    for l in lines:
        liq_w[LIQUIDITY_BUCKETS.get(norm(l["type"]), "Listed")] += l["weight"]
    liquidity_profile = {k: pct(v) for k, v in sorted(liq_w.items(), key=lambda kv: -kv[1])}
    liquid_share = pct(sum(v for k, v in liq_w.items() if k in LIQUID_BUCKETS))

    # --- income & rate sensitivity: unavailable in a weights-only model ---
    income = {
        "running_yield_pct": None,
        "fi_duration_years": None,
        "rate_impact_100bp_eur": None,
        "income_by_sleeve": [],
        "fi_duration_subsleeves": [],
    }

    # --- risk profile / growth-defensive split ---
    growth_w = sum(v for k, v in class_weight.items() if k in GROWTH_CLASSES)
    defensive_w = sum(v for k, v in class_weight.items() if k in DEFENSIVE_CLASSES)
    risk_profile = {
        "selected_profile": args.profile,
        "growth_assets_pct": pct(growth_w),
        "defensive_assets_pct": pct(defensive_w),
    }

    # --- line items (slides 9-10) ---
    line_items = []
    for l in lines:
        ac = l["asset_class"]
        if ac == "Fixed Income":
            group = classify_fi_bucket(l["name"])
        elif ac == "Equities":
            group = classify_equity_region(l["name"])
        elif ac == "Cash":
            group = "Cash"
        elif ac == "Structured Products":
            group = "Autocallables (ARC)"
        elif ac == "Private Assets":
            group = classify_private_asset_group(l["name"], l["type"])
        else:
            group = ac
        line_items.append({
            "top_category": ac, "group": group, "isin": None,
            "instrument": l["name"], "weight_pct": pct(l["weight"]),
        })

    def sort_key(li):
        tc = TOP_CAT_ORDER.index(li["top_category"]) if li["top_category"] in TOP_CAT_ORDER else 99
        if li["top_category"] == "Fixed Income":
            gi = FI_BUCKET_ORDER.index(li["group"]) if li["group"] in FI_BUCKET_ORDER else 99
        elif li["top_category"] == "Equities":
            gi = EQUITY_REGION_ORDER.index(li["group"]) if li["group"] in EQUITY_REGION_ORDER else 99
        else:
            gi = 0
        return (tc, gi, -li["weight_pct"])

    line_items.sort(key=sort_key)

    import datetime as _dt
    valuation_date = args.valuation_date or _dt.date.today().strftime("%-d %B %Y")

    parsed = {
        "client_name": args.client_name,
        "risk_profile_input": args.profile,
        "valuation_date": valuation_date,
        "base_currency": base_ccy,
        "dominant_holding_currency": dominant_ccy,
        "total_value_eur": round(nominal, 2),   # nominal base; schema field name kept for the builder
        "num_positions": len(lines),
        "largest_position_pct": pct(by_weight[0]["weight"]) if by_weight else 0,
        "liquid_share_pct": liquid_share,
        "asset_allocation_pct": asset_allocation,
        "currency_exposure_pct": currency_exposure,
        # Portfolio-wide geography/sector need fund look-through absent here.
        "geographic_exposure_pct": {"Not available": 100.0},
        "sector_exposure_pct": None,
        "sleeves": sleeves,
        "equity_breakdown": equity_breakdown,
        "concentration": concentration,
        "liquidity_profile_pct": liquidity_profile,
        "income": income,
        "risk_profile": risk_profile,
        "uncalled_commitments": [],
        "line_items": line_items,
        "proposed_bond_selection": None,   # no curated bond ladder in a model of funds
        # --- extras consumed only by the model-allocation build path ---
        "is_indicative": True,
        "basis_note": f"Indicative amounts shown on a nominal {base_ccy} {nominal/1_000_000:.1f}m base.",
    }

    Path(args.output).write_text(json.dumps(parsed, indent=2, default=str))
    print(f"Wrote {args.output}: {len(lines)} lines, nominal {base_ccy} {nominal:,.0f}, "
          f"growth assets {risk_profile['growth_assets_pct']:.1f}%, dominant currency {dominant_ccy}")


if __name__ == "__main__":
    main()
