#!/usr/bin/env python3
"""
Parse a client custodian Excel export into parsed.json — the single data
file every other build script (update_chart.py, build_line_items_tables.py,
build_proposal.py) reads from. Never edit parsed.json by hand; re-run this
script if the source Excel changes.

Every rule this script applies is documented in ../references/assumptions.md
by section number; the code comments below reference those sections instead
of re-explaining the rule.

Usage:
    python parse_portfolio.py Portfolio.xlsx --profile Balanced -o parsed.json
    python parse_portfolio.py Portfolio.xlsx --profile Balanced \
        --bucket-overrides overrides.csv --market-cap-overrides mcap.csv \
        -o parsed.json
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "references"))
from geography_map import region_of, equity_region_of, equity_region_from_name  # noqa: E402

PORTFOLIO_SECTIONS = {
    "Cash",
    "Short-term instruments",
    "Bonds",
    "Equities",
    "Structured products",
    "Private Assets",
    "Alternative funds",
    "Gold and other commodities",
    "Other investments",
}

# French-language custodian export (assumptions.md §1b): section headers in
# column A are translated to the English section names above, so every rule
# downstream of load_sheet_rows() sees one vocabulary.
SECTION_ALIASES = {
    "avoirs en compte": "Cash",
    "placements à court terme": "Short-term instruments",
    "obligations": "Bonds",
    "actions": "Equities",
    "produits structurés": "Structured products",
    "actifs privés": "Private Assets",
    "fonds alternatifs": "Alternative funds",
    "or et autres matières premières": "Gold and other commodities",
    "autres placements": "Other investments",
}

# French column headers -> the English normalized header names every rule in
# this script reads. The base-currency valuation column is matched separately
# by VALUATION_HEADER_RE because its currency code varies by file.
HEADER_ALIASES = {
    "solde": "balance",
    "libellé": "description",
    "devise": "currency",
    "cours (mc)": "price (qc)",
    "poids (%)": "weight (%)",
}
VALUATION_HEADER_RE = re.compile(
    r"^(?:valuation \+ accr\. interest|estimation \+ int\. courus) \(([a-z]{3})\)$")

# Maps a raw Portfolio-sheet section header to the asset-class bucket used
# throughout parsed.json and the slide 9/10 line-items table (assumptions.md §4-5).
ASSET_CLASS_MAP = {
    "Cash": "Cash",
    "Short-term instruments": "Fixed Income",
    "Bonds": "Fixed Income",
    "Equities": "Equities",
    "Structured products": "Structured Products",
    "Private Assets": "Private Assets",
    "Alternative funds": "Alternatives",
    "Gold and other commodities": "Commodities",
    "Other investments": "Other",
}

GROWTH_CLASSES = {"Equities", "Alternatives", "Commodities", "Structured Products", "Private Assets"}
DEFENSIVE_CLASSES = {"Cash", "Fixed Income"}

RISK_PROFILES = ["Fixed Income", "Conservative", "Moderate", "Balanced", "Growth", "Equity"]

IG_RATINGS = {"AAA", "AA+", "AA", "AA-"}
CORP_IG_RATINGS = {"A+", "A", "A-", "BBB+", "BBB", "BBB-"}

FUND_KEYWORDS = ("FUND", " FD ", "SICAV", "ETF", "ETC")
STRUCTURED_KEYWORDS = ("CERTIFICATE", " CERT", "AMC", "TRACKER", "DCI-", "RAF ", "CPN ", " CLN ")
# A direct bond line in the French export carries its maturity date in the
# instrument name ("4.85% HEWLETT PACK OCT31  15.10.2031"); funds never do.
BOND_MATURITY_IN_NAME_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")


def norm_header(h) -> str:
    return " ".join(str(h or "").split()).strip().lower()


def canonical_section(value) -> str | None:
    """English section name for a column-A header cell, or None if the cell
    isn't a section header (assumptions.md §1, §1b)."""
    if not isinstance(value, str):
        return None
    v = value.strip()
    if v in PORTFOLIO_SECTIONS:
        return v
    return SECTION_ALIASES.get(" ".join(v.split()).lower())


def load_sheet_rows(ws) -> tuple[dict[str, int], list[dict], str]:
    """Find the header row (the first row carrying both a description column
    and a weight column, in either the English or the French export) then
    yield every row below it as a dict keyed by normalized English header
    name, tagging each with its section from column A.

    Returns (headers, rows, base_currency). The base currency is read from
    the valuation column's own header, e.g. "Valuation + accr. interest
    (EUR)" or "Estimation + int. courus (USD)" (assumptions.md §3); each
    row's value in that column lands under the "_value_base" key."""
    header_row_idx = None
    headers: dict[str, int] = {}
    base_currency = None
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        normed = [HEADER_ALIASES.get(norm_header(v), norm_header(v)) for v in vals]
        if "description" in normed and "weight (%)" in normed:
            header_row_idx = r
            for c, (v, h) in enumerate(zip(vals, normed), 1):
                if v in (None, ""):
                    continue
                m = VALUATION_HEADER_RE.match(norm_header(v))
                if m:
                    base_currency = m.group(1).upper()
                    headers["_value_base"] = c
                else:
                    headers[h] = c
            break
    if header_row_idx is None:
        raise ValueError(f"Could not find a header row with Description/Libellé and Weight/Poids "
                         f"columns in sheet '{ws.title}'")
    if base_currency is None:
        raise ValueError(f"Could not find a base-currency valuation column "
                         f"('Valuation + accr. interest (XXX)' or 'Estimation + int. courus (XXX)') "
                         f"in sheet '{ws.title}'")

    rows = []
    section = None
    for r in range(header_row_idx + 1, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        sec = canonical_section(col_a)
        if sec:
            section = sec
            continue
        row = {h: ws.cell(r, c).value for h, c in headers.items()}
        if isinstance(col_a, str) and col_a.strip() and all(v in (None, "") for v in row.values()):
            # A label in column A with nothing else on the row is a section
            # header we don't know. Failing loudly beats silently filing its
            # holdings under the previous section's asset class.
            raise ValueError(f"Unknown section header '{col_a.strip()}' in sheet '{ws.title}' "
                             f"(row {r}); add it to PORTFOLIO_SECTIONS or SECTION_ALIASES")
        if all(v in (None, "") for v in row.values()):
            continue
        row["_section"] = section
        rows.append(row)
    return headers, rows, base_currency


def load_bucket_overrides(path: str | None) -> dict[str, str]:
    if not path:
        return {}
    out = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            out[row["isin"].strip()] = row["bucket"].strip()
    return out


def load_market_cap_overrides(path: str | None) -> dict[str, str]:
    if not path:
        return {}
    out = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            out[row["isin"].strip()] = row["market_cap"].strip()
    return out


def load_sector_overrides(path: str | None) -> dict[str, str]:
    if not path:
        return {}
    out = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            out[row["isin"].strip()] = row["sector"].strip()
    return out


def instrument_key(value) -> str:
    return " ".join(str(value or "").split()).upper()


def load_instrument_overrides(path: str | None) -> dict[str, dict[str, str]]:
    """assumptions.md §14. CSV columns: instrument,region,vehicle,fi_group.
    `instrument` is an ISIN or the exact instrument description; blank cells
    mean "no override for this field"."""
    if not path:
        return {}
    out = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            key = instrument_key(row.get("instrument"))
            if not key:
                continue
            out[key] = {k: (row.get(k) or "").strip() for k in ("region", "vehicle", "fi_group")}
    return out


def instrument_override(row: dict, overrides: dict[str, dict[str, str]], field: str) -> str | None:
    for key in (row.get("isin code"), row.get("description")):
        hit = overrides.get(instrument_key(key))
        if hit and hit.get(field):
            return hit[field]
    return None


def classify_fi_bucket_by_structure(row: dict) -> str:
    """assumptions.md §4b: used when the export has no rating column, so the
    only honest grouping is by instrument structure, read from the section
    header and the instrument name."""
    desc = f" {str(row.get('description') or '').upper()} "
    if row.get("_section") == "Short-term instruments":
        return "Money market & short-term"
    if " CLN " in desc:
        return "Credit-linked notes"
    if BOND_MATURITY_IN_NAME_RE.search(desc):
        return "Direct bonds"
    return "Bond funds"


def classify_fi_bucket(row: dict, overrides: dict[str, str], has_rating: bool = True,
                       instrument_overrides: dict | None = None) -> str:
    """assumptions.md §4. Note: the Portfolio sheet carries no Sector column
    (only the separate 'Fixed Income' proposal tab does), so this rule uses
    Composite Rating + Duration + Geographical breakdown only."""
    isin = str(row.get("isin code") or "").strip()
    if isin in overrides:
        return overrides[isin]
    forced = instrument_override(row, instrument_overrides or {}, "fi_group")
    if forced:
        return forced
    if not has_rating:
        return classify_fi_bucket_by_structure(row)

    geography = row.get("geographical breakdown")
    if region_of(geography) == "MEA EM":
        return "EM Debt"

    rating = str(row.get("composite rating") or "").strip()
    duration = row.get("duration")

    if rating in IG_RATINGS:
        try:
            d = float(duration)
        except (TypeError, ValueError):
            d = 0.0
        return "Govies 10+ (local)" if d > 10 else "Govies 1-10 (local)"
    if rating in CORP_IG_RATINGS:
        return "Corporate IG (local)"
    return "High Yield (local or global hdg)"


def classify_vehicle(description: str, rating, coupon, has_rating: bool = True) -> str:
    """assumptions.md §6. Without a rating column (§6b) the blank-rating-and-
    coupon signature means nothing, so a line is "Direct line" only when its
    name carries a bond maturity date; single stocks and physical metal need
    an --instrument-overrides entry to be recognised as direct lines."""
    d = f" {(description or '').upper()} "
    if any(k in d for k in FUND_KEYWORDS):
        return "Fund"
    if any(k in d for k in STRUCTURED_KEYWORDS):
        return "Structured / AMC"
    if not has_rating:
        return "Direct line" if BOND_MATURITY_IN_NAME_RE.search(d) else "Fund"
    if (rating in (None, "", "Not Rated")) and coupon in (None, ""):
        return "Fund"
    return "Direct line"


def vehicle_of(row: dict, has_rating: bool, instrument_overrides: dict) -> str:
    return (instrument_override(row, instrument_overrides, "vehicle")
            or classify_vehicle(row.get("description"), row.get("composite rating"),
                                row.get("coupon (%)"), has_rating))


def equity_region_for(row: dict, has_geography: bool, instrument_overrides: dict) -> str:
    """assumptions.md §8: override, else the file's own geography column,
    else (French export, no geography column) a keyword read of the name."""
    forced = instrument_override(row, instrument_overrides, "region")
    if forced:
        return forced
    if has_geography:
        return equity_region_of(row.get("geographical breakdown"))
    return equity_region_from_name(row.get("description"))


def describe_bond_subsleeve(row: dict) -> str:
    """Cosmetic label only, assumptions.md §11."""
    desc = str(row.get("description") or "")
    sector = str(row.get("sector") or "").strip()
    du = desc.upper()
    keyword_map = [
        ("AUTO", "Corporate bond (auto)"),
        ("OIL", "Corporate bond (oil)"),
        ("PORT", "Corporate bond (port)"),
        ("AIRCRAFT", "Corporate bond (aircraft)"),
        ("AIR ", "Corporate bond (aircraft)"),
        ("COMMOD", "Corporate bond (commodity)"),
        ("GAMING", "Corporate bond (gaming)"),
        ("CASINO", "Corporate bond (gaming)"),
        ("TELE", "Corporate bond (telecom)"),
        ("REIT", "Corporate bond (REIT)"),
        ("BANK", "Bank bond"),
        ("INSUR", "Insurance-linked security"),
        ("ILS ", "Insurance-linked security"),
    ]
    for kw, label in keyword_map:
        if kw in du:
            return label
    is_fund = any(k in du for k in FUND_KEYWORDS)
    if is_fund:
        return "High yield bond fund"
    if "SUPRANATIONAL" in sector.upper() or "DEVELOPMENT" in du:
        return "Supranational bond"
    if any(k in du for k in STRUCTURED_KEYWORDS):
        return "Structured note / certificate"
    if sector.lower() in ("government", ""):
        return "Quasi-sovereign bond"
    return "Corporate bond"


def maturity_bucket(maturity_val, as_of=None) -> str:
    """Bucket a maturity date into the 0-3 / 3-5 / 5-7 / 7-10 / 10+ years bands
    used on the 'Proposed bond selection' slide (slide_recipe.md)."""
    import datetime as _dt

    as_of = as_of or _dt.date.today()
    if isinstance(maturity_val, _dt.datetime):
        maturity = maturity_val.date()
    elif isinstance(maturity_val, _dt.date):
        maturity = maturity_val
    elif isinstance(maturity_val, str) and maturity_val.strip():
        maturity = None
        for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d"):
            try:
                maturity = _dt.datetime.strptime(maturity_val.strip(), fmt).date()
                break
            except ValueError:
                continue
        if maturity is None:
            return "Unknown"
    else:
        return "Unknown"
    years = (maturity - as_of).days / 365.25
    if years <= 3:
        return "0 - 3 years"
    if years <= 5:
        return "3 - 5 years"
    if years <= 7:
        return "5 - 7 years"
    if years <= 10:
        return "7 - 10 years"
    return "10+ years"


def parse_proposed_bonds(wb) -> dict:
    """The 'Fixed Income' tab is a curated bond *proposal*, not part of the
    client's current holdings (assumptions.md §1) — every row equally
    weighted, matching the reference deck's 'N issues, equally weighted at
    X each' framing on the Proposed Bond Selection slide."""
    ws = wb["Fixed Income"]
    headers, rows_dict = {}, []
    header_row_idx = None
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(isinstance(v, str) and "isin" in v.lower() for v in vals):
            header_row_idx = r
            for c, v in enumerate(vals, 1):
                if v not in (None, ""):
                    headers[norm_header(v)] = c
            break
    if header_row_idx is None:
        return None
    for r in range(header_row_idx + 1, ws.max_row + 1):
        row = {h: ws.cell(r, c).value for h, c in headers.items()}
        if all(v in (None, "") for v in row.values()):
            continue
        rows_dict.append(row)

    n = len(rows_dict)
    if n == 0:
        return None
    equal_w = 100.0 / n
    amount_each = rows_dict[0].get("amount to invest")

    def bucket_sum(key_fn):
        d = defaultdict(float)
        for row in rows_dict:
            d[key_fn(row) or "Unknown"] += equal_w
        return {k: round(v, 4) for k, v in sorted(d.items(), key=lambda kv: -kv[1])}

    return {
        "num_issues": n,
        "amount_each": amount_each,
        "currency": rows_dict[0].get("currency"),
        "total_amount": amount_each * n if isinstance(amount_each, (int, float)) else None,
        "by_geography_pct": bucket_sum(lambda row: row.get("country")),
        "by_currency_pct": bucket_sum(lambda row: row.get("currency")),
        "by_maturity_pct": bucket_sum(lambda row: maturity_bucket(row.get("maturity"))),
        "by_rating_pct": bucket_sum(lambda row: row.get("rating")),
        "holdings": [
            {
                "name": row.get("company"),
                "isin": row.get("isin"),
                "country": row.get("country"),
                "sector": row.get("sector"),
                "rating": row.get("rating"),
                "currency": row.get("currency"),
                "coupon": row.get("coupon"),
                "yield": row.get("yield"),
                "maturity": str(row.get("maturity")),
            }
            for row in rows_dict
        ],
    }


def pct(x: float) -> float:
    # 6dp, not 4: a genuinely nonzero weight this small (e.g. a residual
    # CAD cash balance at ~0.00003%) must survive rounding as a small
    # positive number, not collapse to exactly 0.0 -- the build scripts'
    # display layer (fmt_pct) is what decides whether to show "<0.1%" for
    # it, and it can only do that if this function hasn't already erased
    # the distinction between "genuinely zero" and "just very small".
    return round(x * 100, 6)


def val_eur(row: dict) -> float:
    """Value in the file's base currency (see load_sheet_rows). The `_eur`
    name is historical: every `value_eur` / `total_value_eur` key in
    parsed.json is in `base_currency`, which is EUR only for EUR-based files."""
    v = row.get("_value_base")
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def weight(row: dict) -> float:
    w = row.get("weight (%)")
    try:
        return float(w or 0) / 100.0
    except (TypeError, ValueError):
        return 0.0


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("excel_path")
    ap.add_argument("--profile", required=True, choices=RISK_PROFILES)
    ap.add_argument("--client-name", default="[Client Name]")
    ap.add_argument("--valuation-date", default=None, help="e.g. '6 August 2026'; defaults to today")
    ap.add_argument("--bucket-overrides", default=None)
    ap.add_argument("--market-cap-overrides", default=None)
    ap.add_argument("--sector-overrides", default=None)
    ap.add_argument("--instrument-overrides", default=None,
                    help="CSV instrument,region,vehicle,fi_group (assumptions.md §14)")
    ap.add_argument("-o", "--output", default="parsed.json")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.excel_path, data_only=True)
    # The English export names its holdings tab "Portfolio"; the French
    # export ships a single unnamed "Sheet1" (assumptions.md §1b).
    holdings_ws = wb["Portfolio"] if "Portfolio" in wb.sheetnames else wb.worksheets[0]
    port_headers, port_rows, base_currency = load_sheet_rows(holdings_ws)
    # Which optional analytics columns this export actually carries. Every
    # rule that needs a missing column reports "not available" instead of
    # running on blanks (assumptions.md §1b).
    data_availability = {
        "isin": "isin code" in port_headers,
        "geography": "geographical breakdown" in port_headers,
        "rating": "composite rating" in port_headers,
        "duration": "duration" in port_headers,
        "yield": "yield (%)" in port_headers or "coupon (%)" in port_headers,
        "value_in_quote_currency": "valuation + accr. interest (in qc)" in port_headers,
    }
    has_rating = data_availability["rating"]
    has_geography = data_availability["geography"]
    proposed_bond_selection = parse_proposed_bonds(wb) if "Fixed Income" in wb.sheetnames else None

    bucket_overrides = load_bucket_overrides(args.bucket_overrides)
    mcap_overrides = load_market_cap_overrides(args.market_cap_overrides)
    sector_overrides = load_sector_overrides(args.sector_overrides)
    instrument_overrides = load_instrument_overrides(args.instrument_overrides)

    included, uncalled = [], []
    for row in port_rows:
        section = row.get("_section")
        if section not in PORTFOLIO_SECTIONS:
            continue
        w = weight(row)
        v = val_eur(row)
        if w <= 0 or v == 0:
            desc = row.get("description") or ""
            if "COMMIT" in str(desc).upper():
                uncalled.append({
                    "isin": row.get("isin code"),
                    "description": desc,
                    "committed_amount": row.get("balance"),
                    "currency": row.get("currency"),
                })
            continue
        row["_asset_class"] = ASSET_CLASS_MAP[section]
        row["_weight"] = w
        row["_value_eur"] = v
        if data_availability["value_in_quote_currency"]:
            try:
                row["_value_qc"] = float(row.get("valuation + accr. interest (in qc)") or 0)
            except (TypeError, ValueError):
                row["_value_qc"] = None
        else:
            # Not derivable without guessing each line's price convention
            # (bond % of nominal vs fund NAV per unit); tables fall back to
            # the base-currency value, labelled in the base currency.
            row["_value_qc"] = None
        included.append(row)

    total_value_eur = sum(r["_value_eur"] for r in included)

    # Base/reporting currency: read from the valuation column's own header
    # ("Valuation + accr. interest (EUR)", "Estimation + int. courus (USD)"),
    # matching the file's own stated portfolio total (assumptions.md §3).
    # This is NOT the same thing as "most-held currency by weight" (that's
    # currency_exposure_pct below, a property of the holdings, not of the
    # reporting base) -- conflating the two mislabels every aggregate figure
    # with whatever currency happens to dominate the book.
    ccy_val = defaultdict(float)
    for r in included:
        ccy_val[str(r.get("currency") or "Other")] += r["_value_eur"]
    dominant_holding_currency = max(ccy_val.items(), key=lambda kv: kv[1])[0] if ccy_val else base_currency

    # --- §5 asset allocation donut ---
    class_weight = defaultdict(float)
    for r in included:
        class_weight[r["_asset_class"]] += r["_weight"]
    asset_allocation = {k: pct(v) for k, v in sorted(class_weight.items(), key=lambda kv: -kv[1])}

    # --- currency exposure donut ---
    ccy_weight = defaultdict(float)
    for r in included:
        ccy_weight[str(r.get("currency") or "Other")] += r["_weight"]
    currency_exposure = {k: pct(v) for k, v in sorted(ccy_weight.items(), key=lambda kv: -kv[1])}

    # --- §7 geography / sector portfolio-wide ---
    # The Portfolio sheet carries no Sector column (unlike the separate
    # 'Fixed Income' proposal tab) -- sector exposure is only computable if
    # the caller supplies --sector-overrides; otherwise it is reported as
    # unavailable rather than fabricated. This matches the reference deck's
    # own disclosed caveat: "these fields are not in the client file."
    geo_weight = defaultdict(float)
    sector_weight = defaultdict(float)
    sector_data_available = bool(sector_overrides)
    for r in included:
        if has_geography:
            geo_weight[region_of(r.get("geographical breakdown"))] += r["_weight"]
        if sector_data_available:
            isin = str(r.get("isin code") or "")
            sector_weight[sector_overrides.get(isin, "Other")] += r["_weight"]
    # No geography column -> null, same "report unavailable, never fabricate"
    # rule as sector (assumptions.md §7).
    geographic_exposure = (
        {k: pct(v) for k, v in sorted(geo_weight.items(), key=lambda kv: -kv[1])}
        if has_geography else None
    )
    sector_exposure = (
        {k: pct(v) for k, v in sorted(sector_weight.items(), key=lambda kv: -kv[1])}
        if sector_data_available else None
    )

    # --- sleeves: top holdings + vehicle breakdown per asset class ---
    sleeves = {}
    for asset_class in sorted(set(r["_asset_class"] for r in included)):
        rows_ac = [r for r in included if r["_asset_class"] == asset_class]
        ac_total_w = sum(r["_weight"] for r in rows_ac)
        top = sorted(rows_ac, key=lambda r: -r["_weight"])[:5]
        vehicle_weight = defaultdict(float)
        for r in rows_ac:
            v = vehicle_of(r, has_rating, instrument_overrides)
            vehicle_weight[v] += r["_weight"] / ac_total_w if ac_total_w else 0
        durations = [float(r["duration"]) for r in rows_ac if isinstance(r.get("duration"), (int, float))]
        n_with_dur = len(durations)
        sleeves[asset_class] = {
            "total_weight_pct": pct(ac_total_w),
            "num_lines": len(rows_ac),
            "top_holdings": [
                {
                    "name": r.get("description"),
                    "isin": r.get("isin code"),
                    "value_eur": r["_value_eur"],
                    "value_qc": r["_value_qc"],
                    "currency": r.get("currency"),
                    "weight_pct": pct(r["_weight"]),
                    "duration": r.get("duration"),
                    "yield_pct": r.get("yield (%)"),
                    "price": r.get("price (qc)"),
                }
                for r in top
            ],
            "vehicle_breakdown_pct": {k: pct(v) for k, v in sorted(vehicle_weight.items(), key=lambda kv: -kv[1])},
            "duration_coverage": (f"{n_with_dur} of {len(rows_ac)}"
                                  if asset_class == "Fixed Income" and data_availability["duration"] else None),
        }

    # --- §8 equity breakdown (renormalized to the equity sleeve) ---
    equity_rows = [r for r in included if r["_asset_class"] == "Equities"]
    eq_total_w = sum(r["_weight"] for r in equity_rows) or 1.0
    eq_geo, eq_sector, eq_mcap = defaultdict(float), defaultdict(float), defaultdict(float)
    mcap_available = bool(mcap_overrides)
    for r in equity_rows:
        w_norm = r["_weight"] / eq_total_w
        eq_geo[equity_region_for(r, has_geography, instrument_overrides)] += w_norm
        if sector_data_available:
            isin = str(r.get("isin code") or "")
            eq_sector[sector_overrides.get(isin, "Other")] += w_norm
        if mcap_available:
            isin = str(r.get("isin code") or "")
            eq_mcap[mcap_overrides.get(isin, "Large cap")] += w_norm
    equity_breakdown = {
        "by_geography_pct": {k: pct(v) for k, v in sorted(eq_geo.items(), key=lambda kv: -kv[1])},
        "by_sector_pct": ({k: pct(v) for k, v in sorted(eq_sector.items(), key=lambda kv: -kv[1])} if sector_data_available else None),
        "by_market_cap_pct": ({k: pct(v) for k, v in sorted(eq_mcap.items(), key=lambda kv: -kv[1])} if mcap_available else None),
    }

    # --- §9 concentration & top holdings ---
    by_weight = sorted(included, key=lambda r: -r["_weight"])
    cum = 0.0
    concentration_rows = []
    for i, r in enumerate(by_weight[:20], 1):
        cum += r["_weight"]
        concentration_rows.append({
            "rank": i, "name": r.get("description"), "value_eur": r["_value_eur"],
            "weight_pct": pct(r["_weight"]), "cumulative_pct": pct(cum),
        })
    concentration = {
        "top_holdings": concentration_rows,
        "top_5_pct": pct(sum(r["_weight"] for r in by_weight[:5])),
        "top_10_pct": pct(sum(r["_weight"] for r in by_weight[:10])),
        "top_20_pct": pct(sum(r["_weight"] for r in by_weight[:20])),
        "positions_above_5pct": sum(1 for r in included if r["_weight"] > 0.05),
    }

    # --- §10 liquidity profile ---
    liq_weight = defaultdict(float)
    for r in included:
        desc = str(r.get("description") or "").upper()
        if r["_asset_class"] == "Private Assets" or "COMMIT" in desc:
            bucket = "Illiquid (lock-up)"
        elif r["_asset_class"] == "Cash":
            # §10 lists cash under "Listed"; without this branch a current
            # account (blank rating and coupon) fell through to the §6 fund
            # signature and was filed as a daily-liquid fund.
            bucket = "Listed"
        elif r["_asset_class"] == "Alternatives":
            # §10: the export says these are alternative funds but not their
            # dealing terms (daily UCITS vs gated quarterly), so they get
            # their own bucket rather than a guess either way.
            bucket = "Alternative funds (dealing terms vary)"
        elif vehicle_of(r, has_rating, instrument_overrides) == "Fund":
            bucket = "Daily-liquid fund"
        else:
            bucket = "Listed"
        liq_weight[bucket] += r["_weight"]
    liquidity_profile = {k: pct(v) for k, v in sorted(liq_weight.items(), key=lambda kv: -kv[1])}

    # --- §11 income & rate sensitivity ---
    fi_rows = [r for r in included if r["_asset_class"] == "Fixed Income"]
    running_yield_num, running_yield_den = 0.0, 0.0
    for r in included:
        y = r.get("yield (%)")
        if isinstance(y, (int, float)):
            running_yield_num += r["_weight"] * y
            running_yield_den += r["_weight"]
    running_yield_pct = round(running_yield_num, 4) if running_yield_den else None

    fi_total_val = sum(r["_value_eur"] for r in fi_rows)
    dur_num, dur_den = 0.0, 0.0
    duration_subsleeves = defaultdict(lambda: {"value": 0.0, "dur_weighted": 0.0})
    for r in fi_rows:
        d = r.get("duration")
        if isinstance(d, (int, float)):
            dur_num += r["_value_eur"] * d
            dur_den += r["_value_eur"]
        label = (describe_bond_subsleeve(r) if has_rating
                 else classify_fi_bucket(r, bucket_overrides, has_rating, instrument_overrides))
        duration_subsleeves[label]["value"] += r["_value_eur"]
        if isinstance(d, (int, float)):
            duration_subsleeves[label]["dur_weighted"] += r["_value_eur"] * d
    fi_duration = round(dur_num / dur_den, 2) if dur_den else None
    rate_impact_eur = round(-fi_duration * 0.01 * fi_total_val, 2) if fi_duration else None

    income_by_sleeve = defaultdict(float)
    for r in included:
        y = r.get("yield (%)")
        c = r.get("coupon (%)")
        rate = y if isinstance(y, (int, float)) else (c if isinstance(c, (int, float)) else None)
        if rate is not None:
            income_by_sleeve[r["_asset_class"]] += r["_value_eur"] * rate / 100.0
    total_income = sum(income_by_sleeve.values()) or 1.0

    income = {
        "running_yield_pct": running_yield_pct,
        "fi_duration_years": fi_duration,
        "rate_impact_100bp_eur": rate_impact_eur,
        "income_by_sleeve": [
            {"sleeve": k, "income_eur": round(v, 2), "pct_of_income": pct(v / total_income)}
            for k, v in sorted(income_by_sleeve.items(), key=lambda kv: -kv[1])
        ],
        "fi_duration_subsleeves": [
            {
                "label": k,
                "value_eur": round(v["value"], 2),
                "avg_duration": ((round(v["dur_weighted"] / v["value"], 2) if v["value"] and v["dur_weighted"] else 0.0)
                                 if data_availability["duration"] else None),
                "contribution": ((round(v["dur_weighted"] / fi_total_val, 2) if fi_total_val else 0.0)
                                 if data_availability["duration"] else None),
            }
            for k, v in sorted(duration_subsleeves.items(), key=lambda kv: -kv[1]["value"])
        ],
    }

    # --- §12 risk profile / growth-defensive split ---
    growth_w = sum(v for k, v in class_weight.items() if k in GROWTH_CLASSES)
    defensive_w = sum(v for k, v in class_weight.items() if k in DEFENSIVE_CLASSES)
    risk_profile = {
        "selected_profile": args.profile,
        "growth_assets_pct": pct(growth_w),
        "defensive_assets_pct": pct(defensive_w),
    }

    # --- line items table source (assumptions.md §4, slides 9-10) ---
    line_items = []
    fi_bucket_order = ["Govies 1-10 (local)", "Govies 10+ (local)", "High Yield (local or global hdg)",
                        "Corporate IG (local)", "EM Debt",
                        # §4b structure-based groups (exports without a rating column)
                        "Direct bonds", "Credit-linked notes", "Bond funds", "Money market & short-term"]
    equity_region_order = ["US", "Eurozone", "UK", "Switzerland", "Japan", "EM", "Global / thematic"]
    for r in included:
        ac = r["_asset_class"]
        if ac == "Fixed Income":
            group = classify_fi_bucket(r, bucket_overrides, has_rating, instrument_overrides)
        elif ac == "Cash":
            group = "Cash"
        elif ac == "Equities":
            group = equity_region_for(r, has_geography, instrument_overrides)
        else:
            group = ac
        line_items.append({
            "top_category": ac,
            "group": group,
            "isin": r.get("isin code"),
            "currency": r.get("currency"),
            "instrument": r.get("description"),
            "weight_pct": pct(r["_weight"]),
        })
    # stable sort: top_category (Cash, Fixed Income, Equities, Alternatives,
    # Structured Products, Private Assets, Commodities, Other), then FI bucket
    # order, then descending weight within group.
    top_cat_order = ["Cash", "Fixed Income", "Equities", "Alternatives", "Structured Products",
                      "Private Assets", "Commodities", "Other"]

    def sort_key(li):
        tc = top_cat_order.index(li["top_category"]) if li["top_category"] in top_cat_order else 99
        if li["top_category"] == "Fixed Income":
            gi = fi_bucket_order.index(li["group"]) if li["group"] in fi_bucket_order else 99
        elif li["top_category"] == "Equities":
            gi = equity_region_order.index(li["group"]) if li["group"] in equity_region_order else 99
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
        "base_currency": base_currency,
        "data_availability": data_availability,
        "dominant_holding_currency": dominant_holding_currency,
        "total_value_eur": round(total_value_eur, 2),
        "num_positions": len(included),
        "largest_position_pct": pct(by_weight[0]["_weight"]) if by_weight else 0,
        "liquid_share_pct": pct(liq_weight.get("Listed", 0) + liq_weight.get("Daily-liquid fund", 0)),
        "asset_allocation_pct": asset_allocation,
        "currency_exposure_pct": currency_exposure,
        "geographic_exposure_pct": geographic_exposure,
        "sector_exposure_pct": sector_exposure,
        "sleeves": sleeves,
        "equity_breakdown": equity_breakdown,
        "concentration": concentration,
        "liquidity_profile_pct": liquidity_profile,
        "income": income,
        "risk_profile": risk_profile,
        "uncalled_commitments": uncalled,
        "line_items": line_items,
        "proposed_bond_selection": proposed_bond_selection,
    }

    Path(args.output).write_text(json.dumps(parsed, indent=2, default=str))
    print(f"Wrote {args.output}: {len(included)} positions, "
          f"{parsed['total_value_eur']:,.0f} {base_currency}, dominant holding currency {dominant_holding_currency}")


if __name__ == "__main__":
    main()
