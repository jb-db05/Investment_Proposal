#!/usr/bin/env python3
"""
Turn a client's CURRENT custodian export plus an advisor's rebalance spec
into (a) a PROPOSED portfolio Excel in the same custodian format, which
build_proposal.py then turns into the deck, and (b) rebalance.json, the
before/after statistics and trade list that the deck's "recommended
changes" slides are built from (build_rebalance_slides.py).

    python build_rebalance.py current.xlsx rebalance_spec.json \
        --out-excel proposed.xlsx --out-json rebalance.json

The spec (see proposals/*/rebalance_spec.json for a worked example) is a
judgment input, like market_update.json: which lines to sell, which to
resize to a target weight, which new funds to add, and how each equity line
is tagged (sector, region, AI camp, vehicle). This script only does the
arithmetic, deterministically:

- Portfolio total is unchanged (transaction costs and taxes ignored).
- `sell` removes the line; its value is credited to base-currency cash.
- `resize` scales balance and valuations to `target_weight_pct` of the total.
- New positions are valued at `target_weight_pct` of the total; their
  quote-currency value uses the FX rate implied by the file itself
  (base valuation / QC valuation of existing lines in that currency).
  Unit counts are left blank: they depend on the execution price.
- Whatever is sold and not reinvested lands in the base-currency current
  account. Every weight is recomputed as valuation / total.

Statistics are computed on the `Equities` section only (the same scope as
the deck's Equities sleeve slides).
"""
from __future__ import annotations

import argparse
import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import parse_portfolio as pp  # noqa: E402

EQUITY_SECTION = "Equities"
CASH_SECTION = "Cash"


def read_raw(ws):
    """Return (header_row_values, list of (section, raw_row_list)) keeping
    the original column order, so the proposed file mirrors the export."""
    header_idx = None
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in vals if v not in (None, "")]
        if len(non_empty) >= 5 and any(isinstance(v, str) and "isin" in v.lower() for v in vals):
            header_idx = r
            header = vals
            break
    if header_idx is None:
        raise ValueError("no ISIN header row")
    rows, section, order = [], None, []
    for r in range(header_idx + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        a = vals[0]
        if isinstance(a, str) and a.strip() in pp.PORTFOLIO_SECTIONS:
            section = a.strip()
            order.append(section)
            continue
        if all(v in (None, "") for v in vals):
            continue
        rows.append([section, vals])
    return header, rows, order


def col_index(header, normalized_name_prefix):
    for i, h in enumerate(header):
        if pp.norm_header(h).startswith(normalized_name_prefix):
            return i
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("current_excel")
    ap.add_argument("spec_json")
    ap.add_argument("--out-excel", required=True)
    ap.add_argument("--out-json", required=True)
    ap.add_argument("--out-mcap-csv", default=None,
                    help="also write isin,market_cap for proposed equity lines (parse_portfolio --market-cap-overrides); "
                         "single stocks use their 'market_cap' tag, funds/certificates are 'Funds (multi-cap)'")
    ap.add_argument("--out-sector-csv", default=None,
                    help="also write isin,sector for every proposed line (parse_portfolio --sector-overrides)")
    args = ap.parse_args()

    spec = json.loads(Path(args.spec_json).read_text())
    wb = openpyxl.load_workbook(args.current_excel, data_only=True)
    ws = pp.find_portfolio_sheet(wb)
    header, rows, section_order = read_raw(ws)
    norm = {pp.norm_header(h): i for i, h in enumerate(header) if h not in (None, "")}
    val_hdr, base_ccy = pp.detect_base_currency(norm)

    C = {
        "isin": norm["isin code"], "desc": norm["description"], "bal": norm["balance"],
        "val": norm[val_hdr], "qc": norm["valuation + accr. interest (in qc)"],
        "ccy": norm["currency"], "w": norm["weight (%)"], "geo": norm["geographical breakdown"],
        "rating": norm.get("composite rating"), "pl": col_index(header, "unrealized p / l total"),
        "cost": col_index(header, "total cost price"), "perf": col_index(header, "tw perf. (" + base_ccy.lower()),
        "last_op": norm.get("last operation date"),
    }

    def num(x):
        return float(x) if isinstance(x, (int, float)) else 0.0

    total = sum(num(v[C["val"]]) for _, v in rows)

    # FX implied by the file itself: base-ccy value per 1 unit of quote ccy.
    fx_samples = defaultdict(list)
    for _, v in rows:
        b, q = num(v[C["val"]]), num(v[C["qc"]])
        if b and q:
            fx_samples[v[C["ccy"]]].append(b / q)
    fx = {k: statistics.median(s) for k, s in fx_samples.items()}
    fx[base_ccy] = 1.0

    tags = spec.get("tags", {})
    actions = {a["isin"]: a for a in spec.get("actions", [])}
    unknown = set(actions) - {v[C["isin"]] for _, v in rows}
    if unknown:
        sys.exit(f"spec references ISINs not in the file: {sorted(unknown)}")

    def equity_snapshot(eq_rows):
        """eq_rows: list of dicts {isin, name, value, tag}"""
        eq_total = sum(r["value"] for r in eq_rows)
        thr = spec.get("single_line_min_weight_pct", 2.0)

        def agg(key):
            d = defaultdict(float)
            for r in eq_rows:
                d[r["tag"].get(key, "Unclassified")] += r["value"]
            return {k: round(v / total * 100, 4) for k, v in sorted(d.items(), key=lambda kv: -kv[1])}

        by_val = sorted(eq_rows, key=lambda r: -r["value"])
        singles = [r for r in eq_rows if r["tag"].get("vehicle") == "Single stock"]
        return {
            "equity_weight_pct": round(eq_total / total * 100, 4),
            "num_lines": len(eq_rows),
            "num_single_stocks": len(singles),
            "num_funds": sum(1 for r in eq_rows if r["tag"].get("vehicle") == "Fund / ETF"),
            "num_certificates": sum(1 for r in eq_rows if r["tag"].get("vehicle") == "Certificate"),
            "num_lines_below_threshold": sum(1 for r in eq_rows if r["value"] / total * 100 < thr),
            "threshold_pct": thr,
            "single_stock_weight_pct": round(sum(r["value"] for r in singles) / total * 100, 4),
            "ai_exposed_pct_of_portfolio": round(sum(r["value"] for r in eq_rows if r["tag"].get("ai") == "AI-exposed") / total * 100, 4),
            "ai_exposed_pct_of_equities": round(sum(r["value"] for r in eq_rows if r["tag"].get("ai") == "AI-exposed") / eq_total * 100, 4),
            "top5_pct_of_equities": round(sum(r["value"] for r in by_val[:5]) / eq_total * 100, 4),
            "by_sector_pct": agg("sector"),
            "by_region_pct": agg("region"),
            "by_ai_pct": agg("ai"),
            "by_vehicle_pct": agg("vehicle"),
        }

    def eq_rows_of(rowlist):
        out = []
        for sec, v in rowlist:
            if sec != EQUITY_SECTION:
                continue
            isin = v[C["isin"]]
            if isin not in tags:
                sys.exit(f"equity line {isin} ({v[C['desc']]}) has no tag in the spec")
            out.append({"isin": isin, "name": v[C["desc"]], "value": num(v[C["val"]]), "tag": tags[isin]})
        return out

    before_eq = equity_snapshot(eq_rows_of(rows))
    cash_before = sum(num(v[C["val"]]) for s, v in rows if s == CASH_SECTION)

    # --- apply actions ---
    sells, resizes, buys = [], [], []
    new_rows = []
    for sec, v in rows:
        isin = v[C["isin"]]
        a = actions.get(isin)
        old_val = num(v[C["val"]])
        if a is None:
            new_rows.append([sec, list(v)])
            continue
        if a["action"] == "sell":
            sells.append({
                "isin": isin, "name": v[C["desc"]], "section": sec,
                "value": round(old_val, 2), "weight_pct": round(old_val / total * 100, 4),
                "unrealized_pl": num(v[C["pl"]]) if C["pl"] is not None else None,
                "tag": tags.get(isin, {}), "reason": a.get("reason", ""), "group": a.get("group"),
            })
            continue
        if a["action"] == "resize":
            new_val = a["target_weight_pct"] / 100 * total
            f = new_val / old_val
            nv = list(v)
            for k in ("bal", "val", "qc", "pl", "cost"):
                if C[k] is not None and isinstance(nv[C[k]], (int, float)):
                    nv[C[k]] = nv[C[k]] * f
            new_rows.append([sec, nv])
            resizes.append({
                "isin": isin, "name": v[C["desc"]], "section": sec,
                "value_before": round(old_val, 2), "value_after": round(new_val, 2),
                "change": round(new_val - old_val, 2),
                "weight_before_pct": round(old_val / total * 100, 4), "weight_after_pct": a["target_weight_pct"],
                "unrealized_pl_realized": (round(num(v[C["pl"]]) * (1 - f), 2) if C["pl"] is not None and f < 1 else None),
                "tag": tags.get(isin, {}), "reason": a.get("reason", ""), "group": a.get("group"),
            })
            continue
        sys.exit(f"unknown action {a['action']!r} for {isin}")

    for np_ in spec.get("new_positions", []):
        val = np_["target_weight_pct"] / 100 * total
        ccy = np_["currency"]
        if ccy not in fx:
            sys.exit(f"no FX rate for {ccy} can be implied from the file")
        nv = [None] * len(header)
        nv[C["isin"]] = np_["isin"]
        nv[C["desc"]] = np_["description"]
        nv[C["val"]] = val
        nv[C["qc"]] = val / fx[ccy]
        nv[C["ccy"]] = ccy
        nv[C["geo"]] = np_["geography"]
        if C["rating"] is not None:
            nv[C["rating"]] = "Not Rated"
        if C["last_op"] is not None:
            nv[C["last_op"]] = "PROPOSED"
        new_rows.append([EQUITY_SECTION, nv])
        tags[np_["isin"]] = {k: np_[k] for k in ("sector", "region", "ai", "vehicle")}
        buys.append({"isin": np_["isin"], "name": np_["description"], "currency": ccy,
                     "value": round(val, 2), "weight_pct": np_["target_weight_pct"],
                     "tag": tags[np_["isin"]], "reason": np_.get("reason", ""), "new": True})

    # cash residual
    residual = total - sum(num(v[C["val"]]) for _, v in new_rows)
    cash_row = next((v for s, v in new_rows if s == CASH_SECTION and v[C["ccy"]] == base_ccy
                     and "CURRENT ACCOUNT" in str(v[C["desc"]]).upper()), None)
    if cash_row is None:
        cash_row = [None] * len(header)
        cash_row[C["desc"]] = f"CURRENT ACCOUNT IN {base_ccy}"
        cash_row[C["ccy"]] = base_ccy
        new_rows.insert(0, [CASH_SECTION, cash_row])
    cash_row[C["val"]] = num(cash_row[C["val"]]) + residual
    cash_row[C["qc"]] = cash_row[C["val"]]
    cash_row[C["bal"]] = cash_row[C["val"]]

    for _, v in new_rows:
        v[C["w"]] = num(v[C["val"]]) / total * 100

    after_eq = equity_snapshot(eq_rows_of(new_rows))
    cash_after = sum(num(v[C["val"]]) for s, v in new_rows if s == CASH_SECTION)

    # Keep resized rows that are top-ups in the "buys" list too, for the slide.
    topups = [r for r in resizes if r["change"] > 0]
    trims = [r for r in resizes if r["change"] < 0]

    rebalance = {
        "as_of": spec.get("as_of"),
        "house_view_source": spec.get("house_view_source"),
        "base_currency": base_ccy,
        "total_value": round(total, 2),
        "objectives": spec.get("objectives", []),
        "sale_groups": spec.get("sale_groups", []),
        "fx_implied": {k: round(v, 6) for k, v in fx.items()},
        "cash_before_pct": round(cash_before / total * 100, 4),
        "cash_after_pct": round(cash_after / total * 100, 4),
        "equities_before": before_eq,
        "equities_after": after_eq,
        "sells": sorted(sells, key=lambda r: -r["value"]),
        "trims": trims,
        "topups": sorted(topups, key=lambda r: -r["change"]),
        "new_positions": buys,
        "gross_sales": round(sum(r["value"] for r in sells) - sum(r["change"] for r in trims), 2),
        "gross_purchases": round(sum(r["value"] for r in buys) + sum(r["change"] for r in topups), 2),
        "unrealized_pl_on_sales": round(sum((r["unrealized_pl"] or 0) for r in sells)
                                        + sum((r["unrealized_pl_realized"] or 0) for r in trims), 2),
    }
    Path(args.out_json).write_text(json.dumps(rebalance, indent=2, default=str))

    # --- write the proposed Excel in the custodian layout ---
    out = openpyxl.Workbook()
    ws_out = out.active
    ws_out.title = "Portfolio"
    n_pos = sum(1 for s, v in new_rows if num(v[C["val"]]) != 0)
    total_str = f"{total:,.0f}".replace(",", "'")
    ws_out.append([f"PROPOSED PORTFOLIO as at {spec.get('as_of', '')} - Total : {total_str} {base_ccy} "
                   f"Non-zero-balance positions - Segment Total ({n_pos})"])
    ws_out.append(header)
    for c in ws_out[2]:
        c.font = Font(bold=True)
    for sec in section_order:
        sec_rows = [v for s, v in new_rows if s == sec]
        if not sec_rows:
            continue
        ws_out.append([sec])
        ws_out.cell(ws_out.max_row, 1).font = Font(bold=True)
        for v in sec_rows:
            ws_out.append(v)
            if v[C["last_op"]] == "PROPOSED" if C["last_op"] is not None else False:
                for c in ws_out[ws_out.max_row]:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")

    # Trades sheet
    wt = out.create_sheet("Trades")
    wt.append(["Action", "ISIN", "Instrument", "Sector", "Region", "AI camp",
               f"Weight before (%)", f"Weight after (%)", f"Trade ({base_ccy})",
               f"Unrealised P/L realised ({base_ccy})", "Rationale"])
    for c in wt[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="202945")
    for r in rebalance["sells"]:
        t = r["tag"]
        wt.append(["Sell", r["isin"], r["name"], t.get("sector"), t.get("region"), t.get("ai"),
                   r["weight_pct"], 0, -r["value"], r["unrealized_pl"], r["reason"]])
    for r in trims:
        t = r["tag"]
        wt.append(["Trim", r["isin"], r["name"], t.get("sector"), t.get("region"), t.get("ai"),
                   r["weight_before_pct"], r["weight_after_pct"], r["change"], r["unrealized_pl_realized"], r["reason"]])
    for r in rebalance["topups"]:
        t = r["tag"]
        wt.append(["Top up", r["isin"], r["name"], t.get("sector"), t.get("region"), t.get("ai"),
                   r["weight_before_pct"], r["weight_after_pct"], r["change"], None, r["reason"]])
    for r in buys:
        t = r["tag"]
        wt.append(["Buy (new)", r["isin"], r["name"], t.get("sector"), t.get("region"), t.get("ai"),
                   0, r["weight_pct"], r["value"], None, r["reason"]])
    wt.append(["Cash", "", f"CURRENT ACCOUNT IN {base_ccy} (residual)", "", "", "",
               rebalance["cash_before_pct"], rebalance["cash_after_pct"],
               round((rebalance["cash_after_pct"] - rebalance["cash_before_pct"]) / 100 * total, 2), None,
               "Proceeds not reinvested: cash buffer to redeploy in stages."])
    for row in wt.iter_rows(min_row=2):
        for c in row[6:10]:
            c.number_format = '#,##0.00'
    for col, w in zip("ABCDEFGHIJK", (10, 15, 44, 22, 13, 13, 12, 12, 14, 16, 90)):
        wt.column_dimensions[col].width = w

    # Before vs after sheet
    wb_ = out.create_sheet("Before vs After")
    wb_.append(["Equity sleeve metric", "Current", "Proposed"])
    for c in wb_[1]:
        c.font = Font(bold=True)
    b, a = before_eq, after_eq
    for label, key in [("Equity weight (% of portfolio)", "equity_weight_pct"),
                       ("Number of equity lines", "num_lines"),
                       ("  of which single stocks", "num_single_stocks"),
                       ("  of which funds / ETFs", "num_funds"),
                       ("  of which certificates", "num_certificates"),
                       (f"Lines below {b['threshold_pct']}% of portfolio", "num_lines_below_threshold"),
                       ("Single stocks (% of portfolio)", "single_stock_weight_pct"),
                       ("AI-exposed (% of portfolio)", "ai_exposed_pct_of_portfolio"),
                       ("AI-exposed (% of equity sleeve)", "ai_exposed_pct_of_equities"),
                       ("Top 5 lines (% of equity sleeve)", "top5_pct_of_equities")]:
        wb_.append([label, b[key], a[key]])
    wb_.append(["Cash (% of portfolio)", rebalance["cash_before_pct"], rebalance["cash_after_pct"]])
    for dim, key in [("Sector", "by_sector_pct"), ("Region", "by_region_pct"), ("AI camp", "by_ai_pct"),
                     ("Vehicle", "by_vehicle_pct")]:
        wb_.append([])
        wb_.append([f"{dim} (% of portfolio)", "Current", "Proposed"])
        for c in wb_[wb_.max_row]:
            c.font = Font(bold=True)
        for k in sorted(set(b[key]) | set(a[key]), key=lambda k: -(a[key].get(k, 0) + b[key].get(k, 0))):
            wb_.append([k, b[key].get(k, 0), a[key].get(k, 0)])
    for row in wb_.iter_rows(min_row=2):
        for c in row[1:]:
            if isinstance(c.value, float):
                c.number_format = '0.0'
    wb_.column_dimensions["A"].width = 38
    out.save(args.out_excel)

    if args.out_mcap_csv:
        with open(args.out_mcap_csv, "w") as f:
            f.write("isin,market_cap\n")
            for sec, v in new_rows:
                isin = v[C["isin"]]
                if sec != EQUITY_SECTION or not isin:
                    continue
                t = tags.get(isin, {})
                mcap = t.get("market_cap") or ("Funds (multi-cap)" if t.get("vehicle") != "Single stock" else "Unclassified")
                f.write(f"{isin},{mcap}\n")

    if args.out_sector_csv:
        section_default = {"Cash": "Cash", "Short-term instruments": "Fixed income", "Bonds": "Fixed income"}
        with open(args.out_sector_csv, "w") as f:
            f.write("isin,sector\n")
            for sec, v in new_rows:
                isin = v[C["isin"]]
                if not isin:
                    continue
                sector = tags.get(isin, {}).get("sector") or section_default.get(sec, "Other")
                f.write(f"{isin},{sector}\n")

    print(f"Total {base_ccy} {total:,.0f}. Equity lines {b['num_lines']} -> {a['num_lines']}, "
          f"single stocks {b['num_single_stocks']} -> {a['num_single_stocks']}, "
          f"AI-exposed {b['ai_exposed_pct_of_equities']:.1f}% -> {a['ai_exposed_pct_of_equities']:.1f}% of equities, "
          f"cash {rebalance['cash_before_pct']:.1f}% -> {rebalance['cash_after_pct']:.1f}%")
    print(f"Wrote {args.out_excel} and {args.out_json}")


if __name__ == "__main__":
    main()
