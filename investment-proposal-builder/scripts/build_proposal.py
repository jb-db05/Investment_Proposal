#!/usr/bin/env python3
"""
Main orchestrator: portfolio Excel + market PDF summary + risk profile ->
a fully populated Investment Proposal deck.

    python build_proposal.py \
        --excel Portfolio.xlsx --profile Balanced \
        --client-name "Jane Doe" \
        --market-update market_update.json \
        -o "Investment Proposal - Jane Doe.pptx"

What this script does NOT do: read the market PDF itself. Extracting an
arbitrary week's commentary into structured slide content is a
summarization/judgment task, not a deterministic transform — see
SKILL.md and references/slide_recipe.md for how Claude produces
market_update.json from the PDF before this script runs. Everything
downstream of that JSON (and of parsed.json from parse_portfolio.py) is
a deterministic, no-invented-numbers fill of the template.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import build_line_items_tables as line_items_mod  # noqa: E402
import update_chart as chart_mod  # noqa: E402

TEMPLATE_PATH = SCRIPT_DIR.parent / "assets" / "template.pptx"

RISK_PROFILE_CHARACTERISTICS = {
    "Fixed Income":  {"horizon": "Short term (1 to 3 years)",       "liquidity": "High liquidity, capital preservation priority"},
    "Conservative":  {"horizon": "Short to medium term (3 to 5 years)", "liquidity": "Regular distribution, stable capital"},
    "Moderate":      {"horizon": "Medium to long term (5 to 10 years)", "liquidity": "Regular distribution but stable capital"},
    "Balanced":      {"horizon": "Medium to long term (5 to 10 years)", "liquidity": "Balanced between growth and distribution"},
    "Growth":        {"horizon": "Long term (7 to 12 years)",       "liquidity": "Growth priority, distributions opportunistic"},
    "Equity":        {"horizon": "Long term (10+ years)",           "liquidity": "Full growth priority, minimal liquidity needs"},
}
# assumptions.md §12 / slide8_profile_dial.md: growth-asset band each profile is centered on.
RISK_PROFILE_GROWTH_BAND = {
    "Fixed Income": (0, 10), "Conservative": (10, 25), "Moderate": (25, 45),
    "Balanced": (45, 65), "Growth": (65, 85), "Equity": (85, 100),
}


def find_shape(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    return None


def set_plain_text(slide, name, text):
    """Simple, robust single/multi-paragraph text replacement: keeps the
    first paragraph's run formatting for every line, drops the rest."""
    sh = find_shape(slide, name)
    if sh is None or not sh.has_text_frame:
        return False
    tf = sh.text_frame
    lines = str(text).split("\n")
    # Set first paragraph text via its first run (preserves that run's style)
    p0 = tf.paragraphs[0]
    for run in list(p0.runs)[1:]:
        run._r.getparent().remove(run._r)
    if not p0.runs:
        p0.add_run()
    p0.runs[0].text = lines[0]

    # remove all paragraphs after the first, then re-add clones of paragraph 0
    # (with its formatting) for each additional line
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    import copy as _copy
    anchor = tf.paragraphs[0]._p
    for line in lines[1:]:
        clone = _copy.deepcopy(anchor)
        anchor.addnext(clone)
        anchor = clone
    # now walk paragraphs in order and assign remaining lines
    for para, line in zip(tf.paragraphs, lines):
        if not para.runs:
            para.add_run()
        for run in list(para.runs)[1:]:
            run._r.getparent().remove(run._r)
        para.runs[0].text = line
    return True


def replace_token(slide, old: str, new: str):
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for run in p.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)


def fmt_money_m(value_eur: float, currency: str) -> str:
    return f"{currency} {value_eur / 1_000_000:.1f}m"


def fmt_money_k(value: float, currency: str) -> str:
    return f"{currency} {value / 1_000:.0f}k"


def get_slide(prs, index_1based):
    return list(prs.slides)[index_1based - 1]


# ---------------------------------------------------------------- charts ---

def update_donut_by_name(slide, chart_name, data, label_style="whole"):
    sh = find_shape(slide, chart_name)
    if sh is None or not sh.has_chart:
        return
    chart_mod.update_donut(sh, data, label_style=label_style)


def update_bar_by_name(slide, chart_name, data):
    sh = find_shape(slide, chart_name)
    if sh is None or not sh.has_chart:
        return
    chart_mod.update_bar(sh, data)


# ----------------------------------------------------------------- tables --

HEADER_FILL = RGBColor(0x1E, 0x27, 0x61)
HEADER_FONT = RGBColor(0xFF, 0xFF, 0xFF)


def _cell(cell, text, *, bold=False, fill=None, font_color=RGBColor(0x22, 0x22, 0x22), size=10, align=PP_ALIGN.LEFT):
    cell.margin_left = Emu(45720)
    cell.margin_right = Emu(45720)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    if not p.runs:
        p.add_run()
    for run in list(p.runs)[1:]:
        run._r.getparent().remove(run._r)
    p.runs[0].text = str(text)
    p.runs[0].font.size = Pt(size)
    p.runs[0].font.bold = bold
    p.runs[0].font.color.rgb = font_color


def rebuild_holdings_table(slide, table_shape_name, holdings, currency_fallback="USD"):
    """Replace the sleeve's 'largest holdings' table with one sized to the
    actual number of top holdings (assumptions.md doesn't fabricate empty
    rows, and the reference template's tables are too small for a sleeve
    with more than 1-5 real holdings)."""
    sh = find_shape(slide, table_shape_name)
    if sh is None or not sh.has_table:
        return
    left, top, width, height = sh.left, sh.top, sh.width, sh.height
    sh._element.getparent().remove(sh._element)

    n_rows = 1 + max(1, len(holdings))
    graphic_frame = slide.shapes.add_table(n_rows, 4, left, top, width, height)
    table = graphic_frame.table
    col_w = [int(width * f) for f in (0.46, 0.18, 0.14, 0.22)]
    for i, w in enumerate(col_w):
        table.columns[i].width = Emu(w)
    headers = ["Holding", "Value", "Share", "Yield / Coupon"]
    for c, h in enumerate(headers):
        _cell(table.cell(0, c), h, bold=True, fill=HEADER_FILL, font_color=HEADER_FONT, size=9.5)
    if not holdings:
        _cell(table.cell(1, 0), "No holdings in this sleeve", size=9)
        for c in range(1, 4):
            _cell(table.cell(1, c), "")
        return
    for r, h in enumerate(holdings, 1):
        ccy = h.get("currency") or currency_fallback
        _cell(table.cell(r, 0), h.get("name") or "", size=9)
        _cell(table.cell(r, 1), fmt_money_k(h.get("value_qc") or h.get("value_eur") or 0, ccy), size=9)
        _cell(table.cell(r, 2), f"{h['weight_pct']:.1f}%", size=9, align=PP_ALIGN.RIGHT)
        yld = h.get("yield_pct")
        _cell(table.cell(r, 3), (f"{yld:.2f}%" if isinstance(yld, (int, float)) else "—"), size=9, align=PP_ALIGN.RIGHT)


def fill_concentration_table(slide, table_shape_name, rows, currency="EUR"):
    sh = find_shape(slide, table_shape_name)
    if sh is None or not sh.has_table:
        return
    table = sh.table
    max_rows = len(table.rows) - 1
    for i in range(max_rows):
        r = i + 1
        if i < len(rows):
            row = rows[i]
            _cell(table.cell(r, 0), str(row["rank"]), size=9)
            _cell(table.cell(r, 1), row["name"] or "", size=9)
            _cell(table.cell(r, 2), f"{row['value_eur']:,.0f}", size=9, align=PP_ALIGN.RIGHT)
            _cell(table.cell(r, 3), f"{row['weight_pct']:.1f}%", size=9, align=PP_ALIGN.RIGHT)
            _cell(table.cell(r, 4), f"{row['cumulative_pct']:.1f}%", size=9, align=PP_ALIGN.RIGHT)
        else:
            for c in range(5):
                _cell(table.cell(r, c), "")


def fill_income_tables(slide, sleeve_table_name, duration_table_name, income):
    sh = find_shape(slide, sleeve_table_name)
    if sh is not None and sh.has_table:
        table = sh.table
        rows = income["income_by_sleeve"]
        for i in range(len(table.rows) - 1):
            r = i + 1
            if i < len(rows):
                row = rows[i]
                _cell(table.cell(r, 0), row["sleeve"], size=9)
                _cell(table.cell(r, 1), f"{row['income_eur']:,.0f}", size=9, align=PP_ALIGN.RIGHT)
                _cell(table.cell(r, 2), f"{row['pct_of_income']:.1f}%", size=9, align=PP_ALIGN.RIGHT)
            else:
                for c in range(3):
                    _cell(table.cell(r, c), "")

    sh2 = find_shape(slide, duration_table_name)
    if sh2 is not None and sh2.has_table:
        left, top, width, height = sh2.left, sh2.top, sh2.width, sh2.height
        sh2._element.getparent().remove(sh2._element)
        rows = income["fi_duration_subsleeves"]
        n_rows = 1 + max(1, len(rows))
        gf = slide.shapes.add_table(n_rows, 4, left, top, width, height)
        table = gf.table
        headers = ["Sub-sleeve", "Value", "Avg duration", "Contribution"]
        for c, h in enumerate(headers):
            _cell(table.cell(0, c), h, bold=True, fill=HEADER_FILL, font_color=HEADER_FONT, size=9)
        for r, row in enumerate(rows, 1):
            _cell(table.cell(r, 0), row["label"], size=8.5)
            _cell(table.cell(r, 1), f"{row['value_eur']:,.0f}", size=8.5, align=PP_ALIGN.RIGHT)
            _cell(table.cell(r, 2), f"{row['avg_duration']:.1f}", size=8.5, align=PP_ALIGN.RIGHT)
            _cell(table.cell(r, 3), f"{row['contribution']:.2f}", size=8.5, align=PP_ALIGN.RIGHT)


# ------------------------------------------------------------- slide fills -

def fill_cover(prs, parsed):
    slide = get_slide(prs, 1)
    ccy = parsed["base_currency"]
    amount_m = parsed["total_value_eur"] / 1_000_000
    set_plain_text(slide, "Text Placeholder 5",
                    f"{parsed['risk_profile_input']} {ccy} {amount_m:.1f}m\n\n"
                    f"Prepared for {parsed['client_name']}\n\n"
                    f"Your Relationship Manager: [Name]\n\n"
                    f"Your Advisor: ")


def fill_profile_dial(prs, parsed):
    slide = get_slide(prs, 8)
    profile = parsed["risk_profile_input"]
    growth_pct = parsed["risk_profile"]["growth_assets_pct"]
    set_plain_text(slide, "TextBox 23", f"Proposed profile for this portfolio: {profile} - {growth_pct:.0f}% in growth assets")
    chars = RISK_PROFILE_CHARACTERISTICS[profile]
    sh = find_shape(slide, "Table 44")
    if sh is not None and sh.has_table:
        table = sh.table
        mapping = {
            "Investment horizon": chars["horizon"],
            "Reference of currency": parsed["base_currency"],
            "Risk level": profile,
            "Liquidity needs": chars["liquidity"],
        }
        for row in table.rows:
            label = row.cells[0].text.strip()
            if label in mapping:
                _cell(row.cells[1], mapping[label], size=10)


def fill_portfolio_overview(prs, parsed):
    slide = get_slide(prs, 11)
    ccy = parsed["base_currency"]
    set_plain_text(slide, "Rectangle 3", f"{ccy} {parsed['total_value_eur']/1_000_000:.1f}M\nTotal value")
    set_plain_text(slide, "Rectangle 4", f"{parsed['num_positions']}\nPositions")
    set_plain_text(slide, "Rectangle 5", f"{parsed['largest_position_pct']:.1f}%\nLargest position")
    set_plain_text(slide, "Rectangle 6", f"{parsed['liquid_share_pct']:.1f}%\nLiquid share")
    ry = parsed["income"]["running_yield_pct"]
    set_plain_text(slide, "Rectangle 7", f"{ry:.1f}%\nRunning yield" if ry is not None else "n/a\nRunning yield")
    update_donut_by_name(slide, "Chart 10", parsed["asset_allocation_pct"])
    update_donut_by_name(slide, "Chart 14", parsed["currency_exposure_pct"])


def set_donut_legend(slide, textbox_names, data: dict[str, float]):
    """Fill the fixed set of companion 'legend' textboxes next to a donut
    with one-decimal percent labels, in data order. If there are more/fewer
    categories than textboxes (common — the template was sized for one
    example's category count), fill what fits and blank the rest rather
    than leaving stale numbers from the reference client behind."""
    items = list(data.items())
    for i, name in enumerate(textbox_names):
        text = chart_mod.pct_label_one_decimal(*items[i]) if i < len(items) else ""
        set_plain_text(slide, name, text)


def fill_geo_sector(prs, parsed):
    slide = get_slide(prs, 12)
    update_bar_by_name(slide, "Chart 5", parsed["geographic_exposure_pct"])
    if parsed["sector_exposure_pct"]:
        update_bar_by_name(slide, "Chart 8", parsed["sector_exposure_pct"])
    else:
        set_plain_text(slide, "TextBox 6", "Sector exposure\n(not available: the custodian file carries no "
                                            "sector field for these holdings — see references/assumptions.md §7)")
        update_bar_by_name(slide, "Chart 8", {"Not available": 100.0})


SLEEVE_SLIDES = {
    # slide_number: (parsed.json sleeve key, description text is left as-is from template)
    13: "Fixed Income",
    15: "Equities",
    17: "Private Assets",
    18: "Commodities",
    19: "Structured Products",
}


def fill_sleeve_slides(prs, parsed):
    for slide_num, sleeve_key in SLEEVE_SLIDES.items():
        slide = get_slide(prs, slide_num)
        sleeve = parsed["sleeves"].get(sleeve_key)
        if sleeve is None:
            set_plain_text(slide, "TextBox 5", "No holdings in this sleeve for this portfolio.")
            continue
        rebuild_holdings_table(slide, "Table 6", sleeve["top_holdings"], parsed["base_currency"])
        vehicle_chart = find_shape(slide, "Chart 9")
        if vehicle_chart is not None and vehicle_chart.has_chart:
            update_donut_by_name(slide, "Chart 9", sleeve["vehicle_breakdown_pct"])
        cov = sleeve.get("duration_coverage")
        caveat_box = find_shape(slide, "TextBox 11")
        if caveat_box is not None:
            if sleeve_key == "Fixed Income" and cov:
                set_plain_text(slide, "TextBox 11",
                                f"⚠ Duration coverage: {cov} fixed-income line(s) report a modified duration; "
                                f"portfolio duration reflects only those.")
            else:
                set_plain_text(slide, "TextBox 11", "")


def fill_proposed_bond_selection(prs, parsed):
    slide = get_slide(prs, 14)
    pb = parsed.get("proposed_bond_selection")
    if not pb:
        return
    set_plain_text(slide, "Text 1",
                    f"Proposed bond selection: {pb['num_issues']} issues, "
                    f"{pb['currency']} {pb['total_amount']:,.0f}, equally weighted at "
                    f"{pb['currency']} {pb['amount_each']:,.0f} each.")
    update_donut_by_name(slide, "Chart 0", pb["by_geography_pct"], label_style="one_decimal")
    update_donut_by_name(slide, "Chart 1", pb["by_currency_pct"], label_style="one_decimal")
    update_donut_by_name(slide, "Chart 2", pb["by_maturity_pct"], label_style="one_decimal")
    update_donut_by_name(slide, "Chart 3", pb["by_rating_pct"], label_style="one_decimal")
    set_plain_text(slide, "Text 16", "By currency")
    set_donut_legend(slide, ["Text 5", "Text 7", "Text 9", "Text 11", "Text 13", "Text 15"], pb["by_geography_pct"])
    set_donut_legend(slide, ["Text 18", "Text 20", "Text 22"], pb["by_currency_pct"])
    set_donut_legend(slide, ["Text 25", "Text 27", "Text 29"], pb["by_maturity_pct"])
    set_donut_legend(slide, ["Text 32", "Text 34", "Text 36", "Text 38", "Text 40"], pb["by_rating_pct"])


def fill_equity_breakdown(prs, parsed):
    slide = get_slide(prs, 16)
    eb = parsed["equity_breakdown"]
    update_donut_by_name(slide, "Chart 0", eb["by_geography_pct"], label_style="one_decimal")
    set_donut_legend(slide, ["Text 5", "Text 7", "Text 9", "Text 11", "Text 13"], eb["by_geography_pct"])
    if eb["by_sector_pct"]:
        update_donut_by_name(slide, "Chart 1", eb["by_sector_pct"], label_style="one_decimal")
        set_donut_legend(slide, ["Text 16", "Text 18", "Text 20", "Text 22", "Text 24"], eb["by_sector_pct"])
    else:
        set_plain_text(slide, "Text 14", "By sector (not available for this custodian file)")
        update_donut_by_name(slide, "Chart 1", {"Not available": 100.0})
        set_donut_legend(slide, ["Text 16", "Text 18", "Text 20", "Text 22", "Text 24"], {})
    if eb["by_market_cap_pct"]:
        update_donut_by_name(slide, "Chart 2", eb["by_market_cap_pct"], label_style="one_decimal")
        set_donut_legend(slide, ["Text 27", "Text 29", "Text 31"], eb["by_market_cap_pct"])
    else:
        set_plain_text(slide, "Text 25", "By market cap (not available for this custodian file)")
        update_donut_by_name(slide, "Chart 2", {"Not available": 100.0})
        set_donut_legend(slide, ["Text 27", "Text 29", "Text 31"], {})
    sleeve = parsed["sleeves"].get("Equities", {})
    ccy = parsed["base_currency"]
    n_lines = sleeve.get("num_lines", 0)
    val_m = (sleeve.get("total_weight_pct", 0) / 100) * parsed["total_value_eur"] / 1_000_000
    set_plain_text(slide, "Text 1", f"Equity sleeve: {n_lines} lines, {ccy} {val_m:.1f}m, "
                                     f"{sleeve.get('total_weight_pct', 0):.1f}% of the portfolio.")


def fill_liquidity(prs, parsed):
    slide = get_slide(prs, 20)
    update_donut_by_name(slide, "Chart 5", parsed["liquidity_profile_pct"])
    illiquid = parsed["liquidity_profile_pct"].get("Illiquid (lock-up)", 0)
    set_plain_text(slide, "TextBox 8",
                    f"About {illiquid:.1f}% of the portfolio sits in private, hedge-fund and lock-up vehicles "
                    f"with redemption gates and notice periods.\n\n"
                    f"Liquidity events should be planned around these constraints; the daily-liquid sleeve "
                    f"covers near-term needs.")


def fill_concentration(prs, parsed):
    slide = get_slide(prs, 21)
    conc = parsed["concentration"]
    fill_concentration_table(slide, "Table 3", conc["top_holdings"])
    set_plain_text(slide, "Rectangle 4", f"{conc['top_5_pct']:.1f}%\nTop 5")
    set_plain_text(slide, "Rectangle 5", f"{conc['top_10_pct']:.1f}%\nTop 10")
    set_plain_text(slide, "Rectangle 6", f"{conc['top_20_pct']:.1f}%\nTop 20")
    set_plain_text(slide, "Rectangle 7", f"{conc['positions_above_5pct']}\nPositions > 5%")


def fill_income(prs, parsed):
    slide = get_slide(prs, 22)
    income = parsed["income"]
    ry = income["running_yield_pct"]
    set_plain_text(slide, "Rectangle 5", f"{ry:.1f}%\nRunning yield" if ry is not None else "n/a\nRunning yield")
    fd = income["fi_duration_years"]
    set_plain_text(slide, "Rectangle 6", f"{fd:.1f}y\nFI duration" if fd is not None else "n/a\nFI duration")
    ri = income["rate_impact_100bp_eur"]
    ccy = parsed["base_currency"]
    set_plain_text(slide, "Rectangle 7", f"{ccy} {ri:,.0f}\nImpact +100bp" if ri is not None else "n/a\nImpact +100bp")
    fill_income_tables(slide, "Table 4", "Table 9", income)


# ------------------------------------------------------------- market slides

def fill_market_slides(prs, market):
    if not market:
        return
    s3 = get_slide(prs, 3)
    h = market["headline"]
    set_plain_text(s3, "Text Placeholder 2", h["intro_sentence"])
    rects = [sh for sh in s3.shapes if sh.name == "Rectangle"]
    for sh, stat in zip(rects, h["stats"]):
        set_plain_text_shape(sh, f"{stat['value']}\n{stat['label']}")
    textboxes = [sh for sh in s3.shapes if sh.name == "TextBox"]
    # textboxes[0]='Cross-asset scoreboard' label, [1]=index col, [2]=week col, [3]=ytd col, [4]='Three observations', [5..7]=observation Rectangles handled below
    if len(textboxes) >= 4:
        rows = market["scoreboard"]["rows"]
        set_plain_text_shape(textboxes[1], "Index\n" + "\n".join(r["index"] for r in rows))
        set_plain_text_shape(textboxes[2], "Week\n" + "\n".join(r["week"] for r in rows))
        set_plain_text_shape(textboxes[3], "YTD\n" + "\n".join(r["ytd"] for r in rows))
    obs_rects = [sh for sh in s3.shapes if sh.name == "Rectangle"][4:]  # observation cards follow the 4 stat cards
    for sh, obs in zip(obs_rects, market["three_observations"]):
        set_plain_text_shape(sh, f"{obs['headline']}\n{obs['body']}")

    s4 = get_slide(prs, 4)
    set_plain_text(s4, "Text Placeholder 2", market["four_drivers"]["intro_sentence"])
    drv_rects = [sh for sh in s4.shapes if sh.name == "Rectangle"]
    for sh, drv in zip(drv_rects, market["four_drivers"]["drivers"]):
        stat_line = f"\n{drv['stat']}" if drv.get("stat") else ""
        set_plain_text_shape(sh, f"{drv['number']} - {drv['title']}\n{drv['body']}{stat_line}")

    s5 = get_slide(prs, 5)
    set_plain_text(s5, "Text Placeholder 2", market["scenario"]["intro_sentence"])

    s6 = get_slide(prs, 6)
    set_plain_text(s6, "Text Placeholder 2", market["house_view"]["intro_sentence"])
    rect_names = ["Rectangle 4", "Rectangle 8", "Rectangle 12"]
    box_names = ["TextBox 6", "TextBox 10", "TextBox 14"]
    for rn, bn, view in zip(rect_names, box_names, market["house_view"]["views"]):
        set_plain_text(s6, rn, view["asset_class"])
        set_plain_text(s6, bn, f"{view['stance']}: {view['body']}")


def set_plain_text_shape(shape, text):
    if shape is None or not shape.has_text_frame:
        return
    lines = str(text).split("\n")
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    for run in list(p0.runs)[1:]:
        run._r.getparent().remove(run._r)
    if not p0.runs:
        p0.add_run()
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    import copy as _copy
    anchor = tf.paragraphs[0]._p
    for _ in lines[1:]:
        clone = _copy.deepcopy(anchor)
        anchor.addnext(clone)
        anchor = clone
    for para, line in zip(tf.paragraphs, lines):
        if not para.runs:
            para.add_run()
        for run in list(para.runs)[1:]:
            run._r.getparent().remove(run._r)
        para.runs[0].text = line


# --------------------------------------------------------------------- main

def build(excel_path, profile, client_name, output_path, market_update_path=None,
          valuation_date=None, bucket_overrides=None, market_cap_overrides=None,
          sector_overrides=None, keep_parsed_json=None):
    parse_cmd = [sys.executable, str(SCRIPT_DIR / "parse_portfolio.py"), excel_path,
                 "--profile", profile, "--client-name", client_name,
                 "-o", keep_parsed_json or str(SCRIPT_DIR.parent / "work" / "_parsed_tmp.json")]
    if valuation_date:
        parse_cmd += ["--valuation-date", valuation_date]
    if bucket_overrides:
        parse_cmd += ["--bucket-overrides", bucket_overrides]
    if market_cap_overrides:
        parse_cmd += ["--market-cap-overrides", market_cap_overrides]
    if sector_overrides:
        parse_cmd += ["--sector-overrides", sector_overrides]
    subprocess.run(parse_cmd, check=True)

    parsed_path = keep_parsed_json or str(SCRIPT_DIR.parent / "work" / "_parsed_tmp.json")
    parsed = json.loads(Path(parsed_path).read_text())

    market = None
    if market_update_path:
        market = json.loads(Path(market_update_path).read_text())

    prs = Presentation(str(TEMPLATE_PATH))

    # IMPORTANT: line_items_mod.build() inserts/deletes slides at position
    # 9-10, shifting every fixed slide number after it. Every other fill_*
    # function below addresses slides by their ORIGINAL template position
    # (get_slide(prs, N)), so it must run first, while slide 11 is still
    # slide 11 etc. line_items_mod.build() runs last, once every other
    # slide's content is already in place.
    fill_cover(prs, parsed)
    fill_market_slides(prs, market)
    fill_profile_dial(prs, parsed)
    fill_portfolio_overview(prs, parsed)
    fill_geo_sector(prs, parsed)
    fill_sleeve_slides(prs, parsed)
    fill_proposed_bond_selection(prs, parsed)
    fill_equity_breakdown(prs, parsed)
    fill_liquidity(prs, parsed)
    fill_concentration(prs, parsed)
    fill_income(prs, parsed)
    line_items_mod.build(prs, parsed)  # slides 9-10(+): rebuilt as real tables, runs last

    prs.save(output_path)
    print(f"Wrote {output_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--excel", required=True)
    ap.add_argument("--profile", required=True,
                     choices=["Fixed Income", "Conservative", "Moderate", "Balanced", "Growth", "Equity"])
    ap.add_argument("--client-name", default="[Client Name]")
    ap.add_argument("--market-update", default=None, help="market_update.json produced from the market PDF")
    ap.add_argument("--valuation-date", default=None)
    ap.add_argument("--bucket-overrides", default=None)
    ap.add_argument("--market-cap-overrides", default=None)
    ap.add_argument("--sector-overrides", default=None)
    ap.add_argument("--keep-parsed-json", default=None, help="also write parsed.json to this path")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()

    build(args.excel, args.profile, args.client_name, args.output,
          market_update_path=args.market_update, valuation_date=args.valuation_date,
          bucket_overrides=args.bucket_overrides, market_cap_overrides=args.market_cap_overrides,
          sector_overrides=args.sector_overrides, keep_parsed_json=args.keep_parsed_json)


if __name__ == "__main__":
    main()
