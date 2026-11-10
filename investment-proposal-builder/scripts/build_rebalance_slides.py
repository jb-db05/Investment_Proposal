#!/usr/bin/env python3
"""
Insert the "recommended changes" slides (built from rebalance.json, see
build_rebalance.py) right after the risk-profile dial and before the
proposed holdings list:

  1. Why we propose changes   - before -> after KPI boxes + objectives
  2. Current vs proposed       - native clustered-bar charts by region,
                                 sector and AI camp (% of portfolio)
  3+. Proposed sales           - grouped tables with value, P/L, rationale
                                 (paginated, a theme group never splits)
  N.  Proposed purchases       - new positions, top-ups, cash buffer

Called from build_proposal.py (--rebalance) AFTER every other fill and
after the holdings-list pagination, because inserting slides shifts every
fixed template position the other fill_* functions rely on.
"""
from __future__ import annotations

import copy

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

import build_line_items_tables as li

NAVY = RGBColor(0x20, 0x29, 0x45)
SLATE = RGBColor(0x4B, 0x5F, 0x80)
TIGER = RGBColor(0xFF, 0x6C, 0x0E)
GOLD = RGBColor(0xFF, 0xC5, 0x45)
MINT = RGBColor(0x3B, 0xAF, 0x90)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF3, 0xF4, 0xF7)
FONT = "Nunito Sans"
FONT_BOLD = "Nunito Sans ExtraBold"

LEFT = Emu(665163)
WIDTH = Emu(9398001)
TOP = Emu(2300000)
ROW_H = Emu(250000)
MAX_TABLE_H = Emu(4600000)


def _new_slide(prs, after_index, title, subtitle):
    """Insert a slide with the blank line-items layout (Title + subtitle
    placeholders only) at position after_index + 1."""
    _, anchor = li.find_slide(prs, lambda s: s.shapes.title is not None
                              and s.shapes.title.text.startswith("Proposed portfolio"))
    if anchor is None:
        raise RuntimeError("holdings-list slide not found (build the line-items slides first)")
    layout = anchor.slide_layout
    slide = prs.slides.add_slide(layout)
    lst = prs.slides._sldIdLst
    el = list(lst)[-1]
    lst.remove(el)
    lst.insert(after_index + 1, el)
    # drop any placeholder the layout adds besides title/subtitle
    for sh in list(slide.shapes):
        if sh.is_placeholder and sh.placeholder_format.idx not in (0, 1) and sh != slide.shapes.title:
            sh._element.getparent().remove(sh._element)
    slide.shapes.title.name = "Title 1"
    for sh in slide.shapes:
        if sh.is_placeholder and sh != slide.shapes.title:
            sh.name = "Text Placeholder 2"
    li.set_placeholder_text(slide, "Title 1", title)
    li.set_subtitle(slide, subtitle)
    return slide


def _text(slide, left, top, width, height, paragraphs, *, anchor=MSO_ANCHOR.TOP):
    """paragraphs: list of (text, size, bold, color, font) tuples."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(45720)
    for i, (text, size, bold, color, font) in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
        p.space_after = Pt(4)
    return tb


def _kpi(slide, left, top, width, height, value, caption):
    box = slide.shapes.add_shape(1, left, top, width, height)  # rectangle
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = RGBColor(0xD9, 0xDC, 0xE3)
    box.line.width = Pt(0.75)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (text, size, bold, color, font) in enumerate([
        (value, 20, True, TIGER, FONT_BOLD), (caption, 11, False, NAVY, FONT)]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font


def _pct0(x):
    return f"{x:.0f}%"


def _money(x):
    return f"{x:,.0f}"


def slide_diagnosis(prs, idx, rb):
    b, a = rb["equities_before"], rb["equities_after"]
    ccy = rb["base_currency"]
    china_b = b["by_region_pct"].get("China", 0)
    china_a = a["by_region_pct"].get("China", 0)
    slide = _new_slide(
        prs, idx, "Why we propose changes to the equity sleeve",
        f"The equity sleeve ({b['equity_weight_pct']:.0f}% of the portfolio) holds {b['num_lines']} lines, "
        f"{b['num_single_stocks']} of them single stocks, and {b['num_lines_below_threshold']} positions "
        f"below {b['threshold_pct']:.0f}% each. We propose fewer, larger positions and a broader spread "
        f"of regions and sectors.")
    kpis = [
        (f"{b['num_lines']} → {a['num_lines']}", "Equity lines"),
        (f"{b['num_single_stocks']} → {a['num_single_stocks']}", "Single stocks"),
        (f"{b['num_lines_below_threshold']} → {a['num_lines_below_threshold']}",
         f"Lines below {b['threshold_pct']:.0f}%"),
        (f"{_pct0(b['ai_exposed_pct_of_equities'])} → {_pct0(a['ai_exposed_pct_of_equities'])}",
         "AI-exposed share of equities"),
        (f"{china_b:.0f}% → {china_a:.0f}%", "China (% of portfolio)"),
    ]
    w, gap = Emu(1792800), Emu(108000)
    for i, (v, c) in enumerate(kpis):
        _kpi(slide, Emu(LEFT + i * (w + gap)), Emu(2412000), w, Emu(1008000), v, c)
    paras = [("What the proposal does", 13, True, NAVY, FONT_BOLD)]
    paras += [(f"•  {o}", 11, False, NAVY, FONT) for o in rb["objectives"]]
    _text(slide, LEFT, Emu(3640000), WIDTH, Emu(3000000), paras)
    _text(slide, LEFT, Emu(6900000), WIDTH, Emu(400000), [(
        f"Portfolio total unchanged at {ccy} {rb['total_value']:,.0f}; transaction costs and taxes not included. "
        f"AI-exposed / AI-neutral / AI-insulated follows the house classification "
        f"({rb.get('house_view_source') or 'house view'}).", 9, False, SLATE, FONT)])
    return slide


def _bar_chart(slide, left, top, width, height, title, before: dict, after: dict, order=None):
    cats = order or sorted(set(before) | set(after), key=lambda k: -(after.get(k, 0) + before.get(k, 0)))
    data = CategoryChartData()
    data.categories = cats
    data.add_series("Current", [round(before.get(c, 0), 1) for c in cats])
    data.add_series("Proposed", [round(after.get(c, 0), 1) for c in cats])
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, data)
    ch = gf.chart
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    tp = ch.chart_title.text_frame.paragraphs[0]
    tp.runs[0].font.size = Pt(12)
    tp.runs[0].font.bold = True
    tp.runs[0].font.color.rgb = NAVY
    tp.runs[0].font.name = FONT
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(9)
    ch.legend.font.color.rgb = NAVY
    ch.font.name = FONT
    ch.font.size = Pt(9)
    ch.font.color.rgb = NAVY
    cat_ax = ch.category_axis
    cat_ax.reverse_order = True
    cat_ax.tick_labels.font.size = Pt(9)
    cat_ax.format.line.color.rgb = RGBColor(0xD9, 0xDC, 0xE3)
    cat_ax.has_major_gridlines = False
    val_ax = ch.value_axis
    val_ax.visible = False
    val_ax.has_major_gridlines = False
    plot = ch.plots[0]
    plot.gap_width = 60
    plot.overlap = -10
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0.0"%"'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(8)
    dl.font.color.rgb = NAVY
    for s, color in zip(plot.series, (SLATE, TIGER)):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = color
    return gf


def slide_charts(prs, idx, rb):
    b, a = rb["equities_before"], rb["equities_after"]
    slide = _new_slide(
        prs, idx, "Current vs proposed equity sleeve",
        f"Equities as % of the total portfolio ({b['equity_weight_pct']:.1f}% today, "
        f"{a['equity_weight_pct']:.1f}% proposed). Diversified funds are shown as one bucket: "
        f"their own sector mix is spread across the market.")
    half = Emu(int(WIDTH / 2) - 60000)
    _bar_chart(slide, LEFT, TOP, half, Emu(2650000), "By region", b["by_region_pct"], a["by_region_pct"])
    _bar_chart(slide, LEFT, Emu(TOP + 2700000), half, Emu(1900000), "By AI camp", b["by_ai_pct"], a["by_ai_pct"],
               order=[k for k in ("AI-exposed", "AI-neutral", "AI-insulated") if k in b["by_ai_pct"] or k in a["by_ai_pct"]])
    _bar_chart(slide, Emu(LEFT + half + 120000), TOP, half, Emu(4600000), "By sector",
               b["by_sector_pct"], a["by_sector_pct"])
    return slide


def _cell(cell, text, *, bold=False, fill=None, color=NAVY, size=9, align=PP_ALIGN.LEFT, font=FONT):
    li._set_cell(cell, text, bold=bold, fill=fill, font_color=color, size=size, align=align, font_name=font)


def _table(slide, headers, widths, blocks, total_row=None):
    """blocks: list of (group_title, rows) where rows are lists of cell
    strings. Group title rows span the full width in light grey."""
    n_rows = 1 + sum(1 + len(r) for _, r in blocks) + (1 if total_row else 0)
    gf = slide.shapes.add_table(n_rows, len(headers), LEFT, TOP, WIDTH, Emu(ROW_H * n_rows))
    t = gf.table
    t.first_row = False
    t.horz_banding = False
    for i, w in enumerate(widths):
        t.columns[i].width = Emu(w)
    for r in range(n_rows):
        t.rows[r].height = ROW_H
    numeric = {i for i, h in enumerate(headers) if h.startswith(("Weight", "Value", "Trade", "Unrealised", "Before", "After"))}
    for c, h in enumerate(headers):
        _cell(t.cell(0, c), h, bold=True, fill=NAVY, color=WHITE, size=9, font=FONT_BOLD,
              align=PP_ALIGN.RIGHT if c in numeric else PP_ALIGN.LEFT)
    r = 1
    for title, rows in blocks:
        t.cell(r, 0).merge(t.cell(r, len(headers) - 1))
        _cell(t.cell(r, 0), title, bold=True, fill=GOLD, size=9.5, font=FONT_BOLD)
        r += 1
        for row in rows:
            for c, v in enumerate(row):
                _cell(t.cell(r, c), v, size=8.5 if c == len(headers) - 1 else 9,
                      align=PP_ALIGN.RIGHT if c in numeric else PP_ALIGN.LEFT)
            r += 1
    if total_row:
        for c, v in enumerate(total_row):
            _cell(t.cell(r, c), v, bold=True, fill=LIGHT, size=9, font=FONT_BOLD,
                  align=PP_ALIGN.RIGHT if c in numeric else PP_ALIGN.LEFT)
    return gf


SALES_HEADERS = ["Instrument", "Weight", "Value", "Unrealised P/L", "Rationale"]
SALES_WIDTHS = [2500000, 700000, 950000, 1150000, 4098001]


def _sale_rows(items):
    rows = []
    for s in items:
        if "value_after" in s:  # trim
            rows.append([f"{s['name']} (trim to {s['weight_after_pct']:.1f}%)",
                         f"-{s['weight_before_pct'] - s['weight_after_pct']:.1f}%",
                         _money(-s["change"]),
                         _money(s["unrealized_pl_realized"]) if s.get("unrealized_pl_realized") is not None else "",
                         s["reason"]])
        else:
            rows.append([s["name"], f"-{s['weight_pct']:.1f}%", _money(s["value"]),
                         _money(s["unrealized_pl"]) if s.get("unrealized_pl") is not None else "", s["reason"]])
    return rows


def slides_sales(prs, idx, rb, rows_budget=15):
    items = rb["sells"] + rb["trims"]
    groups = rb.get("sale_groups") or []
    blocks = []
    for g in groups + [None]:
        members = [s for s in items if (s.get("group") or None) == g]
        if members:
            blocks.append((g or "Other sales", members))
    # paginate: a group never splits; one group-title row + rows per group
    pages, cur, used = [], [], 0
    for title, members in blocks:
        cost = 1 + len(members)
        if cur and used + cost > rows_budget:
            pages.append(cur)
            cur, used = [], 0
        cur.append((title, members))
        used += cost
    if cur:
        pages.append(cur)
    ccy = rb["base_currency"]
    total_sales = rb["gross_sales"]
    pl = rb["unrealized_pl_on_sales"]
    slides = []
    for n, page in enumerate(pages, 1):
        slide = _new_slide(
            prs, idx + len(slides), "Proposed sales" + (f" ({n} of {len(pages)})" if len(pages) > 1 else ""),
            f"{len(items)} sales and trims raise {ccy} {total_sales:,.0f}. Unrealised P/L crystallised: "
            f"{ccy} {pl:,.0f} (cost basis as reported by the custodian; check the tax position before trading). "
            f"Values in {ccy}.")
        _table(slide, SALES_HEADERS, SALES_WIDTHS,
               [(t, _sale_rows(m)) for t, m in page],
               total_row=(["Total sales", f"-{sum(s.get('weight_pct', 0) for s in items) + sum(s['weight_before_pct'] - s['weight_after_pct'] for s in rb['trims']):.1f}%",
                           _money(total_sales), _money(pl), ""] if n == len(pages) else None))
        slides.append(slide)
    return slides


BUY_HEADERS = ["Instrument", "Before", "After", "Trade", "Rationale"]
BUY_WIDTHS = [2900000, 650000, 650000, 950000, 4248001]


def slide_purchases(prs, idx, rb):
    ccy = rb["base_currency"]
    new_rows = [[p["name"], "-", f"{p['weight_pct']:.1f}%", _money(p["value"]), p["reason"]]
                for p in rb["new_positions"]]
    top_rows = [[p["name"], f"{p['weight_before_pct']:.1f}%", f"{p['weight_after_pct']:.1f}%",
                 _money(p["change"]), p["reason"]] for p in rb["topups"]]
    cash_change = (rb["cash_after_pct"] - rb["cash_before_pct"]) / 100 * rb["total_value"]
    cash_rows = [[f"Current account in {ccy}", f"{rb['cash_before_pct']:.1f}%", f"{rb['cash_after_pct']:.1f}%",
                  _money(cash_change),
                  "House view: overweight cash. Buffer to redeploy into the same targets in stages, "
                  "e.g. after Nvidia results (25 Nov) and 2027 capex guidance (late Jan)."]]
    slide = _new_slide(
        prs, idx, "Proposed purchases",
        f"{len(rb['new_positions'])} new funds and {len(rb['topups'])} top-ups, {ccy} "
        f"{rb['gross_purchases']:,.0f} in total; the balance stays in cash. Weights are % of the total "
        f"portfolio, trade values in {ccy}; unit counts depend on the execution price.")
    _table(slide, BUY_HEADERS, BUY_WIDTHS,
           [("New positions", new_rows), ("Top-ups to funds already held", top_rows), ("Cash buffer", cash_rows)])
    return slide


def build(prs, rb):
    """Insert the slides after the risk-profile dial (template slide 8)."""
    idx, _ = li.find_slide(prs, lambda s: any(
        sh.has_text_frame and "Proposed profile for this portfolio" in sh.text_frame.text for sh in s.shapes))
    if idx is None:
        raise RuntimeError("risk-profile dial slide not found")
    s = slide_diagnosis(prs, idx, rb)
    idx += 1
    slide_charts(prs, idx, rb)
    idx += 1
    sales = slides_sales(prs, idx, rb)
    idx += len(sales)
    slide_purchases(prs, idx, rb)
    return prs
