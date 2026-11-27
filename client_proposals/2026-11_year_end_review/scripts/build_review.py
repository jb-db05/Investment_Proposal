#!/usr/bin/env python3
"""Build the client portfolio review & diversification proposal on the Syz template.

Usage: build_deck.py analysis.json template.pptx out.pptx
"""
import json, sys, copy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION

A = json.load(open(sys.argv[1]))
prs = Presentation(sys.argv[2])
OUT = sys.argv[3]

NAVY = RGBColor(0x20, 0x29, 0x45); GOLD = RGBColor(0xFF, 0xC5, 0x45)
SLATE = RGBColor(0x4B, 0x5F, 0x80); MINT = RGBColor(0x3B, 0xAF, 0x90)
ORANGE = RGBColor(0xFF, 0xA4, 0x00); TIGER = RGBColor(0xFF, 0x6C, 0x0E)
MAUVE = RGBColor(0xAC, 0x5D, 0x85); SKY = RGBColor(0x79, 0xD6, 0xFF)
PEACH = RGBColor(0xE6, 0xA4, 0xAD); GREY = RGBColor(0x7A, 0x80, 0x90)
LIGHT = RGBColor(0xF4, 0xF5, 0xF8); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x39, 0x2B)
BODY = "Nunito Sans"; BOLD = "Nunito Sans ExtraBold"

AC_COLORS = {"Equities": SKY, "Commodities": MINT, "Private assets": MAUVE, "Fixed income": SLATE,
             "Structured products": PEACH, "Cash & money market": GOLD, "Digital assets": ORANGE,
             "Hedge funds": NAVY}

L = 666000; R = 10026000; W = R - L          # content box
TOTAL = A["total_eur"]

def eur_m(w):  # weight % -> EUR m string
    return f"EUR {w/100*TOTAL/1e6:,.2f}m"

def pct(v, d=1, sign=False):
    if v is None: return "n/a"
    s = f"{v:+.{d}f}%" if sign else f"{v:.{d}f}%"
    if not sign and 0 < v < 0.05: s = "<0.1%"
    return s

# ------------------------------------------------------------------ helpers
def tb(slide, x, y, w, h, paras, *, fill=None, line=None, anchor=MSO_ANCHOR.TOP, margin=60000, shape=MSO_SHAPE.RECTANGLE):
    """paras: list of (text, size, bold, color[, align]) or str."""
    if fill is None and line is None:
        sh = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    else:
        sh = slide.shapes.add_shape(shape, Emu(x), Emu(y), Emu(w), Emu(h))
        sh.shadow.inherit = False
        if fill is None: sh.fill.background()
        else:
            sh.fill.solid(); sh.fill.fore_color.rgb = fill
        if line is None: sh.line.fill.background()
        else:
            sh.line.color.rgb = line; sh.line.width = Pt(1)
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(margin); tf.margin_top = tf.margin_bottom = Emu(margin // 2)
    first = True
    for p in paras:
        if isinstance(p, str): p = (p, 11, False, NAVY)
        text, size, bold, color = p[:4]
        align = p[4] if len(p) > 4 else PP_ALIGN.LEFT
        para = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        para.alignment = align
        para.space_after = Pt(3)
        r = para.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        r.font.name = BOLD if bold else BODY
    return sh

def new_slide(title, subtitle=None):
    s = prs.slides.add_slide(prs.slide_layouts[21])  # "Title" layout
    s.shapes.title.text = title
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 14:  # never edit subtitle placeholder in place (repo rule)
            ph._element.getparent().remove(ph._element)
    if subtitle:
        tb(s, 665163, 1728947, 9398001, 520000, [(subtitle, 13, False, NAVY)], margin=0)
    return s

def divider(title, sub=None):
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = title
    if sub:
        tb(s, 665163, 4600000, 9398001, 600000, [(sub, 16, False, WHITE)], margin=0)
    return s

def header(slide, x, y, w, text):
    tb(slide, x, y, w, 300000, [(text, 13, True, NAVY)], margin=0)

def source(slide, text):
    tb(slide, L, 7080000, W, 260000, [(text, 8, False, GREY)], margin=0)

def kpi(slide, x, y, w, h, value, caption, color=TIGER):
    tb(slide, x, y, w, h, [(value, 20, True, color, PP_ALIGN.CENTER), (caption, 10, False, NAVY, PP_ALIGN.CENTER)],
       fill=WHITE, line=GOLD, anchor=MSO_ANCHOR.MIDDLE)

def card(slide, x, y, w, h, title, lines, accent=ORANGE, title_size=12, body_size=10):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(55000), Emu(h))
    bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background(); bar.shadow.inherit = False
    paras = [(title, title_size, True, NAVY)] + [(("• " + l) if not l.startswith("§") else l[1:], body_size, False, NAVY) for l in lines]
    return tb(slide, x + 55000, y, w - 55000, h, paras, fill=LIGHT, margin=90000)

def table(slide, x, y, w, col_w, rows, *, header_fill=GOLD, size=9, row_h=255000, colors=None, bold_rows=(), aligns=None, fills=None):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Emu(x), Emu(y), Emu(w), Emu(row_h * nr))
    t = gt.table
    tblPr = t._tbl.tblPr
    # drop the default table style banding look
    tblPr.set("bandRow", "0"); tblPr.set("firstRow", "0")
    tot = sum(col_w)
    for i, cw in enumerate(col_w): t.columns[i].width = Emu(int(w * cw / tot))
    for ri, row in enumerate(rows):
        t.rows[ri].height = Emu(row_h)
        for ci, val in enumerate(row):
            c = t.cell(ri, ci); c.margin_left = c.margin_right = Emu(50000)
            c.margin_top = c.margin_bottom = Emu(15000); c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            if ri == 0: c.fill.fore_color.rgb = header_fill
            elif fills and fills.get(ri): c.fill.fore_color.rgb = fills[ri]
            else: c.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
            tf = c.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = ""
            r = p.add_run(); r.text = str(val)
            b = ri == 0 or ri in bold_rows
            r.font.size = Pt(size); r.font.bold = b; r.font.name = BOLD if b else BODY
            col = NAVY
            if colors and ri > 0: col = colors(ri, ci, val) or NAVY
            r.font.color.rgb = col
            if aligns: p.alignment = aligns[ci]
    return gt

def style_chart_text(chart, size=9):
    chart.has_title = False
    chart.font.size = Pt(size); chart.font.name = BODY; chart.font.color.rgb = NAVY

def donut(slide, x, y, w, h, data, colors, hole_label=None):
    cd = CategoryChartData(); cd.categories = list(data.keys()); cd.add_series("w", list(data.values()))
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart; ch.has_legend = False; style_chart_text(ch)
    plot = ch.plots[0]
    from lxml import etree
    hs = plot._element.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}holeSize')
    if hs is not None: hs.set('val', '62')
    for i, pt in enumerate(plot.series[0].points):
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = colors[i]
        pt.format.line.color.rgb = WHITE
    if hole_label:
        tb(slide, x + w // 2 - 700000, y + h // 2 - 250000, 1400000, 500000,
           [(hole_label, 11, True, NAVY, PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, margin=0)
    return ch

def legend(slide, x, y, w, data, colors, size=10, row=270000, fmt=lambda v: pct(v)):
    for i, (k, v) in enumerate(data.items()):
        sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y + i * row + 70000), Emu(130000), Emu(130000))
        sq.fill.solid(); sq.fill.fore_color.rgb = colors[i]; sq.line.fill.background(); sq.shadow.inherit = False
        tb(slide, x + 200000, y + i * row, w - 200000 - 800000, row, [(k, size, False, NAVY)], margin=0)
        tb(slide, x + w - 800000, y + i * row, 800000, row, [(fmt(v), size, True, NAVY, PP_ALIGN.RIGHT)], margin=0)

def hbar(slide, x, y, w, h, data, color=SLATE, colors=None, num_fmt='0.0"%"', size=9, gap=60, minv=None, maxv=None, series=None, series_colors=None, legend_on=False):
    cd = CategoryChartData(); cd.categories = list(data.keys()) if series is None else data
    if series is None:
        cd.add_series("v", list(data.values()))
    else:
        for name, vals in series: cd.add_series(name, vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    ch = gf.chart; style_chart_text(ch, size)
    ch.has_legend = legend_on
    if legend_on:
        ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(size)
    ca = ch.category_axis; ca.reverse_order = True; ca.has_major_gridlines = False
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ca.format.line.color.rgb = RGBColor(0xD0, 0xD3, 0xDA); ca.tick_labels.font.size = Pt(size)
    va = ch.value_axis; va.visible = False; va.has_major_gridlines = False
    if minv is not None: va.minimum_scale = minv
    if maxv is not None: va.maximum_scale = maxv
    plot = ch.plots[0]; plot.gap_width = gap
    if series is not None: plot.overlap = -10
    plot.has_data_labels = True
    dl = plot.data_labels; dl.number_format = num_fmt; dl.number_format_is_linked = False
    dl.font.size = Pt(size); dl.font.color.rgb = NAVY; dl.position = XL_LABEL_POSITION.OUTSIDE_END
    for si, s in enumerate(plot.series):
        s.invert_if_negative = False
        if series_colors:
            s.format.fill.solid(); s.format.fill.fore_color.rgb = series_colors[si]
        else:
            for i, pt in enumerate(s.points):
                pt.format.fill.solid(); pt.format.fill.fore_color.rgb = colors[i] if colors else color
    return ch


# ================================================================== YEAR-END REVIEW BODY
H = A["holdings"]
AC_ORDER = ["Equities", "Commodities", "Private assets", "Fixed income", "Structured products",
            "Cash & money market", "Digital assets", "Hedge funds"]
ac = {k: A["asset_class"].get(k, 0) for k in AC_ORDER}
REG_LABEL = {"Eurozone": "Europe ex-UK & CH", "Commodities (no region)": "Gold & commodities (no region)",
             "Global": "Global / multi-region", "EM": "Emerging markets"}
region = {REG_LABEL.get(k, k): v for k, v in A["region"].items() if v >= 0.05}
eq_region = {REG_LABEL.get(k, k): v for k, v in A["eq_region"].items()}
eq_sector = A["eq_sector"]
ccy = {k: v for k, v in A["ccy"].items() if v >= 2}
ccy["Other"] = sum(v for k, v in A["ccy"].items() if v < 2)
gold_phys = A["commodity_sub"].get("Gold (physical)", 0)
gold_miners = A["commodity_sub"].get("Gold miners", 0)
other_metals = A["commodity_sub"].get("Other metals & commodities", 0)
theme = A["theme"]
EUROPE = ("Eurozone", "Switzerland", "UK")

# ---- estimated YTD (buy-and-hold on current holdings; start weight = w / (1+r))
def perf_ok(h):
    return h["perf_eur"] is not None and not (h["is_account"] and h["ccy"] not in ("XAG", "XPT"))
start = {}; contrib = {}
for h in H:
    r = (h["perf_eur"] or 0) / 100 if perf_ok(h) else 0.0
    sw = h["weight"] / (1 + r)
    start[h["desc"]] = sw
    contrib.setdefault(h["asset_class"], 0.0)
    contrib[h["asset_class"]] += sw * r
SW = sum(start.values())
contrib = {k: v / SW * 100 for k, v in contrib.items()}      # % points of portfolio
EST_YTD = sum(contrib.values())

def sleeve_ret(f):
    sw = sum(start[h["desc"]] for h in H if f(h) and perf_ok(h))
    ew = sum(h["weight"] for h in H if f(h) and perf_ok(h))
    return (ew / sw - 1) * 100 if sw else 0
eq_eu = sleeve_ret(lambda h: h["asset_class"] == "Equities" and h["region"] in EUROPE)
eq_rest = sleeve_ret(lambda h: h["asset_class"] == "Equities" and h["region"] not in EUROPE)
eq_all = sleeve_ret(lambda h: h["asset_class"] == "Equities")
eq_eu_w = sum(h["weight"] for h in H if h["asset_class"] == "Equities" and h["region"] in EUROPE)
ai_ret = sleeve_ret(lambda h: h["theme"] == "AI")

# ------------------------------------------------------------------ 1: cover
cover = prs.slides[0]
for sh in cover.shapes:
    if sh.name == "Title 1":
        sh.text_frame.text = "Year-End\vPortfolio Review 2026"
    if sh.name == "Text Placeholder 5":
        tf = sh.text_frame
        lines = [f"Value-oriented EUR portfolio · EUR {TOTAL/1e6:.1f}m", "", "Prepared for our client", "",
                 "Portfolio data as of 26 November 2026"]
        paras = tf.paragraphs
        for i, line in enumerate(lines):
            p = paras[i] if i < len(paras) else tf.add_paragraph()
            if p.runs:
                p.runs[0].text = line
                for r in p.runs[1:]: r.text = ""
            else:
                r = p.add_run(); r.text = line
        for p in paras[len(lines):]:
            for r in p.runs: r.text = ""
template_n = len(prs.slides)

# ------------------------------------------------------------------ 2: executive summary
s = new_slide("2026 at a glance",
              f"A modestly positive year (≈ {EST_YTD:+.1f}%), built on European value stocks and private equity – and a portfolio that held up well in the November correction.")
kw = (W - 3 * 120000) // 4
kpi(s, L, 2350000, kw, 780000, f"EUR {TOTAL/1e6:.1f}m", "Portfolio value")
kpi(s, L + kw + 120000, 2350000, kw, 780000, f"≈ {EST_YTD:+.1f}%", "Estimated YTD return (EUR)", color=MINT)
kpi(s, L + 2 * (kw + 120000), 2350000, kw, 780000, f"{eq_eu:+.1f}%", "European & Swiss equities YTD", color=MINT)
kpi(s, L + 3 * (kw + 120000), 2350000, kw, 780000, pct(gold_phys + gold_miners), "Gold (incl. miners)", color=TIGER)
cw = (W - 2 * 150000) // 3; y = 3330000; h = 3550000
card(s, L, y, cw, h, "Your portfolio", [
    f"Equities {pct(ac['Equities'])}, two-thirds in Europe & Switzerland, tilted to insurers, banks, defence and Swiss defensives",
    f"Commodities {pct(ac['Commodities'])}, mostly physical gold ({pct(gold_phys)})",
    f"Private assets {pct(ac['Private assets'])}; fixed income {pct(ac['Fixed income'])}"], accent=SKY, body_size=12.5)
card(s, L + cw + 150000, y, cw, h, "What drove 2026", [
    f"Private equity (+{contrib.get('Private assets',0):.1f} pts) and equities (+{contrib.get('Equities',0):.1f} pts) carried the year",
    f"Commodities cost {contrib.get('Commodities',0):.1f} pts: gold −12%, silver and platinum −31% to −39%",
    "Low exposure to AI and tech limited the damage of the sell-off"], accent=MINT, body_size=12.5)
card(s, L + 2 * (cw + 150000), y, cw, h, "Looking to 2027", [
    "Cycle extended, with a dip first; inflation close to peaking",
    "Valuations back to average after the de-rating; earnings still rising",
    "Value, Europe and gold remain consistent with our 2027 themes"], accent=ORANGE, body_size=12.5)
source(s, "Source: client portfolio statement (26 Nov 2026), Syz calculations; Syz Outlook 2027; Syz Investment Conclusions, Nov 2026. "
          "YTD return estimated from each holding's time-weighted EUR return – indicative, excludes flows and positions sold during the year.")

# ------------------------------------------------------------------ 3: divider
divider("Your portfolio in 2026", "Allocation, exposures and results")

# ------------------------------------------------------------------ 4: overview
s = new_slide("Allocation by asset class",
              "Commodities include silver and platinum accounts, BGF World Mining, VanEck Gold Miners, M&R Earths ETF and the DCI XAG-USD notes.")
kw = (W - 4 * 120000) // 5; ky = 2380000
kpi(s, L, ky, kw, 720000, f"EUR {TOTAL/1e6:.1f}m", "Total value")
kpi(s, L + (kw + 120000), ky, kw, 720000, str(A["n"]), "Positions")
kpi(s, L + 2 * (kw + 120000), ky, kw, 720000, pct(ac["Equities"]), "Equities")
kpi(s, L + 3 * (kw + 120000), ky, kw, 720000, pct(gold_phys), "Physical gold")
kpi(s, L + 4 * (kw + 120000), ky, kw, 720000, pct(ac["Private assets"]), "Private assets")
header(s, L, 3300000, 5000000, "Allocation by asset class")
acd = {k: v for k, v in ac.items() if v > 0}
cols = [AC_COLORS[k] for k in acd]
donut(s, L, 3620000, 2500000, 2500000, acd, cols)
legend(s, L + 2600000, 3700000, 2700000, acd, cols)
header(s, 6300000, 3300000, 3700000, "Inside the commodities sleeve")
cs = {"Physical gold": gold_phys, "Gold miners": gold_miners, "Other metals & commodities": other_metals}
donut(s, 6300000, 3620000, 1650000, 1650000, cs, [GOLD, ORANGE, SLATE])
legend(s, 6300000, 5400000, 3700000, cs, [GOLD, ORANGE, SLATE])
tb(s, 6300000, 6250000, 3700000, 700000, [(
    "Silver: half is held in a DCI maturing 27 Nov (strike USD 58), the other half was rolled on 26 Nov "
    "into a new DCI (16% p.a., strike USD 47.707, 29 Jan 2027).", 9, False, NAVY)], margin=0)
source(s, "Source: client portfolio statement, 26 Nov 2026; weights as reported by the custodian. Classification of the six listed securities as commodities per client instruction.")

# ------------------------------------------------------------------ 5: region & currency
s = new_slide("Allocation by region and currency",
              "Your equities are anchored in Europe and Switzerland – the region that held up best in this year's AI-led correction.")
header(s, L, 2350000, 4600000, "Whole portfolio by region")
hbar(s, L, 2650000, 4600000, 3000000, region, colors=[MINT if "Gold" in k else SLATE for k in region])
header(s, 5418000, 2350000, 4600000, "Equity sleeve by region")
hbar(s, 5418000, 2600000, 4600000, 1950000, eq_region,
     colors=[SKY if k in ("Europe ex-UK & CH", "Switzerland", "UK") else RGBColor(0xC5, 0xCB, 0xD8) for k in eq_region], size=8.5)
header(s, 5418000, 4600000, 4600000, "Currency exposure")
hbar(s, 5418000, 4880000, 4600000, 2150000, ccy, color=GOLD, size=8.5, gap=40)
tb(s, L, 5750000, 4600000, 1250000, [
    ("What stands out", 11, True, NAVY),
    (f"Europe, Switzerland and the UK make up {pct(A['eq_region'].get('Eurozone',0)+A['eq_region'].get('Switzerland',0)+A['eq_region'].get('UK',0))} "
     f"of your equities; US equities are only {pct(A['eq_region']['US'])}, mostly Berkshire Hathaway.", 9.5, False, NAVY),
    (f"USD is {pct(A['ccy']['USD'])} of assets, largely through the private-asset funds.", 9.5, False, NAVY)],
   fill=LIGHT, margin=90000)
source(s, "Source: client portfolio statement; region assigned at issuer / fund-mandate level by Syz. 'Global / multi-region' is mainly private-equity, hedge and thematic funds.")

# ------------------------------------------------------------------ 6: sectors & themes
s = new_slide("Sector allocation and themes",
              "A value profile: financials and industrials dominate, technology is small and energy is almost absent.")
header(s, L, 2350000, 4600000, "Equity sleeve by sector")
hbar(s, L, 2650000, 4600000, 3300000, eq_sector, colors=[TIGER if k == "Information technology" else SKY for k in eq_sector])
tb(s, L, 6000000, 4600000, 700000, [(
    "'Diversified' = broad equity funds (Swiss income, EU high dividend, EU mid-caps, US small caps). "
    "Mining and gold-miner funds are shown under commodities.", 8.5, False, GREY)], margin=0)
header(s, 5418000, 2350000, 4600000, "Thematic exposure (% of total portfolio)")
kw2 = (4600000 - 2 * 100000) // 3
kpi(s, 5418000, 2700000, kw2, 800000, pct(theme.get("Defence", 0)), "Defence", color=SLATE)
kpi(s, 5418000 + kw2 + 100000, 2700000, kw2, 800000, pct(theme.get("AI", 0)), "AI / mega-cap tech", color=TIGER)
kpi(s, 5418000 + 2 * (kw2 + 100000), 2700000, kw2, 800000, pct(theme.get("Energy", 0)), "Energy", color=MINT)
card(s, 5418000, 3650000, 4600000, 3200000, "Reading the exposures", [
    "Financials: Allianz, Zurich, Hannover Re, Berkshire, Lloyds – the backbone of the value tilt",
    "Defence: Airbus, BAE, Dassault, Leonardo, Saab and two defence funds – a 2026 winner overall",
    "AI: limited to the Meta/Amazon bonds, two tech notes, the Digital Security ETF and China Next Winners – in line with your preference",
    "Energy: only the PKN Orlen bond; given your career in energy, we do not count it as a diversifier"],
     accent=ORANGE, body_size=10)
source(s, "Source: client portfolio statement; sector and theme tags assigned by Syz at issuer / fund-mandate level (the custodian file has no sector field).")

# ------------------------------------------------------------------ 7: results
s = new_slide("2026 results: what added and what cost",
              f"Estimated YTD return of about {EST_YTD:+.1f}% in EUR: private equity and European equities more than offset the fall in gold and metals.")
header(s, L, 2350000, 5300000, "Contribution to YTD return by asset class (% points)")
cdata = {k: contrib.get(k, 0) for k in sorted(contrib, key=lambda k: -contrib[k]) if abs(contrib.get(k, 0)) >= 0.005}
hbar(s, L, 2650000, 5300000, 3700000, cdata, colors=[MINT if v >= 0 else TIGER for v in cdata.values()],
     num_fmt='+0.0;-0.0', minv=-5, maxv=6.5)
rows = [["Sleeve", "Weight", "YTD (EUR)"]]
for k in AC_ORDER:
    r = sleeve_ret(lambda h, k=k: h["asset_class"] == k)
    rows.append([k, pct(ac[k]), pct(r, 1, True)])
def rcol(ri, ci, v):
    if ci == 2: return MINT if v.startswith("+") else TIGER
table(s, 6200000, 2650000, 3826000, [2.0, 0.9, 1.0], rows, size=9.5, row_h=330000, colors=rcol,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
tb(s, 6200000, 5700000, 3826000, 1100000, [
    ("Within equities", 11, True, NAVY),
    (f"European, Swiss & UK stocks {eq_eu:+.1f}% vs {eq_rest:+.1f}% for the rest of the equity sleeve.", 10, False, NAVY)],
   fill=LIGHT, margin=90000)
source(s, "Source: client portfolio statement, Syz calculations. Estimate based on each current holding's time-weighted EUR return (YTD, or since purchase for 2026 purchases); "
          "excludes cash flows and positions sold during the year. Private-asset valuations are reported with a lag.")

# ------------------------------------------------------------------ 8: best & worst
s = new_slide("Best and worst performers year to date",
              "Private equity, Swiss defensives, insurers and European defence led; silver, platinum, rare earths and China tech lagged.")
ps = [x for x in A["perf_sorted"] if x[3] is not None]
SHORT = {"COA KONA III REF3 LP": "CoA Kona III (PE)", "ISHS DIGITAL SECURITY ETF": "iShares Digital Security",
         "SWISSCOM (NOM)": "Swisscom", "BERKSHIRE HATHAWAY B": "Berkshire Hathaway B", "ALLIANZ (NOM)": "Allianz",
         "BIL OC 211.5% NESN  13.12.2027": "Nestlé outperf. cert. 2027", "HANNOVER RUECK (NOM)": "Hannover Re",
         "BAE SYSTEMS": "BAE Systems", "UBS(CH) EQ-SWISS INC CHFP": "UBS Swiss Income Equity", "LEFT LANE CP II CLT2": "Left Lane II (PE)",
         "ZURICH INSURANCE GRP(NOM)": "Zurich Insurance", "SAAB AB B": "Saab B",
         "GOLD KG 1 999.9 CHF": "Gold bullion (kg)", "VANECK GOLD MINERS ETF": "VanEck Gold Miners",
         "OR KRUGERRAND(PCE)1OZ USD": "Krugerrand coins",
         "16,77%DCI-XAGUSD-58-UBS26  27.11.2026": "DCI XAG-USD 27.11.26", "SG CHINA NEXT WINNERS  01.06.2099": "SG China Next Winners",
         "DECALIA AEGIS DEFENSE A1": "Decalia Aegis Defense", "ISHARES BITCOIN ETF": "iShares Bitcoin ETF",
         "CURRENT ACCOUNT IN XPT": "Platinum (XPT) account", "SWC(CH)PLATINUM ETF EACHF": "SWC Platinum ETF",
         "WISD STRAT M&R EARTHS ETF": "WisdomTree M&R Earths", "CURRENT ACCOUNT IN XAG": "Silver (XAG) account"}
def wlab(w): return "<0.1%" if w < 0.05 else f"{w:.1f}%"
top = {f"{SHORT.get(d, d.title())}  ({wlab(w)})": p for d, a_, w, p, q, lo in ps[:10]}
bot = {f"{SHORT.get(d, d.title())}  ({wlab(w)})": p for d, a_, w, p, q, lo in ps[-10:][::-1]}
header(s, L, 2350000, 4600000, "Top 10 – performance in EUR (weight)")
hbar(s, L, 2650000, 4600000, 4200000, top, color=MINT, num_fmt='+0.0"%";-0.0"%"', maxv=58)
header(s, 5418000, 2350000, 4600000, "Bottom 10 – performance in EUR (weight)")
hbar(s, 5418000, 2650000, 4600000, 4200000, bot, color=TIGER, num_fmt='+0.0"%";-0.0"%"', minv=-50, maxv=0)
source(s, "Source: client portfolio statement (26 Nov 2026), time-weighted performance in EUR (YTD, or since purchase for 2026 purchases); cash and fiduciary deposits excluded. "
          "The silver account is now a small residual after the DCI roll.")

# ------------------------------------------------------------------ 9: divider
divider("The November correction", "What happened in the last days – and why your portfolio held up")

# ------------------------------------------------------------------ 10: what happened
s = new_slide("What happened: an AI-led correction, not an earnings scare",
              "The S&P 500 lost 8.6% in the week to 20 November – its worst week since April 2025 – and closed in correction after seven straight losses.")
header(s, L, 2350000, 4600000, "Since the last Investment Committee (to 20 Nov)")
since = {"NASDAQ 100": -9.4, "MSCI EM": -6.1, "S&P 500": -6.4, "MSCI AC World": -5.8, "Japan (MSCI)": -5.5,
         "S&P 500 Equal Weight": -2.6, "STOXX Europe 600": -0.2, "Swiss SPI": 1.1}
hbar(s, L, 2650000, 4600000, 2500000, since, colors=[MINT if k in ("STOXX Europe 600", "Swiss SPI") else (TIGER if v < -5 else SLATE) for k, v in since.items()],
     num_fmt='+0.0"%";-0.0"%"', size=8.5, minv=-14, maxv=4)
header(s, L, 5200000, 4600000, "Value vs growth since the last IC")
sty = {"MSCI ACWI Value": -4.2, "MSCI ACWI Growth": -7.5}
hbar(s, L, 5450000, 4600000, 1000000, sty, colors=[MINT, TIGER], num_fmt='+0.0"%";-0.0"%"', size=8.5, minv=-14, maxv=4, gap=40)
obs = [("Earnings delivered, prices did not",
        "Q3: 84% of S&P 500 companies beat, EPS +31.9%, a third quarter above 25%. The 'other 493' now grow faster than the Magnificent 7."),
       ("The culprit: AI capex and cash flow",
        "Hyperscaler 2027 capex up 91% since January; a ~USD 595bn funding gap for 2027–28 now has to be financed at higher rates. Nvidia −33% since 6 Nov."),
       ("Rates add pressure",
        "US 10-year at 5.13% (16 Nov) and a December Fed hike almost fully priced."),
       ("Valuations reset",
        "S&P 500 at 16.5–18.3x forward earnings, below its 5-year average (19.8x); Europe 13.5x.")]
oy = 2350000
for t, b in obs:
    card(s, 5418000, oy, 4600000, 1080000, t, ["§" + b], accent=ORANGE, title_size=11, body_size=9.5)
    oy += 1160000
source(s, "Source: Syz Earnings Pulse, 23 Nov 2026; Syz Research Forum, 23 Nov 2026 (FactSet, Bloomberg); Syz Investment Conclusions, Nov 2026. Total return; local currency for Stoxx 600 and SPI, USD otherwise.")

# ------------------------------------------------------------------ 11: Europe resilience
s = new_slide("Why Europe – and your portfolio – held up",
              "Europe's market is built on banks, insurers, health care, consumer goods and industrials, with far less technology than the US – exactly where the sell-off hit.")
header(s, L, 2350000, 4600000, "World sectors since the last IC (to 20 Nov)")
sec = {"Cons. discretionary": 1.2, "Cons. staples": 1.1, "Utilities": 0.7, "Health care": -0.7, "Financials": -1.8,
       "Industrials": -3.4, "Materials": -5.8, "Comm. services": -7.2, "Energy": -12.2, "Information tech.": -13.0}
hbar(s, L, 2650000, 4600000, 3000000, sec, colors=[MINT if v > -2 else (SLATE if v > -6 else TIGER) for v in sec.values()],
     num_fmt='+0.0"%";-0.0"%"', size=8.5, minv=-18, maxv=4)
tb(s, L, 5750000, 4600000, 1150000, [
    ("Europe was 'the resilient region'", 11, True, NAVY),
    ("Stoxx 600 −0.2% and SPI +1.1% since the IC, vs −6.4% for the S&P 500 and −9.4% for the Nasdaq 100. "
     "Energy fell too as oil dropped from ~$116 to ~$91 – your portfolio holds almost none.", 9.5, False, NAVY)],
   fill=LIGHT, margin=90000)
header(s, 5418000, 2350000, 4600000, "Your equities, YTD in EUR")
cmp = {"European, Swiss & UK equities": eq_eu, "Rest of equity sleeve": eq_rest, "AI-linked holdings": ai_ret}
hbar(s, 5418000, 2650000, 4600000, 1300000, cmp, colors=[MINT, SLATE, TIGER], num_fmt='+0.0"%";-0.0"%"', size=9, minv=-5, maxv=18, gap=40)
card(s, 5418000, 4050000, 4600000, 2850000, "How your portfolio benefited", [
    f"{pct(eq_eu_w)} of the portfolio sits in European, Swiss and UK stocks – insurers (Allianz, Zurich, Hannover Re), "
    "Swisscom, Swiss income and EU high-dividend funds",
    f"Information technology is only {pct(eq_sector.get('Information technology',0))} of your equities and AI-linked holdings "
    f"{pct(theme.get('AI',0))} of the portfolio",
    "The year's drag came from gold and metals, not from the equity sell-off"],
     accent=MINT, body_size=11)
source(s, "Source: Syz Research Forum, 23 Nov 2026 (FactSet); client portfolio statements (Nov 2026), Syz calculations. Sector returns in local currency. "
          "Equity returns: time-weighted EUR, YTD or since purchase.")

# ------------------------------------------------------------------ 12: house view now
s = new_slide("Our house view after the correction",
              "November Investment Committee: valuations upgraded to positive, equity exposure kept neutral, gold still overweight.")
rows = [["Pillar", "Stance", "Why"],
        ["Valuations", "Positive ↑", "S&P 500 16.5x (22x a year ago); Europe 13.5x, Japan 13.7x at or below averages"],
        ["Earnings growth", "Positive", "2026 EPS +34.8% US, +24.1% Europe; 2027 still revised up"],
        ["Macro", "Neutral", "Cycle intact but cooling; energy prices start to bite outside the US"],
        ["Market factors", "Neutral ↓", "Trend and technicals turned negative in US and Europe"],
        ["Liquidity", "Moderately negative ↓", "Policy and financial conditions neutral to mildly restrictive"]]
def hcol(ri, ci, v):
    if ci != 1: return None
    return MINT if v.startswith("Positive") else (TIGER if "negative" in v else SLATE)
table(s, L, 2350000, 5900000, [1.4, 1.5, 4.0], rows, size=9.5, row_h=430000, colors=hcol)
card(s, 6750000, 2350000, 3276000, 4500000, "Allocation stance", [
    "Equities: neutral – strong earnings vs higher real yields and credit stress",
    "Government bonds 1–10y: upgraded to neutral",
    "Corporates: neutral, prefer short maturities",
    "Commodities: downgraded to neutral (energy and industrial metals weaker)",
    "Gold: overweight – our preferred hedge against currency debasement",
    "USD: positive short term, weaker over the medium term"], accent=ORANGE, body_size=11)
tb(s, L, 5100000, 5900000, 1750000, [
    ("For your portfolio", 11, True, NAVY),
    ("Our research confirms the profile you chose: value and Europe were rewarded, gold stays a core holding, "
     "and the cautious view on industrial metals and energy matches your preferences. No changes are proposed in this review.", 10, False, NAVY)],
   fill=LIGHT, margin=90000)
source(s, "Source: Syz Investment Conclusions – Summary, November 2026 (five asset-allocation pillars and investment views).")

# ------------------------------------------------------------------ 13: divider
divider("Outlook 2027", "Macro scenario and investment themes")

# ------------------------------------------------------------------ 14: macro 2027
s = new_slide("2027 macro: a dip first, then a longer cycle",
              "Growth and inflation should moderate; central banks finish their hiking cycle; 'moderation' is our base case.")
cw3 = (W - 2 * 150000) // 3
card(s, L, 2350000, cw3, 2500000, "Growth", [
    "Early 2027 slowdown from rate hikes and lagged energy prices",
    "Re-acceleration later in the year as energy eases; capex and fiscal support outside the US"], accent=SKY, body_size=11.5)
card(s, L + cw3 + 150000, 2350000, cw3, 2500000, "Inflation & rates", [
    "Headline inflation close to peaking; should fall materially in 2027",
    "Fed, ECB, BoJ, BoE: ~2 more hikes to next summer; SNB one – not an aggressive cycle",
    "Main risk: a disorderly rise in long-term yields"], accent=ORANGE, body_size=11.5)
card(s, L + 2 * (cw3 + 150000), 2350000, cw3, 2500000, "Markets", [
    "Risk of a 10–15% correction, not a bear market: no red lights",
    "Year 3 of the US presidential cycle: positive but sub-average US returns, volatile H1"], accent=MINT, body_size=11.5)
header(s, L, 5000000, W, "Syz Macro Clock – 2027 scenarios")
rows = [["Moderation (base case) 45%", "Overheating → recession 25%", "Slowdown 20%", "Stagflationary trap 10%"],
        ["Equities (broadening), duration, EM, gold", "Cyclicals, value, commodities, banks", "Long sovereigns, quality, defensives, CHF", "Gold, trend-following, cash, TIPS"]]
table(s, L, 5300000, W, [1, 1, 1, 1], rows, size=10.5, row_h=600000, header_fill=GOLD)
source(s, "Source: Syz Outlook 2027 (Chief Investment Office); Syz Investment Conclusions, Nov 2026; Syz Research.")

# ------------------------------------------------------------------ 15: themes 2027
s = new_slide("2027 themes – and where your portfolio stands",
              "Five structural trends drive our five portfolio priorities. Your value and gold profile already fits three of them.")
rows = [["Portfolio priority", "What it means", "Your portfolio today"],
        ["1  Long innovation", "AI capex super-cycle; rotation within AI, not collapse", "Small by choice (~6%): cyber-security, two tech notes"],
        ["2  Long pricing power", "Grids, defence, physical infrastructure", "Strong: defence ~10%, insurers, Swisscom"],
        ["3  Own the source of inflation", "Commodities, infrastructure", "Via gold; industrial metals de-emphasised (house view neutral)"],
        ["4  Bonds – income first", "Corporates over sovereigns; manage duration", "Modest (~12%): IG corporates and EM bond certificate"],
        ["5  Rethinking diversification", "Low-beta stocks, hedge funds, gold, Bitcoin", "Gold core holding, Swiss defensives, small Bitcoin"]]
def tc(ri, ci, v):
    if ci != 2: return None
    return MINT if v.startswith(("Strong", "Gold", "Via")) else SLATE
table(s, L, 2350000, W, [1.6, 2.4, 2.8], rows, size=11, row_h=570000, colors=tc)
tb(s, L, 5900000, W, 950000, [
    ("Five structural trends behind them: ", 10, True, NAVY),
    ("a new geopolitical and economic order · surging sovereign debt and fiscal dominance · the AI capex super-cycle · "
     "the return of nominal growth · from demand to supply constraints.", 10, False, NAVY)], fill=LIGHT, margin=90000)
source(s, "Source: Syz Outlook 2027 – 'Five structural trends' and '5 portfolio priorities'; client portfolio statement (26 Nov 2026).")

# ------------------------------------------------------------------ 16: conclusions 2027
s = new_slide("Our investment conclusions for 2027",
              "Stay invested and rebalance away from index concentration – a backdrop that suits a diversified, value-oriented portfolio.")
cw4 = (W - 3 * 150000) // 4
card(s, L, 2350000, cw4, 2650000, "Regions", [
    "All neutral, with a US–Switzerland barbell",
    "Europe via mid-caps (13.8x); avoid France",
    "Japan: strongest trend, value leads"], accent=SKY, body_size=11.5)
card(s, L + cw4 + 150000, 2350000, cw4, 2650000, "Style & sectors", [
    "Broaden beyond mega-cap tech",
    "Health care, financials and industrials preferred",
    "Value led growth by ~15 pts in 2026"], accent=MINT, body_size=11.5)
card(s, L + 2 * (cw4 + 150000), 2350000, cw4, 2650000, "Duration & gold", [
    "Build duration gradually, 5–10y first",
    "Keep gold as the core hedge against currency debasement"], accent=GOLD, body_size=11.5)
card(s, L + 3 * (cw4 + 150000), 2350000, cw4, 2650000, "Hedges & credit", [
    "Add convexity: trend-following, macro",
    "Underweight private credit",
    "Keep digital assets small"], accent=MAUVE, body_size=11.5)
card(s, L, 5150000, W, 1700000, "What this means for you", [
    "Your European value equities, Swiss defensives and gold are aligned with our 2027 view",
    "Points we will keep monitoring together: the size of the gold position (~20%), USD exposure (~43%), "
    "the liquidity of private assets (~19%) and the silver DCIs at maturity"], accent=ORANGE, body_size=11.5)
source(s, "Source: Syz Outlook 2027 – 'Our investment conclusions for 2027' and regional / sector views (FactSet consensus, 3 Nov 2026).")

# ------------------------------------------------------------------ reorder
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
final_page = ids[template_n - 1]
keep = [ids[0]] + ids[template_n:] + [final_page]
for el in ids:
    if el not in keep:
        prs.part.drop_rel(el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'))
        sldIdLst.remove(el)
for el in keep[1:]:
    sldIdLst.remove(el); sldIdLst.append(el)
prs.save(OUT)
print("saved", OUT, len(prs.slides), "EST_YTD", round(EST_YTD, 2), {k: round(v, 2) for k, v in contrib.items()},
      "eq_eu", round(eq_eu, 1), "eq_rest", round(eq_rest, 1), "ai", round(ai_ret, 1), "eq_all", round(eq_all, 1))
