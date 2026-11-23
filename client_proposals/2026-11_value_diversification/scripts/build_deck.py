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

# ------------------------------------------------------------------ data prep
H = A["holdings"]
ac = A["asset_class"]
AC_ORDER = ["Equities", "Commodities", "Private assets", "Fixed income", "Structured products",
            "Cash & money market", "Digital assets", "Hedge funds"]
ac = {k: ac.get(k, 0) for k in AC_ORDER}
REG_LABEL = {"Eurozone": "Europe ex-UK & CH", "Commodities (no region)": "Gold & commodities (no region)",
             "Global": "Global / multi-region", "EM": "Emerging markets"}
region = {}
for k, v in A["region"].items():
    if v < 0.05: continue
    region[REG_LABEL.get(k, k)] = v
eq_region = {REG_LABEL.get(k, k): v for k, v in A["eq_region"].items()}
eq_sector = A["eq_sector"]
ccy = {}
for k, v in A["ccy"].items():
    if v < 0.05: continue
    ccy[{"XAG": "Silver (XAG)", "XPT": "Platinum (XPT)"}.get(k, k)] = v

def w_of(substr):
    return sum(h["weight"] for h in H if substr in h["desc"])

gold_phys = A["commodity_sub"].get("Gold (physical)", 0)
gold_miners = A["commodity_sub"].get("Gold miners", 0)
other_metals = A["commodity_sub"].get("Other metals & commodities", 0)
theme = A["theme"]

# --- proposal (weights in % of portfolio)
SELL = [  # (instrument, weight, reason)
    ("Silver account (XAG)", w_of("CURRENT ACCOUNT IN XAG"), "Other metal – client preference; house UW silver vs gold"),
    ("Platinum account (XPT)", w_of("CURRENT ACCOUNT IN XPT"), "Other metal – white metals in a bear market"),
    ("SWC Platinum ETF", w_of("PLATINUM ETF"), "Other metal – PGM outlook challenging"),
    ("16.77% DCI XAG-USD (matures 27.11.26)", w_of("DCI-XAGUSD"), "Do not roll – take proceeds in EUR"),
    ("BGF World Mining fund", w_of("BRGF WORLD MINING"), "Base-metal / mining beta, cyclical"),
    ("WisdomTree M&R Earths ETF", w_of("M&R EARTHS"), "Rare earths – volatile, −37% YTD"),
    ("WisdomTree Copper ETC", w_of("COPPER COM"), "House view: range-bound"),
    ("WisdomTree Aluminium ETC", w_of("ALUMIN"), "House view: bearish bias"),
    ("WisdomTree Broad Commodities (AIGCI)", w_of("AIGCI"), "Includes energy & base metals"),
    ("SG China Next Winners note", w_of("CHINA NEXT WINNERS"), "China tech / AI theme; house: avoid adding China"),
    ("VanEck Defense ETF (trim half)", w_of("VANECK VECTORS DEF") / 2, "Defence concentration ~9.5% of book"),
]
sell_total = sum(s[1] for s in SELL)
CASH_USE = round(A["asset_class"]["Cash & money market"] - 1.3, 2)   # keep ~1.3% liquidity
BUY = [  # (block, instrument type, weight, rationale)
    ("Fixed income", "Quality government bonds 5–10y (Germany, Switzerland, supranationals; no France)", 3.5, "Lock in yields, add duration gradually"),
    ("Fixed income", "Inflation-linked bonds (EUR linkers / EUR-hedged TIPS)", 2.0, "House OW; hedge if inflation persists"),
    ("Fixed income", "EUR investment-grade corporates 3–6y (ex-energy, ex-tech)", 1.5, "Carry with a quality buffer"),
    ("Hedge funds", "Global macro / trend-following (liquid UCITS)", 2.5, "Convexity; works in stagflation & sell-offs"),
    ("Equities", "Global health care – value-priced, defensive", 1.4, "House OW; defensive leadership"),
    ("Equities", "Japan value equities (EUR-hedged)", 1.1, "Strongest trend, value leads; 14x P/E"),
    ("Equities", "US equal-weight / US value (ex mega-cap tech)", 1.1, "US exposure without AI concentration"),
]
buy_total = sum(b[2] for b in BUY)
# balance: buys funded by sales + cash
assert abs(buy_total - (sell_total + CASH_USE)) < 0.6, (buy_total, sell_total, CASH_USE)
CASH_USE = buy_total - sell_total

prop = dict(ac)
prop["Commodities"] -= sum(s[1] for s in SELL[:9])
prop["Equities"] -= SELL[9][1] + SELL[10][1]
prop["Cash & money market"] -= CASH_USE
for blk, _, w, _ in BUY:
    prop[{"Fixed income": "Fixed income", "Hedge funds": "Hedge funds", "Equities": "Equities"}[blk]] += w

# ------------------------------------------------------------------ slide 1: cover (template)
cover = prs.slides[0]
for sh in cover.shapes:
    if sh.name == "Title 1":
        sh.text_frame.text = "Portfolio Review &\vDiversification Proposal"
    if sh.name == "Text Placeholder 5":
        tf = sh.text_frame
        lines = [f"Value-oriented EUR portfolio · EUR {TOTAL/1e6:.1f}m", "",
                 "Prepared for our client", "", "November 2026"]
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
s = new_slide("Executive summary",
              "A well-performing, value-tilted portfolio – but concentrated in gold, metals and private assets, with too little true diversification.")
cw = (W - 3 * 150000) // 4; y = 2450000; h = 4450000
card(s, L, y, cw, h, "Portfolio today", [
    f"EUR {TOTAL/1e6:.1f}m across {A['n']} positions",
    f"Commodities {pct(ac['Commodities'])} incl. physical gold {pct(gold_phys)} and other metals {pct(other_metals)}",
    f"Equities {pct(ac['Equities'])}, already value-tilted (insurers, Berkshire, Swiss income)",
    f"Private assets {pct(ac['Private assets'])}; fixed income only {pct(ac['Fixed income'])}",
    f"Hedge funds {pct(ac['Hedge funds'])}; USD {pct(A['ccy']['USD'])} of the book"], accent=SKY, body_size=12.5)
card(s, L + cw + 150000, y, cw, h, "Year to date", [
    "Winners: private equity (CoA Kona +46%), Swisscom, Digital Security, Berkshire, Allianz",
    "Losers: silver & platinum (−31% to −39%), M&R Earths (−37%), Bitcoin ETF, China Next Winners",
    "Gold −12% in EUR: a drag this year, but still a core long-term hedge"], accent=MINT, body_size=12.5)
card(s, L + 2 * (cw + 150000), y, cw, h, "Our market view", [
    "Cycle extended into 2027, with a dip first; moderation is our base case (45%)",
    "Value beats growth by ~15 pts YTD; AI trade now a debate on returns on capex",
    "Equities over bonds; lock in 5–10y yields; gold and hedge funds as diversifiers",
    "Unwind USD exposure"], accent=ORANGE, body_size=12.5)
card(s, L + 3 * (cw + 150000), y, cw, h, "Our proposal", [
    f"Sell other metals and mining ({pct(sum(x[1] for x in SELL[:9]))}) – keep all gold",
    f"Redeploy {pct(buy_total)} into quality bonds, inflation-linkers, macro/trend hedge funds and value equities",
    "No new AI or energy exposure",
    f"Fixed income {pct(ac['Fixed income'])} → {pct(prop['Fixed income'])}; hedge funds → {pct(prop['Hedge funds'])}"], accent=TIGER, body_size=12.5)
source(s, "Source: client portfolio statement (November 2026); Syz Outlook 2027; Syz Research Forum, 23 November 2026.")

# ------------------------------------------------------------------ 3: divider
divider("Portfolio today", "Allocation, exposures and year-to-date performance")

# ------------------------------------------------------------------ 4: overview
s = new_slide("Portfolio overview: allocation by asset class",
              "Commodities include silver and platinum accounts, BGF World Mining, VanEck Gold Miners, M&R Earths ETF and the DCI XAG-USD note.")
kw = (W - 4 * 120000) // 5; ky = 2380000
kpi(s, L, ky, kw, 720000, f"EUR {TOTAL/1e6:.1f}m", "Total value")
kpi(s, L + (kw + 120000), ky, kw, 720000, str(A["n"]), "Positions")
kpi(s, L + 2 * (kw + 120000), ky, kw, 720000, pct(gold_phys), "Physical gold")
kpi(s, L + 3 * (kw + 120000), ky, kw, 720000, pct(ac["Private assets"]), "Private assets")
kpi(s, L + 4 * (kw + 120000), ky, kw, 720000, pct(ac["Fixed income"] + ac["Hedge funds"]), "Bonds + hedge funds")
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
    f"Gold (bullion, Krugerrands, miners) is {pct(gold_phys+gold_miners)} of the book – far above the ~3% "
    "gold/commodities weight in our balanced SAA – and stays in place by client choice.", 9.5, False, NAVY)], margin=0)
source(s, "Source: client portfolio statement; weights as reported by the custodian. Classification of the six listed securities as commodities per client instruction.")

# ------------------------------------------------------------------ 5: region & currency
s = new_slide("Allocation by region and currency",
              "Excluding gold, the book is spread across Europe, the US and Switzerland; currency risk is dominated by the USD (largely the private-asset funds).")
header(s, L, 2350000, 4600000, "Whole portfolio by region")
reg_cols = [MINT if "Gold" in k else SLATE for k in region]
hbar(s, L, 2650000, 4600000, 3000000, region, colors=reg_cols)
header(s, 5418000, 2350000, 4600000, "Equity sleeve by region")
hbar(s, 5418000, 2600000, 4600000, 1950000, eq_region, color=SKY, size=8.5)
header(s, 5418000, 4600000, 4600000, "Currency exposure")
cc = {k: v for k, v in ccy.items() if v >= 2}
cc['Other'] = sum(v for k, v in ccy.items() if v < 2)
hbar(s, 5418000, 4880000, 4600000, 2150000, cc, color=GOLD, size=8.5, gap=40)
tb(s, L, 5750000, 4600000, 1250000, [
    ("What stands out", 11, True, NAVY),
    (f"US equities are only {pct(A['eq_region']['US']*ac['Equities']/100)} of the portfolio ({pct(A['eq_region']['US'])} of equities) – well below our strategic US weight; Switzerland and Europe dominate.", 9.5, False, NAVY),
    (f"USD is {pct(A['ccy']['USD'])} of assets – house view is to unwind USD exposure.", 9.5, False, NAVY)],
   fill=LIGHT, margin=90000)
source(s, "Source: client portfolio statement; region assigned at issuer / fund-mandate level by Syz. 'Global / multi-region' is mainly private-equity, hedge and thematic funds.")

# ------------------------------------------------------------------ 6: sectors & themes
s = new_slide("Sector allocation and thematic exposures",
              "Equities lean to financials and industrials (value); AI-linked exposure is modest, but defence has become a sizeable theme.")
header(s, L, 2350000, 4600000, "Equity sleeve by sector")
sec_cols = [SKY] * len(eq_sector)
hbar(s, L, 2650000, 4600000, 3300000, eq_sector, colors=sec_cols)
tb(s, L, 6000000, 4600000, 700000, [(
    "'Diversified' = broad equity funds (Swiss income, EU high dividend, EU mid-caps, US small caps). "
    "Mining and gold-miner funds are shown under commodities.", 8.5, False, GREY)], margin=0)
header(s, 5418000, 2350000, 4600000, "Thematic exposure (% of total portfolio)")
kw2 = (4600000 - 2 * 100000) // 3
kpi(s, 5418000, 2700000, kw2, 800000, pct(theme.get("Defence", 0)), "Defence", color=TIGER)
kpi(s, 5418000 + kw2 + 100000, 2700000, kw2, 800000, pct(theme.get("AI", 0)), "AI / mega-cap tech", color=SLATE)
kpi(s, 5418000 + 2 * (kw2 + 100000), 2700000, kw2, 800000, pct(theme.get("Energy", 0)), "Energy", color=MINT)
card(s, 5418000, 3650000, 4600000, 3200000, "Reading the exposures", [
    "Defence: Airbus, BAE, Dassault Aviation, Leonardo (share + note), Saab, VanEck Defense and Decalia Aegis – strong run, now a concentration",
    "AI: Meta & Amazon bonds, Google/Microsoft/Amazon and Microsoft notes, Digital Security ETF, China Next Winners – consistent with the client's low-AI preference; no adds proposed",
    "Energy: only the PKN Orlen bond. As the client's career is in energy, energy is not treated as a diversifier and no energy exposure is added"],
     accent=ORANGE, body_size=10.5)
source(s, "Source: client portfolio statement; sector and theme tags assigned by Syz at issuer / fund-mandate level (the custodian file has no sector field).")

# ------------------------------------------------------------------ 7: commodities sleeve detail
s = new_slide("The commodities sleeve in detail",
              "Gold is the anchor; the other metals have been the main source of losses this year.")
CNAME = {"GOLD KG 1 999.9 CHF": "Gold bullion, 35 x 1 kg", "BRGF WORLD MINING(REG)A2": "BGF World Mining fund",
         "OR KRUGERRAND(PCE)1OZ USD": "Krugerrand gold coins (70 x 1 oz)", "WISDOMTREE-ETC-ALUMIN PER": "WisdomTree Aluminium ETC",
         "WISDOMT.ETC-COPPER COM ID": "WisdomTree Copper ETC", "SWC(CH)PLATINUM ETF EACHF": "SWC Platinum ETF",
         "CURRENT ACCOUNT IN XPT": "Platinum account (XPT)", "WISDOMTREE-ETC-AIGCI(LDN)": "WisdomTree Broad Commodities ETC (AIGCI)",
         "WISD STRAT M&R EARTHS ETF": "WisdomTree Strategic Metals & Rare Earths ETF", "VANECK GOLD MINERS ETF": "VanEck Gold Miners ETF",
         "CURRENT ACCOUNT IN XAG": "Silver account (XAG)", "16,77%DCI-XAGUSD-58-UBS26  27.11.2026": "16.77% DCI XAG-USD (UBS), 27.11.2026"}
com = sorted([h for h in H if h["asset_class"] == "Commodities"], key=lambda h: -h["weight"])
rows = [["Holding", "Type", "Value (EUR)", "Weight", "YTD (EUR)", "Proposal"]]
keep = {"Gold (physical)", "Gold miners"}
for h in com:
    rows.append([CNAME.get(h["desc"], h["desc"]),
                 h["sub"], f"{h['value_eur']:,.0f}".replace(",", "'"), pct(h["weight"], 2), pct(h["perf_eur"], 1, True),
                 "Keep" if h["sub"] in keep else "Sell"])
rows.append(["Total commodities", "", f"{sum(h['value_eur'] for h in com):,.0f}".replace(",", "'"), pct(ac["Commodities"], 2), "", ""])
def ccol(ri, ci, v):
    if ci == 4 and isinstance(v, str) and v.startswith("-"): return RED
    if ci == 4 and isinstance(v, str) and v.startswith("+"): return MINT
    if ci == 5: return MINT if v == "Keep" else (TIGER if v == "Sell" else None)
table(s, L, 2400000, W, [3.6, 2.2, 1.3, 0.9, 1.0, 0.9], rows, size=9, row_h=300000, colors=ccol,
      bold_rows=(len(rows) - 1,), aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER])
source(s, "Source: client portfolio statement. YTD = time-weighted performance in EUR over the statement period (from purchase date for positions bought in 2026).")

# ------------------------------------------------------------------ 8: best & worst
s = new_slide("Best and worst performers year to date",
              "Private equity, Swiss defensives and insurers led; silver, platinum, rare earths and China tech lagged.")
ps = A["perf_sorted"]
def short(d):
    m = {"COA KONA III REF3 LP": "CoA Kona III (PE)", "ISHS DIGITAL SECURITY ETF": "iShares Digital Security",
         "SWISSCOM (NOM)": "Swisscom", "BERKSHIRE HATHAWAY B": "Berkshire Hathaway B", "ALLIANZ (NOM)": "Allianz",
         "BIL OC 211.5% NESN  13.12.2027": "Nestlé outperf. cert. 2027", "HANNOVER RUECK (NOM)": "Hannover Re",
         "BAE SYSTEMS": "BAE Systems", "UBS(CH) EQ-SWISS INC CHFP": "UBS Swiss Income Equity", "LEFT LANE CP II CLT2": "Left Lane II (PE)",
         "GOLD KG 1 999.9 CHF": "Gold bullion (kg)", "VANECK GOLD MINERS ETF": "VanEck Gold Miners",
         "16,77%DCI-XAGUSD-58-UBS26  27.11.2026": "DCI XAG-USD note", "SG CHINA NEXT WINNERS  01.06.2099": "SG China Next Winners",
         "DECALIA AEGIS DEFENSE A1": "Decalia Aegis Defense", "ISHARES BITCOIN ETF": "iShares Bitcoin ETF",
         "CURRENT ACCOUNT IN XPT": "Platinum (XPT) account", "SWC(CH)PLATINUM ETF EACHF": "SWC Platinum ETF",
         "WISD STRAT M&R EARTHS ETF": "WisdomTree M&R Earths", "CURRENT ACCOUNT IN XAG": "Silver (XAG) account"}
    return m.get(d, d.title())
top = {f"{short(d)}  ({w:.1f}%)": p for d, a_, w, p, q, lo in ps[:10]}
bot = {f"{short(d)}  ({w:.1f}%)": p for d, a_, w, p, q, lo in ps[-10:][::-1]}
header(s, L, 2350000, 4600000, "Top 10 – performance in EUR (weight)")
hbar(s, L, 2650000, 4600000, 4200000, top, color=MINT, num_fmt='+0.0"%";-0.0"%"', maxv=58)
header(s, 5418000, 2350000, 4600000, "Bottom 10 – performance in EUR (weight)")
hbar(s, 5418000, 2650000, 4600000, 4200000, bot, color=TIGER, num_fmt='+0.0"%";-0.0"%"', minv=-50, maxv=0)
source(s, "Source: client portfolio statement, time-weighted performance in EUR for the statement period (YTD, or since purchase for 2026 purchases); cash and fiduciary deposits excluded. Private-asset valuations are reported with a lag.")

# ------------------------------------------------------------------ 9: divider
divider("Our view on global markets", "Syz Outlook 2027 and Research Forum, 23 November 2026")

# ------------------------------------------------------------------ 10: where markets stand
s = new_slide("Where markets stand: value has beaten growth",
              "A rotation out of the year's leaders has turned into a market-wide correction led by AI/technology – a de-rating, not an earnings scare.")
header(s, L, 2350000, 4600000, "YTD total return (to 20 Nov 2026)")
mk = {"MSCI AC World Value": 13.3, "S&P 500 Equal Weight": 12.1, "Swiss SPI": 13.8, "STOXX Europe 600": 11.0,
      "MSCI Japan (USD)": 11.3, "MSCI AC World": 5.8, "S&P 500": 5.9, "NASDAQ 100": 5.6,
      "MSCI AC World Growth": -1.4, "MSCI China": -20.6}
hbar(s, L, 2650000, 4600000, 2900000, mk, colors=[MINT if v > 7 else (SLATE if v >= 0 else TIGER) for v in mk.values()],
     num_fmt='+0.0"%";-0.0"%"', size=8.5, minv=-38, maxv=20)
header(s, L, 5600000, 4600000, "World sectors YTD (local)")
secs = {"Energy": 36.8, "Cons. staples": 16.4, "Health care": 13.3, "Information tech.": 8.3, "Comm. services": -11.0}
hbar(s, L, 5850000, 4600000, 1200000, secs, colors=[MINT if v > 0 else TIGER for v in secs.values()],
     num_fmt='+0.0"%";-0.0"%"', size=8, gap=30, minv=-30, maxv=45)
obs = [("Value leads growth by ~15 points",
        "ACWI Value +13.3% vs Growth −1.4%; the equal-weighted S&P 500 (+12.1%) beats the cap-weighted index (+5.9%). Concentration is unwinding."),
       ("AI: from tailwind to debate",
        "2027 hyperscaler capex ~$1.3tn and still revised up, but a ~$595bn funding gap; Nvidia −33% since 6 Nov. Markets now ask for returns on AI spending."),
       ("Valuations back to average",
        "S&P 500 at 17.2x forward earnings (23x a year ago), Stoxx 600 13.8x, Japan 14.3x – while forward EPS keeps hitting records ($418)."),
       ("Gold: strong structurally, broken tactically",
        "Gold at $3,711 after fresh lows; central banks bought 675t Jan–Sep and 45% plan to add more. Core holding.")]
oy = 2350000
for t, b in obs:
    card(s, 5418000, oy, 4600000, 1080000, t, ["§" + b], accent=ORANGE, title_size=11, body_size=9.5)
    oy += 1160000
source(s, "Source: Syz Research Forum, 23 November 2026 (FactSet, Bloomberg). Total return gross; USD unless stated. Past performance is not a guide to future returns.")

# ------------------------------------------------------------------ 11: 2027 outlook
s = new_slide("Outlook 2027: a dip, then a longer cycle",
              "Five structural trends shape a new macro regime; our Syz Macro Clock makes 'moderation' the base case.")
header(s, L, 2350000, 4600000, "Five structural trends")
trends = ["A new geopolitical and economic order", "Surging sovereign debt & fiscal dominance",
          "The AI capex super-cycle", "The return of nominal growth", "From demand to supply constraints"]
for i, t in enumerate(trends):
    yy = 2700000 + i * 430000
    tb(s, L, yy, 380000, 360000, [(str(i + 1), 14, True, WHITE, PP_ALIGN.CENTER)], fill=TIGER, anchor=MSO_ANCHOR.MIDDLE, margin=0, shape=MSO_SHAPE.OVAL)
    tb(s, L + 470000, yy, 4100000, 360000, [(t, 11, False, NAVY)], anchor=MSO_ANCHOR.MIDDLE, margin=0)
card(s, L, 4950000, 4600000, 1950000, "Our answers for 2027", [
    "Growth: cycle extended, with a dip first; 2026's 'stagflation-lite' gives way to moderation",
    "Inflation close to peaking; the risk is a disorderly rise in long yields",
    "Risk of a 10–15% correction, not a bear market",
    "AI: a rotation within the theme, not a collapse"], accent=SLATE, body_size=9.5)
header(s, 5418000, 2350000, 4600000, "Syz Macro Clock – scenario probabilities")
rows = [["Scenario", "Prob.", "What works"],
        ["Moderation (base case)", "45%", "Equities (broadening), duration, EM, gold"],
        ["Overheating → recession", "25%", "Cyclicals, value, commodities, banks; short duration"],
        ["Slowdown", "20%", "Long sovereigns, quality, defensives, USD/CHF/JPY"],
        ["Stagflationary trap", "10%", "Gold, energy, trend-following, cash, TIPS"]]
table(s, 5418000, 2700000, 4600000, [1.6, 0.55, 2.6], rows, size=9.5, row_h=520000,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT])
tb(s, 5418000, 5450000, 4600000, 1300000, [
    ("Implication", 11, True, NAVY),
    ("No single asset wins in every regime. Diversification means owning duration, gold and trend-following "
     "alongside value equities – exactly the gaps in this portfolio today.", 10, False, NAVY)], fill=LIGHT, margin=90000)
source(s, "Source: Syz Outlook 2027 (Chief Investment Office); Syz Research.")

# ------------------------------------------------------------------ 12: house TAA
s = new_slide("Our 2027 tactical asset allocation",
              "Still favour equities over bonds; keep gold and hedge funds as diversifiers; lock in 5–10y yields; trim the dollar.")
rows = [["Asset class", "House view", "Comment"],
        ["Global equities", "Moderate overweight", "On pullbacks; broaden beyond index concentration"],
        ["Cash", "Move to underweight", "Progressively"],
        ["Government bonds 1–10y", "To neutral, in stages", "Start with 5–10y; long bonds only on yield spikes"],
        ["Inflation-linked bonds", "Overweight", ""],
        ["Investment-grade credit", "Neutral", "Attractive carry, tight spreads"],
        ["High yield / private credit", "Underweight", "Prefer short maturities"],
        ["EM debt", "Favoured FI segment", "Short & medium maturities"],
        ["Hedge funds (macro, trend)", "Neutral / overweight", "Add convexity"],
        ["Gold", "Overweight", "Core hedge against currency debasement"],
        ["Private markets", "Selective overweight", ""],
        ["Digital assets", "Small strategic", ""],
        ["US dollar", "Unwind exposure", "EUR/USD mildly positive over 6m; CHF to strengthen"]]
def tcol(ri, ci, v):
    if ci != 1: return None
    if "nder" in v or "Unwind" in v: return TIGER
    if "verweight" in v or "avoured" in v or "neutral, in" in v: return MINT
table(s, L, 2350000, 5700000, [2.0, 1.7, 3.0], rows, size=9, row_h=330000, colors=tcol)
card(s, 6550000, 2350000, 3476000, 2150000, "Regions: all neutral", [
    "US–Switzerland barbell",
    "Japan: strongest trend, value leads (14.1x)",
    "Europe via mid-caps (13.8x); avoid France",
    "China / HK: avoid adding"], accent=SKY, body_size=9.5)
card(s, 6550000, 4600000, 3476000, 2280000, "Sectors (US)", [
    "Growth leg: IT, Industrials, Health care",
    "Hedge leg: Energy, Financials",
    "Underweight: Cons. discretionary, Staples, Real estate",
    "§For this client: we keep Health care, Financials, Industrials and skip the IT and Energy legs."], accent=MAUVE, body_size=9.5)
source(s, "Source: Syz Investment Committee – 2027 TAA (Outlook 2027); Syz Research Forum, 23 Nov 2026 (FI & FX views). P/E = 2027 consensus (FactSet).")

# ------------------------------------------------------------------ 13: commodities view
s = new_slide("Commodities: keep gold, step away from the other metals",
              "Our research supports the client's preference: gold stays a core holding while white and base metals face a weaker set-up.")
rows = [["Commodity", "Spot 19 Nov", "Sell-side 2027", "Syz 12-month view", "Technical", "For this portfolio"],
        ["Gold ($/oz)", "3,711", "4,775–5,100", "Core holding", "Bearish", "Keep; do not add for now"],
        ["Silver ($/oz)", "45.3", "64–78", "Underweight vs gold", "Bearish", "Sell"],
        ["Platinum ($/oz)", "1,375", "1,920–2,040", "Preferred PGM", "Bearish", "Sell (client preference)"],
        ["Copper ($/t)", "12,471", "12,500–14,700", "Range-bound", "Neutral", "Sell"],
        ["Aluminium ($/t)", "2,823", "2,800–3,175", "Bearish bias", "Bearish", "Sell"],
        ["Oil – Brent ($/bbl)", "93.1", "63–88", "Conflict-driven", "Neutral", "No exposure (energy)"]]
def ccol2(ri, ci, v):
    if ci == 4: return TIGER if v == "Bearish" else SLATE
    if ci == 5: return MINT if v.startswith("Keep") else (TIGER if v.startswith("Sell") else SLATE)
table(s, L, 2400000, W, [1.6, 1.0, 1.3, 1.5, 0.9, 1.9], rows, size=9.5, row_h=340000, colors=ccol2)
cw3 = (W - 2 * 150000) // 3
card(s, L, 4950000, cw3, 1950000, "Why keep gold", [
    "Core hedge against currency debasement",
    "Central banks bought 675t Jan–Sep; GS sees ~60t/month",
    "Wait for the tactical trend to repair before adding"], accent=GOLD, body_size=9.5)
card(s, L + cw3 + 150000, 4950000, cw3, 1950000, "Why sell the rest", [
    "White metals in a bear market; PGM demand hit by EV/hybrid shift",
    "Copper: AI-demand doubts and tariff risk",
    "Mining & rare earths add equity beta, not diversification"], accent=TIGER, body_size=9.5)
card(s, L + 2 * (cw3 + 150000), 4950000, cw3, 1950000, "Execution note", [
    "Sell-side targets sit above spot: selling locks in this year's losses",
    "Stage the sales over 4–6 weeks and use rebounds",
    "Let the DCI XAG-USD note mature on 27.11 and not roll it"], accent=SLATE, body_size=9.5)
source(s, "Source: Syz Research Forum, 23 Nov 2026 – commodities section (sell-side consensus range, Bloomberg, World Gold Council, Syz Group).")

# ------------------------------------------------------------------ 14: divider
divider("Diversification proposal", "Keeping the value approach, with limited AI and no energy")

# ------------------------------------------------------------------ 15: guidelines
s = new_slide("Client guidelines shape the proposal",
              "Five principles translate the client's preferences and our house view into portfolio changes.")
pr = [("1  Stay value-oriented", "Add only reasonably priced assets: health care, Japan value, US equal-weight; keep the insurers, Berkshire and Swiss income funds.", SKY),
      ("2  Limit AI", f"No new AI or mega-cap tech exposure; exit the China Next Winners note. AI-linked exposure falls from {pct(theme.get('AI',0))} to {pct(theme.get('AI',0)-SELL[9][1])}.", SLATE),
      ("3  Keep gold, sell the other metals", f"Gold, Krugerrands and gold miners ({pct(gold_phys+gold_miners)}) stay; {pct(sum(x[1] for x in SELL[:9]))} in silver, platinum, base metals, mining and rare earths is sold.", GOLD),
      ("4  Energy is not a diversifier", "The client already depends on the energy sector through his career, so we add none. Gold, inflation-linkers and trend-following take over its hedging role.", MINT),
      ("5  Diversify what is missing", "Add the drivers this book lacks: duration, inflation protection and uncorrelated hedge-fund returns – and reduce USD and defence concentration.", TIGER)]
yy = 2400000
for t, b, c in pr:
    card(s, L, yy, W, 830000, t, ["§" + b], accent=c, title_size=12, body_size=10)
    yy += 910000
source(s, "Source: client preferences as communicated; Syz Outlook 2027.")

# ------------------------------------------------------------------ 16: transactions
s = new_slide("Proposed transactions",
              f"About {eur_m(buy_total)} ({pct(buy_total)} of the portfolio) moves from metals, China tech, defence and surplus cash into diversifiers.")
rows = [["Sell / reduce", "Weight", "EUR", "Why"]]
for n, w, why in SELL:
    rows.append([n, pct(w, 2), f"{w/100*TOTAL:,.0f}".replace(",", "'"), why])
rows.append(["Surplus cash & money market", pct(CASH_USE, 2), f"{CASH_USE/100*TOTAL:,.0f}".replace(",", "'"), "Cash to underweight; ~1.3% kept"])
rows.append(["Total funding", pct(sell_total + CASH_USE, 2), f"{(sell_total+CASH_USE)/100*TOTAL:,.0f}".replace(",", "'"), ""])
table(s, L, 2330000, 5150000, [2.6, 0.75, 0.95, 2.9], rows, size=8, row_h=290000, header_fill=TIGER,
      bold_rows=(len(rows) - 1,), aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT])
rows = [["Buy", "Weight", "EUR"]]
for blk, n, w, why in BUY:
    rows.append([n, pct(w, 1), f"{w/100*TOTAL:,.0f}".replace(",", "'")])
rows.append(["Total investments", pct(buy_total, 1), f"{buy_total/100*TOTAL:,.0f}".replace(",", "'")])
fills = {1: RGBColor(0xE8, 0xEC, 0xF3), 2: RGBColor(0xE8, 0xEC, 0xF3), 3: RGBColor(0xE8, 0xEC, 0xF3),
         4: RGBColor(0xEC, 0xE9, 0xF5), 5: RGBColor(0xE6, 0xF5, 0xFC), 6: RGBColor(0xE6, 0xF5, 0xFC), 7: RGBColor(0xE6, 0xF5, 0xFC)}
table(s, 5950000, 2330000, 4076000, [3.0, 0.7, 0.9], rows, size=8.5, row_h=400000, header_fill=MINT, fills=fills,
      bold_rows=(len(rows) - 1,), aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
tb(s, 5950000, 6000000, 4076000, 700000, [(
    "New purchases in EUR or CHF, or EUR-hedged share classes. Fund selection follows the Syz approved list.", 8.5, False, GREY)], margin=0)
source(s, f"Weights in % of the portfolio (EUR {TOTAL/1e6:.1f}m at statement valuation); amounts indicative and subject to market prices at execution.")

# ------------------------------------------------------------------ 17: current vs proposed
s = new_slide("Current vs proposed allocation",
              "Commodities shrink to the gold core; fixed income and hedge funds rise; equities stay broadly unchanged but more diversified.")
cats = AC_ORDER
header(s, L, 2350000, 5600000, "Allocation by asset class (% of portfolio)")
hbar(s, L, 2600000, 5600000, 4350000, cats, series=[("Current", [ac[c] for c in cats]), ("Proposed", [prop[c] for c in cats])],
     series_colors=[RGBColor(0xC5, 0xCB, 0xD8), TIGER], legend_on=True, gap=50, maxv=40)
rows = [["", "Current", "Proposed"]]
gold_all = gold_phys + gold_miners
usd_sold = sum(w_of(k) for k in ["BRGF WORLD MINING", "COPPER COM", "ALUMIN", "AIGCI", "CHINA NEXT WINNERS"])
metal_ccy_sold = w_of("CURRENT ACCOUNT IN XAG") + w_of("CURRENT ACCOUNT IN XPT") + w_of("DCI-XAGUSD")
rows += [["Gold (incl. miners)", pct(gold_all), pct(gold_all)],
         ["Other metals & commodities", pct(other_metals), "0.0%"],
         ["Bonds + hedge funds", pct(ac["Fixed income"] + ac["Hedge funds"]), pct(prop["Fixed income"] + prop["Hedge funds"])],
         ["AI-linked", pct(theme.get("AI", 0)), pct(theme.get("AI", 0) - SELL[9][1])],
         ["Defence", pct(theme.get("Defence", 0)), pct(theme.get("Defence", 0) - SELL[10][1])],
         ["Energy", pct(theme.get("Energy", 0)), pct(theme.get("Energy", 0))],
         ["USD-denominated", pct(A["ccy"]["USD"]), pct(A["ccy"]["USD"] - usd_sold)],
         ["Cash & money market", pct(ac["Cash & money market"]), pct(prop["Cash & money market"])]]
def pcol(ri, ci, v):
    return TIGER if ci == 2 else None
table(s, 6450000, 2600000, 3576000, [2.2, 1.0, 1.0], rows, size=9.5, row_h=400000, colors=pcol,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
tb(s, 6450000, 6300000, 3576000, 650000, [(
    "USD falls further if purchases are made in EUR/CHF or hedged share classes.", 8.5, False, GREY)], margin=0)
source(s, "Source: client portfolio statement; Syz calculations. Proposed weights at statement valuation, before market moves and transaction costs.")

# ------------------------------------------------------------------ 18: diversification by scenario
s = new_slide("How the proposal diversifies across scenarios",
              "Each Syz Macro Clock regime now has assets in the portfolio that can work in it – without leaning on energy or AI.")
rows = [["Scenario (probability)", "What works (house view)", "Today", "After the proposal"],
        ["Moderation (45%)", "Broadening equities, duration, EM, gold", "Value equities, gold, EM bonds; little duration",
         "+ 5–10y quality govies, US equal-weight, Japan value"],
        ["Overheating → recession (25%)", "Value, cyclicals, banks, short duration", "Insurers, banks, Berkshire, industrials",
         "+ inflation-linkers; short-dated IG"],
        ["Slowdown (20%)", "Long sovereigns, quality, defensives, CHF", "Swiss defensives, CHF assets",
         "+ government bonds, global health care"],
        ["Stagflationary trap (10%)", "Gold, energy, trend-following, TIPS", "Gold only",
         "+ trend-following / macro hedge funds, inflation-linkers (energy replaced)"]]
def scol(ri, ci, v):
    return MINT if ci == 3 else None
table(s, L, 2400000, W, [1.7, 2.1, 2.2, 2.6], rows, size=9.5, row_h=640000, colors=scol)
card(s, L, 5800000, W, 1050000, "The result", [
    "§A portfolio less dependent on one asset (gold) and one illiquid block (private assets), with more liquid diversifiers, lower USD risk and the client's value bias intact."],
     accent=MINT, body_size=10.5)
source(s, "Source: Syz Outlook 2027 – Syz Macro Clock; Syz calculations.")

# ------------------------------------------------------------------ 19: risks & next steps
s = new_slide("Risks, implementation and next steps", "What could go wrong, and how we would phase the changes.")
cw2 = (W - 150000) // 2
card(s, L, 2450000, cw2, 4450000, "Risks to watch", [
    "Gold stays ~20% of assets: a large single exposure, accepted by the client; tactical trend still bearish",
    "Selling metals after a drawdown: a sharp rebound in silver or platinum would be missed",
    "Duration: a disorderly rise in long yields (US 10y above 5.25%) – hence 5–10y first, long bonds only on spikes",
    "Hedge funds: manager selection and fees; lower liquidity than ETFs",
    "Private assets (~19%) stay illiquid; further capital calls of uncalled commitments to be planned"],
     accent=TIGER, body_size=12)
card(s, L + cw2 + 150000, 2450000, cw2, 4450000, "Implementation plan", [
    "Week 1: let the DCI XAG-USD note mature (27.11); sell copper, aluminium and AIGCI ETCs; convert proceeds to EUR",
    "Weeks 1–4: stage the silver, platinum, mining and rare-earth sales on rebounds",
    "Weeks 2–6: build the bond ladder (5–10y) and linkers; add longer bonds only on yield spikes",
    "Weeks 2–8: add hedge funds and equities, on market pullbacks where possible",
    "Review together at the next quarterly meeting; check tax impact of realised gains/losses before execution"],
     accent=MINT, body_size=12)

# ------------------------------------------------------------------ reorder: cover, new slides, final page; drop the rest
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
final_page = ids[template_n - 1]          # "Final page - blank"
keep = [ids[0]] + ids[template_n:] + [final_page]
for el in ids:
    if el not in keep:
        rId = el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        prs.part.drop_rel(rId)
        sldIdLst.remove(el)
for el in keep[1:]:
    sldIdLst.remove(el); sldIdLst.append(el)
prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides), "buy", round(buy_total, 2), "sell", round(sell_total, 2), "cash", CASH_USE)
print({k: round(v, 2) for k, v in prop.items()}, round(sum(prop.values()), 2))
