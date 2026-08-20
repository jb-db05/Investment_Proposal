# Slide-by-slide recipe

What each of the template's 36 slides is, where its content comes from, and
which build-script function fills it. "Static" means the slide ships as-is
from `assets/template.pptx` — no per-client data goes on it.

| # | Slide | Source | Filled by |
|---|-------|--------|-----------|
| 1 | Cover | `parsed.json` (profile, currency, amount, client name) | `build_proposal.fill_cover` |
| 2 | Agenda | static | — |
| 3 | How markets moved last week | `market_update.json: headline, scoreboard, three_observations` | `build_proposal.fill_market_slides` |
| 4 | What drove the move: four forces | `market_update.json: four_drivers` | `build_proposal.fill_market_slides` |
| 5 | Our economic scenario | **frozen — never updated, see note below** | — |
| 6 | Our current investment views | **frozen — never updated, see note below** | — |
| 7 | Section divider ("Markets this week") | static | — |
| 8 | Portfolio strategies / risk-return trade-off | `parsed.json: risk_profile`, `slide8_profile_dial.md` | `build_proposal.fill_profile_dial` |
| 9-N | Proposed portfolio: full holdings list | `parsed.json: line_items` | `build_line_items_tables.build` (runs **last** — see note below) |
| N+1 | Portfolio overview | `parsed.json`: KPIs, `asset_allocation_pct`, `currency_exposure_pct` | `build_proposal.fill_portfolio_overview` |
| N+2 | Geographic & sector exposure | `parsed.json: geographic_exposure_pct`, `sector_exposure_pct` (may be unavailable, see assumptions.md §7) | `build_proposal.fill_geo_sector` |
| N+3 | Fixed Income: income-type specifics | `parsed.json: sleeves["Fixed Income"]` | `build_proposal.fill_sleeve_slides` |
| N+4 | Fixed Income Breakdown (proposed bond selection) | `parsed.json: proposed_bond_selection` (from the `Fixed Income` Excel tab — a curated proposal, NOT current holdings) | `build_proposal.fill_proposed_bond_selection` |
| N+5 | Equities: income-type specifics | `parsed.json: sleeves["Equities"]` | `build_proposal.fill_sleeve_slides` |
| N+6 | Equity breakdown | `parsed.json: equity_breakdown` | `build_proposal.fill_equity_breakdown` |
| N+7 | Alternatives: income-type specifics | `parsed.json: sleeves["Private Assets"]` (closest conceptual match — see note below) | `build_proposal.fill_sleeve_slides` |
| N+8 | Commodities | `parsed.json: sleeves["Commodities"]` | `build_proposal.fill_sleeve_slides` |
| N+9 | Structured products | `parsed.json: sleeves["Structured Products"]` | `build_proposal.fill_sleeve_slides` |
| N+10 | Liquidity profile | `parsed.json: liquidity_profile_pct` | `build_proposal.fill_liquidity` |
| N+11 | Concentration & top holdings | `parsed.json: concentration` | `build_proposal.fill_concentration` |
| N+12 | Income & interest-rate sensitivity | `parsed.json: income` | `build_proposal.fill_income` |
| N+13 | Appendix divider | static | — |
| N+14 | Syz Group at a Glance | static (firm collateral) | — |
| N+15 | House view TAA (blank preference-matrix template) | static — advisor fills in manually | — |
| N+16..36 | Investment universe, advisory offering, fund recommendation one-pagers | static (bank collateral, not client-specific) | — |

**N** is 10, 11, or higher depending on how many holdings-list slides a
given portfolio needs (assumptions.md's pagination rule) — that's why
`build_line_items_tables.build()` must run **last** in `build_proposal.py`:
every other `fill_*` function addresses a slide by its fixed original
template position (`get_slide(prs, 11)` for Portfolio Overview, etc.), and
those positions are only correct before the holdings-list slide count
changes. This bit us once during development — see the comment above the
call order in `build_proposal.build()`.

## Scope boundary

The appendix (Syz Group at a Glance, advisory service tiers, investment
universe, house-view TAA worksheet) is **not** data-driven by this skill.
It's bank collateral / marketing content that doesn't come from the
client's portfolio or the market PDF, and none of the three inputs (Excel,
PDF, risk profile) contains what would be needed to personalize it. It
ships as-is from the template.

The five individual fund-recommendation one-pagers that used to follow the
appendix (BNY Mellon Global Short-Dated High Yield, UBS Swiss Income
Equity Fund + its characteristics page, Schroder US Large Cap Fund + its
characteristics page) were **removed from the template entirely** per
explicit instruction — they were specific fund picks tied to one advisor's
recommendations, not generic collateral appropriate for every client. If a
future need brings fund one-pagers back, that's a new input (which funds
to recommend, keyed by risk profile or client), not something to
reconstruct from the three existing inputs.

## "Alternatives" naming

The reference deck's slide 17 is titled "Alternatives: income-type
specifics" and in the original example held a single structured note
("Scarce Assets Certificate"). This skill's asset-class taxonomy (see
`assumptions.md` §1, `ASSET_CLASS_MAP` in `parse_portfolio.py`) doesn't
produce an "Alternatives" bucket — the custodian export's own section
headers split that space into `Structured products` and `Private Assets`.
`sleeves["Private Assets"]` is used to fill slide 17 as the closest
conceptual match (illiquid, less-correlated holdings), while `Structured
Products` gets its own slide (19) using its own section header text
verbatim. If your desk's taxonomy genuinely needs a combined "Alternatives"
sleeve, merge those two `ASSET_CLASS_MAP` targets in `parse_portfolio.py`
rather than special-casing it in the slide filler.

## Slides 5-6 are permanently frozen

Per explicit instruction, "Our economic scenario" and "Our current
investment views" always ship exactly as the reference deck has them —
`build_proposal.fill_market_slides()` never touches them, regardless of
what `market_update.json` contains. Only slides 3-4 are dynamic. If a
future requirement wants these unfrozen again, that's a one-line change
(re-add the old scenario/house_view fill calls), not a redesign — the
`market_update.json` schema below still documents the `scenario` and
`house_view` fields in case that day comes, even though nothing currently
reads them.

## Table column simplification (sleeve "largest holdings" tables)

The reference deck used a different table column set per sleeve (Fixed
Income: Holding/Value/Share/Duration/Yield/Price/Bid-ask; Equities:
Holding/Value/Share/Country/Sector/Market Cap — with Country/Sector/Market
Cap always blank, since that data isn't in the client file). This skill
uses **Holding / Value / Share**, plus a **Yield / Coupon** column only on
sleeves where that data can be real — Fixed Income and Structured Products
(slides 13, 19). Equities, Alternatives/Private Assets and Commodities
(slides 15, 17, 18) drop that column entirely rather than show it always
blank. `build_proposal.rebuild_holdings_table(..., include_yield_column=...)`
replaces the template's fixed 1-5-row table with one sized to the sleeve's
actual holding count on every build — the template's own tables can't grow.

## Line-items table layout (slides 9-10+)

Each top-level asset class gets **one** colored header row carrying both
its name (left) and its total weight (right) on the same line — never a
separate subtotal row underneath its holdings. This matches the reference
deck's own slide 9-10 design (a colored bar with the category name and %
together) rather than the earlier draft of this skill, which added a
"Subtotal — X" row after every group. Finer sub-groups (a fixed-income
bucket, an equity region) still show their own label in the merged first
column, but get no subtotal of their own — that breakdown lives on the
sleeve slides (13-19) instead, so showing it twice would be redundant.

Category bar colors (`CATEGORY_COLORS` in `build_line_items_tables.py`):
Cash = gold (`FFC545`, theme `bg2`/`lt2`), Fixed Income = slate (`4B5F80`,
`accent6` — the reference deck's own color, which originally doubled for
Equities too), Private Assets (filling the "Alternatives" slot) = mauve
(`AC5D85`, `accent5`), Commodities = mint (`3BAF90`, `accent2`), Total =
navy (`202945`, `tx2`/`dk2`) — all four pulled directly from the reference
deck's own Rectangle shapes via their theme color references. Equities and
Structured Products depart from the reference deck by explicit design
direction: Equities = Sky Blue (`79D6FF`), Structured Products = Peach
Pink (`E6A4AD`, not a theme color — a one-off custom value), both chosen
to be visually distinct from every other bar rather than reusing a theme
accent. `Other` still falls back to `FFA400` (`accent3`) since no explicit
color was ever specified for it.

## Formatting rules that must survive every rebuild

These are enforced in code (`set_shape_lines()` in `build_proposal.py`),
not just convention — breaking them was a real bug caught during review:

- **Never collapse a shape's paragraphs to one style.** A stat box's
  number and its caption, or a driver card's title and body, are separate
  paragraphs with their own font/size/color in the template. Setting text
  must preserve each paragraph's own formatting by position, only cloning
  a new paragraph (from the last existing one) when there are more lines
  than the shape already has.
- **Slide 3 scoreboard**: each Week/YTD value is colored green (`3BAF90`,
  `accent2`) if positive, tiger-orange (`FF6C0E`, `accent4`) if negative,
  based on the value string's own sign — never the header row.
- **Slide 4 driver cards**: only the title line (`"N - Title"`) keeps
  whatever bold the template gave it; every other line (body, inline stat)
  is forced non-bold, overriding the template if needed.
- **Slide 8 risk-return dial**: the selected profile's dot turns orange
  (`FFA400`, `accent3`, the template's own highlight color — previously
  hardcoded onto "Moderate" regardless of which profile was actually
  chosen) and grows larger (`DOT_SIZE_HIGHLIGHT`); every other dot,
  Moderate included, reverts to the default teal (`3AB7C8`, a literal
  color on the ovals, not a theme reference) at the default size
  (`DOT_SIZE_DEFAULT`). Resizing keeps each dot's own center point fixed
  (`_resize_dot_keep_center()`) so it stays anchored to the curve instead
  of drifting when it grows/shrinks. See `PROFILE_DOT_SHAPE` in
  `build_proposal.py` for the profile → oval-shape mapping (they live
  inside a group named `Group 3`).
- **Subtitle placeholder ("Text Placeholder 2") is never edited in
  place.** Its layout definition uses `<a:spAutoFit/>` with `anchor="b"`
  (bottom-anchored, shape grows to fit text) — a subtitle that wraps to a
  second line visibly shifts *down* instead of the box growing upward.
  `set_subtitle()` (in both `build_proposal.py` and
  `build_line_items_tables.py`) deletes the placeholder and recreates it
  as a plain textbox at the template's own position (H=1.85cm/V=4.8cm,
  i.e. `left=665163, top=1728947` EMU), fixed height, top-anchored, no
  autofit — so it never moves regardless of line count. Apply this same
  pattern to any new subtitle-setting code; editing the placeholder's
  runs in place will reintroduce the bug.
- **Internal data-sourcing footnotes are blanked, not translated.** Two
  footnotes reveal backend methodology the client doesn't need to see:
  the proposed-bond-selection slide's "Source: 'Fixed Income' tab..."
  (`Text 41`) and the equity-breakdown slide's "classified at issuer
  level... these fields are not in the client file" (`Text 32`). Both are
  set to `""` in `fill_proposed_bond_selection()` /
  `fill_equity_breakdown()` rather than removed as shapes, so the layout
  doesn't shift.

## Template-level fixes (applied once, not per-build)

Two things were wrong with `assets/template.pptx` itself, inherited from
the source deck, and got fixed directly on the template rather than
patched around at build time — every deck built from it inherits the fix
automatically:

- **Leftover reviewer comments.** The source deck (`work16.pptx`) carried
  three Google-Slides-export PowerPoint comments plus their
  `commentAuthors.xml`/`authors.xml`/`revisionInfo.xml` parts — irrelevant
  past review notes, not client content. Stripped entirely (parts,
  content-type entries, relationships, and the `commentRel` extension
  blocks on the three slides that referenced them).
- **Literal black instead of the brand navy.** Several table headers
  (sleeve "largest holdings" tables, the income tables) had their text
  color hardcoded as `000000` instead of the theme's actual dark color,
  `202945` (`dk2` in the "Syz 2024" theme — the color everything else on
  the deck actually uses for dark text). Every literal `000000` inside
  `ppt/slides/slideN.xml` was replaced with `202945`; the theme's own
  `dk1` slot (still `000000`, used correctly elsewhere) was left alone.

- **Slide 8's compass/gauge icon(s)** — there were actually two separate
  decorative shapes stacked on top of each other: `Graphic 8` (a Freeform
  group) and `Graphic 54` (an embedded SVG picture literally named
  "Speedometer Middle with solid fill" in its own XML). Both removed —
  they added visual noise without meaning. If a rendered check still shows
  an icon there, look for a third one; there is no principled reason to
  assume two was the final count.
- **A stray empty chart on the Equities sleeve slide** — template slide 15
  ("Equities: income-type specifics") carried a second `graphicFrame`,
  `Chart 14`, stacked almost exactly on top of the real vehicle-breakdown
  donut `Chart 9` (offsets 6.12"/2.93" vs 6.12"/2.95", near-identical
  size). Its chart part (`ppt/charts/chart11.xml`) had zero series and zero
  categories — an empty plot frame that nothing filled, since the builder
  only ever writes `Chart 9` on sleeve slides. Every other sleeve slide
  (Alternatives, Commodities) has `Chart 9` alone, which is what made this
  one an artifact rather than a design. Removed: the `graphicFrame`, the
  `rId4` relationship on `slide15.xml.rels`, the orphaned `chart11.xml` and
  its `_rels`, the now-unreferenced `embeddings/Microsoft_Excel_Worksheet10.xlsx`,
  and the `[Content_Types].xml` override for `chart11.xml`. This one was
  invisible to the text-level QA in `SKILL.md` step 6 (an empty chart
  contributes no text) and to `validate.py` (an unused-but-well-formed part
  is not an error) — it only shows up by comparing chart shapes across the
  sleeve slides, which is worth doing after any template swap.
- **Slide 8 dial dots** are also resized, not just recolored: the selected
  profile's dot grows to `DOT_SIZE_HIGHLIGHT`, every other dot shrinks to
  `DOT_SIZE_DEFAULT`, both resized around their own fixed center point
  (`_resize_dot_keep_center()` in `build_proposal.py`) so they stay
  anchored to the curve.
- **Slides 5 and 6's subtitle font size** was 14pt in the source deck
  while every other subtitle across the deck is 12pt — a genuine
  inconsistency in the reference deck itself, not something this skill's
  code introduced. Fixed by changing only the font size on the template's
  own frozen slide 5/6 subtitles (text and position untouched, honoring
  "never change the text or layout" for those two slides).
- **Sleeve slides' vehicle-breakdown donuts no longer repeat the sleeve
  name inside the donut hole** (the reference deck showed e.g. "Equities"
  centered inside the donut on the Equities sleeve slide, redundant with
  "Equities across investment vehicles" as the slide's own title one line
  above). Removed on Fixed Income, Equities, Alternatives and Commodities
  (Structured Products never had one). By contrast, the Portfolio
  Overview's two donuts ("Allocation by asset class", "Currency exposure")
  now carry their title *inside* the donut hole instead of as a separate
  textbox above it, Nunito Sans 14pt bold, navy — those two donuts sit
  side by side with no other label distinguishing them, so an in-hole
  label is needed there in a way it isn't on the single-donut sleeve
  slides.
- **The "Fixed Income Breakdown" (proposed bond selection) and "Equity
  breakdown" slides had a broken title.** Root cause: they're the only two
  slides using the `DEFAULT` slide layout, and that layout — unlike every
  other layout in the deck — never declared a title placeholder at all
  (confirmed by diffing `ppt/slideLayouts/slideLayout31.xml` against
  `slideLayout22.xml`, the `Title` layout every other data slide uses).
  Equity breakdown's title happened to still show text because its author
  worked around the missing placeholder with a plain hand-positioned
  textbox (`Text 0`) — at the wrong position. Fixed at the root: the
  `Title` layout's title-placeholder XML block was copied verbatim into
  the `DEFAULT` layout, so both slides now inherit the correct position
  (665163, 1032442 EMU) the normal way, through the same layout mechanism
  every other slide uses — not a per-slide workaround. (An earlier attempt
  that cloned a title *shape* directly onto each slide, without fixing the
  layout, silently failed: python-pptx read the clone's position back as
  `None` because a placeholder with no layout to inherit from resolves to
  nothing. If you're extending this area, verify a placeholder fix by
  reading the shape's `.left`/`.top` back after reopening the file, not
  just by confirming the edit ran without error.)
- **Leftover "agenda dots" section-nav cluster** on four appendix slides
  (original template positions 26 "Your Virtual team of Experts", 27 "Our
  Investment Universe", 29 "Syz Advisory Services", 30 "Our Value
  Proposition") — a row of 5 small ovals plus a numbered badge and label
  in the top-right corner, left over from a different presentation's
  section-navigation system and inconsistent with this deck's own design.
  Removed by bounding box (`left > 8000000 and top < 1800000`) rather than
  by name, since the shape names differ per slide.

- **Slide 8's profile-name captions** ("FIXED INCOME", "MODERATE", etc.)
  are now explicitly styled every build (color/size/bold), not left to
  the template's own formatting: only 4 of the 6 had any explicit run
  properties at all in the source deck (Moderate and Balanced had none,
  inheriting an undefined default) — see `PROFILE_LABEL_SHAPE` and the
  loop in `recolor_profile_dial()`. These live nested inside the same
  `Group 3` as the dots; searching `slide.shapes` instead of `group.shapes`
  for them is a real bug that shipped once already — silently no-ops
  instead of raising, so it only shows up as "nothing changed" on
  inspection, not as an error.
- **Portfolio Overview's two donut hole labels** sit at fixed coordinates
  (H=2.75cm/V=13.91cm and H=16.51cm/V=13.91cm) rather than a computed
  donut-center position — an earlier computed-center version came out
  visibly misaligned (the two donuts aren't quite the same size, so "center
  of the chart's bounding box" isn't "center of the visible ring").
- **The "Please enter client TAA" instruction text** was deleted from the
  house-view slide — the blank preference-matrix table stays for the
  advisor to fill in by hand, but the placeholder instruction sentence
  above it was clutter.
- **The "Headquartered in Geneva..." sentence on "Syz Group at a Glance"**
  was 11pt against every other body sentence in the deck being 12pt —
  fixed to 12pt, text and position untouched.

If you re-derive `assets/template.pptx` from a fresh copy of the reference
deck, redo all of these fixes — none of them survive re-deriving the
template from scratch.

## Market update extraction (slides 3-6)

`market_update.json` is **not** produced by a parsing script — extracting
an arbitrary week's PDF into "the four forces that moved markets" or "three
observations" is a summarization/judgment task that belongs to Claude
reading the PDF with the `pdf` skill (or its own multimodal `Read`), not a
deterministic transform. `SKILL.md` has the exact schema and extraction
instructions; `work/market_update.json` in this repo is a worked example
built from a real PDF (`work/20260817_Weekly_Investment_Meeting_Summary.pdf`)
for reference and for testing `build_proposal.py` without re-reading a PDF
every time.

Schema (see the worked example for real values):

```
{
  "week_of": str,
  "headline": {"summary_title": str, "intro_sentence": str,
               "stats": [{"value": str, "label": str}, ...4 of them]},
  "scoreboard": {"rows": [{"index": str, "week": str, "ytd": str}, ...
                           up to 9 — matches the template's 9 table rows]},
  "three_observations": [{"headline": str, "body": str}, ...3 of them],
  "four_drivers": {"intro_sentence": str,
                    "drivers": [{"number": int, "title": str, "body": str,
                                 "stat": str|null}, ...4 of them]},
  "scenario": {"intro_sentence": str,
               "probabilities": [{"name": str, "pct": number}, ...]},
  "house_view": {"intro_sentence": str,
                 "views": [{"asset_class": str, "stance": str, "body": str},
                           ...3 of them, matching the template's 3 columns]}
}
```

Picking exactly 4 headline stats / 9 scoreboard rows / 3 observations / 4
drivers / 3 house-view columns is a judgment call every week — pick the
figures and lines that best support the intro sentence, the same way a
human analyst would when condensing the source deck onto these slides. If
the source PDF doesn't have a clean equivalent for a field (e.g. no
scenario-probability table that week), leave the slide's original template
text in place rather than inventing numbers, and say so in your reply to
the user.
