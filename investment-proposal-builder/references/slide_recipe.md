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
| 5 | Our economic scenario | `market_update.json: scenario` | `build_proposal.fill_market_slides` |
| 6 | Our current investment views | `market_update.json: house_view` | `build_proposal.fill_market_slides` |
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

Slides 24 and 26-36 (appendix, "Syz Group at a Glance", advisory service
tiers, individual fund recommendation one-pagers) are **not** data-driven
by this skill. They are bank collateral / marketing content that doesn't
come from the client's portfolio or the market PDF, and none of the three
inputs (Excel, PDF, risk profile) contains what would be needed to
personalize them (e.g. which specific funds to recommend is an advisor
judgment call, not a data transform). They ship as-is from the template.
If a future version needs to vary the recommended-fund one-pagers by risk
profile, that's a new input (a fund-recommendation-by-profile mapping),
not something to reverse-engineer from the three existing inputs.

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

## Table column simplification (sleeve "largest holdings" tables)

The reference deck used a different table column set per sleeve (Fixed
Income: Holding/Value/Share/Duration/Yield/Price/Bid-ask; Equities:
Holding/Value/Share/Country/Sector/Market Cap — with Country/Sector/Market
Cap always blank, since that data isn't in the client file). This skill
uses one uniform 4-column layout — **Holding / Value / Share / Yield or
Coupon** — across every sleeve slide (13/15/17/18/19), for two reasons:
the template's own tables were sized for exactly 1-5 holdings and can't
grow, so `build_proposal.rebuild_holdings_table()` replaces them with a
freshly-sized table on every build; and a uniform layout means the same
code and the same "don't fabricate a blank column" rule apply everywhere.

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
