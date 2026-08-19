---
name: investment-proposal-builder
description: "Build a fully populated Syz-branded Investment Proposal PowerPoint deck from a client's portfolio Excel export, a weekly market-update PDF, and a chosen risk profile (Fixed Income, Conservative, Moderate, Balanced, Growth, or Equity). Use this skill whenever the user asks to create, generate, or build an investment proposal, investment pitch deck, or portfolio review presentation from a client portfolio spreadsheet and a market commentary PDF. Produces real PowerPoint tables and native charts (not images), with every number traced back to the source Excel — no fabricated data."
---

# Investment Proposal Builder

Turns three client-specific inputs into a client-ready PowerPoint deck:

1. **Portfolio Excel** — a custodian export of the client's current holdings
   (a `Portfolio` sheet with section-header rows per asset class) plus,
   optionally, a `Fixed Income` sheet holding a curated bond *proposal*.
2. **Market-update PDF** — a weekly market commentary (Syz Research style).
3. **Risk profile** — one of `Fixed Income`, `Conservative`, `Moderate`,
   `Balanced`, `Growth`, `Equity`.

Output: a ~36-39 slide `.pptx` (slide count varies with how many holdings
the client has — see "Pagination" below), with every number, chart and
table traced back to the Excel file, and every piece of market commentary
traced back to the PDF. **Nothing is invented.** Where a number genuinely
isn't available in the source file (e.g. equity sector classification —
see `references/assumptions.md` §7), the deck says so instead of guessing.

## Before you start

Skim `references/slide_recipe.md` once per session — it maps every slide
to its data source and the function that fills it, and documents the scope
boundary (which slides are static bank collateral, never data-driven).

## Workflow

### 1. Gather inputs

Confirm you have all three: the portfolio Excel, the market PDF, and the
risk profile (ask if any is missing — don't guess a risk profile). Also
get the client's name for the cover slide, and optionally a valuation date
(defaults to today).

### 2. Parse the portfolio

```bash
python scripts/parse_portfolio.py "<portfolio.xlsx>" \
    --profile "<Fixed Income|Conservative|Moderate|Balanced|Growth|Equity>" \
    --client-name "<Client Name>" \
    --valuation-date "<e.g. 6 August 2026>" \
    -o parsed.json
```

Read the printed summary line (position count, total value, dominant
holding currency) and sanity-check it against the Excel — e.g. does the
position count match the file's own "Segment Total (N)" label if it has
one. If your desk has an authoritative fixed-income bucket mapping,
equity-sector mapping, or market-cap classification, pass
`--bucket-overrides`, `--sector-overrides`, `--market-cap-overrides` (CSV
files — see `references/assumptions.md` §4, §7-8 for the exact column
names). Without them, sector and market-cap breakdowns are correctly
reported as unavailable rather than fabricated.

### 3. Extract the market update (you do this — it's not a script)

Read the market PDF (the `pdf` skill or your own multimodal `Read` tool)
and produce `market_update.json` matching the schema in
`references/slide_recipe.md` (§"Market update extraction"). This step is
judgment, not a deterministic transform: you're picking which 4 headline
stats, which 9 scoreboard rows, which 3 observations, which 4 drivers, and
which 3 house-view columns best represent that week's commentary — the
same summarization call a research analyst makes when condensing their own
full deck onto these four slides. `work/market_update.json` in this repo
is a complete worked example from a real PDF; use it as a template for
field names and tone, not as a source of numbers for a different week.

If the PDF's structure doesn't cleanly map to a field (no explicit
scenario-probability table that week, say), leave that field's slide
un-filled (the template's original text stays) and tell the user instead
of inventing numbers.

### 4. Build the deck

```bash
python scripts/build_proposal.py \
    --excel "<portfolio.xlsx>" \
    --profile "<risk profile>" \
    --client-name "<Client Name>" \
    --market-update market_update.json \
    --valuation-date "<same date as step 2>" \
    -o "Investment Proposal - <Client Name>.pptx"
```

This re-runs `parse_portfolio.py` internally (pass the same override flags
here too if you used them in step 2) and writes the final deck. It also
rebuilds the holdings-list slides as real tables (`build_line_items_tables.py`)
— see "Pagination" below.

### 5. Validate

```bash
python /mnt/skills/public/pptx/scripts/office/validate.py "Investment Proposal - <Client Name>.pptx" --original assets/template.pptx
```

Expect exactly the 3 pre-existing `ppt/comments/*` ID-uniqueness findings
(inherited from the source reference deck, harmless, not introduced by
this skill — see `references/assumptions.md`'s closing note if you want
the history). Any other new error is a real regression — fix it before
delivering.

### 6. Check for leftover placeholder content

Grep the built deck's extracted text for tokens that should never survive
a real build: the reference client's own numbers (`7.0m`, `36\nPositions`,
`86.6%`), the literal string `[Client Name`, and `Please enter client TAA`
only being present on the one slide that's meant to stay a manual
worksheet (the TAA preference-matrix slide — that one is intentionally
static, see `slide_recipe.md`).

```bash
python -c "
from pptx import Presentation
prs = Presentation('Investment Proposal - <Client Name>.pptx')
text = '\n'.join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
bad = ['[Client Name', '7.0m', 'USD 7.0M', '86.6%']
print([b for b in bad if b in text])
"
```

An empty list is a pass.

### 7. Visual QA

Convert to PDF/images and look at every slide (`pptx` skill's
`scripts/office/soffice.py`, or `thumbnail.py`). **Known limitation**:
LibreOffice conversion did not work at all in the sandbox this skill was
built in (`soffice` failed to load even a trivial pptx — an environment
issue, not specific to this deck; confirmed by testing a blank file). If
that's still the case in your environment, say so explicitly rather than
claiming a visual check that didn't happen, and ask the user to open the
file themselves for final sign-off. If `soffice` does work for you, check
specifically: the holdings-list table(s) don't overflow the slide, no
category got split across two slides, donut/bar chart labels aren't
truncated, and the sleeve "largest holdings" tables aren't taller than the
slide.

## Pagination (the hardest part)

The client's full holdings list never fits on the reference deck's
original 2 slides once you have more than ~35-40 real positions.
`scripts/build_line_items_tables.py` rebuilds those slides as real
PowerPoint tables (`a:tbl`, not textboxes) and paginates dynamically: a
*group* (fixed-income bucket, equity region, or asset class — see
`assumptions.md` §4) never splits across two slides. If a page doesn't
have room for the next whole group, that group moves to the next slide,
and a new slide is cloned from the template's blank line-items layout.
This is why it must run **last** in `build_proposal.py` — see the comment
in `build()` and `references/slide_recipe.md`'s note on slide numbering.

## Directory structure

```
investment-proposal-builder/
├── SKILL.md                       — this file
├── assets/
│   └── template.pptx              — branded template, slides 9-10 cleared to blank canvas
├── scripts/
│   ├── parse_portfolio.py         — Excel -> parsed.json (all calculations, no chart/slide code)
│   ├── update_chart.py            — python-pptx chart.replace_data() helpers
│   ├── build_line_items_tables.py — rebuilds the holdings-list slides as real tables
│   └── build_proposal.py          — orchestrator: runs all of the above, fills every other slide
├── references/
│   ├── assumptions.md             — every calculation/classification rule, by section number
│   ├── slide_recipe.md            — slide-by-slide data source map + market_update.json schema
│   ├── slide8_profile_dial.md     — risk-profile bands and the growth-asset % logic
│   └── geography_map.py           — country -> region lookups (portfolio-wide and equity-only)
└── work/                          — worked example: a real portfolio Excel, PDF, market_update.json
                                      and parsed.json, for testing changes to the scripts
```

## Extending this skill

- **New risk-profile bands or characteristics text**: edit
  `RISK_PROFILE_CHARACTERISTICS` / `RISK_PROFILE_GROWTH_BAND` in
  `build_proposal.py`, and update `slide8_profile_dial.md` to match — the
  doc documents the code, so keep them in sync.
- **A different fixed-income or equity taxonomy**: edit the classifier
  functions in `parse_portfolio.py` (`classify_fi_bucket`,
  `geography_map.py`), and update `assumptions.md` §4/§7-8 to match.
- **A custodian export with different column names**: `parse_portfolio.py`
  matches headers case-insensitively with whitespace stripped
  (`norm_header()`) — add the new header's normalized form wherever the
  old one is referenced (search for `.get("...")` calls).
- **Never** hand-edit `parsed.json` or a built `.pptx` for a real client
  delivery — fix the generating script and rebuild, so the same fix
  applies to the next client too.
