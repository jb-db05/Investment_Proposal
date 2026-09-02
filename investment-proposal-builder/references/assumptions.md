# Calculation assumptions

This document is the single source of truth for every number `scripts/parse_portfolio.py`
computes. If a slide shows a figure that looks wrong, check the rule here first —
the code is written to match this document exactly, not the other way round.
No number in the output deck is invented: everything traces back to a column
in the client's Excel export, or to a documented rule applied to that column.

## 1. Input file shape

`parse_portfolio.py` expects the client's custodian export with two tabs:

- **`Portfolio`** — the client's current holdings. Section-header rows in
  column A (`Cash`, `Short-term instruments`, `Bonds`, `Equities`,
  `Structured products`, `Private Assets`, `Gold and other commodities`,
  `Other investments`) mark the top-level asset class of every row beneath
  them, until the next header row. This is the **only** place asset-class
  membership comes from — the script never re-derives it from instrument
  names — with one documented exception, below.
- **`Fixed Income`** — a separate, already-curated bond selection to be
  *proposed* to the client (not part of their current holdings). Every row
  is treated as equally weighted, matching the reference deck's own framing
  ("N issues, equally weighted at X each"). **Optional**: an export without
  this tab is a portfolio review with no bond proposal attached, and the
  "Fixed Income Breakdown" slide is dropped from the deck rather than left
  showing the template's own example selection
  (`build_proposal.drop_proposed_bond_selection`).

Only the holdings tab is required, and only its *shape* is: if the workbook
has no tab literally named `Portfolio` (a custodian export re-saved out of
Excel routinely arrives as `Sheet1`), the one other sheet carrying the
export's header row — a row of >=5 filled cells including an `ISIN code`
column — is used instead, and the script says which one it picked. Two such
sheets is an error, not a guess.

Column names are matched case-insensitively with whitespace stripped
(the export has trailing spaces on some headers, e.g. `"Rating  "`).

### 1a. Precious-metal accounts are commodities, not cash

The one exception to "the section header decides". The custodian files a
metal account (`CURRENT ACCOUNT IN XAG`, `... IN XPT`) under `Cash`,
because it is an account. Economically it is not cash: the balance is
ounces of silver or platinum, its value moves with the metal price, and it
carries no deposit-rate return. Counting it as cash overstates the client's
defensive assets and understates their commodity exposure — the two numbers
the risk profile is actually read from.

So a row is reclassified to `Commodities` when **both** hold
(`is_precious_metal_account()`):

- its `Currency` is an ISO 4217 precious-metal code — `XAU` gold, `XAG`
  silver, `XPT` platinum, `XPD` palladium; and
- its `Description` contains "ACCOUNT".

The currency code is what carries the meaning; the description test keeps
the rule to account balances, so a metal-denominated *security* still goes
wherever its own section header puts it. Nothing else about the row changes
— its weight, valuation and geography come from the same columns as before.

Being a commodity by asset class does not make a metal account *physical*:
it is the currency form of the metal, so §10 groups it with cash rather
than with the sleeve's bars and coins.

## 2. Row filtering

- A holding row is included in the line-items table and every weight-based
  chart only if `Weight (%) > 0`.
- Rows with `Valuation + accr. interest (in QC) == 0` (undrawn private-equity
  commitments, e.g. "... COMMIT") are excluded from weighted breakdowns and
  the line-items table, but their count and total committed amount are
  reported separately as `uncalled_commitments` in `parsed.json` so nothing
  is silently dropped — they just aren't a % of the invested book.
- All weights are read directly from the file's own `Weight (%)` column
  (already computed by the custodian against the portfolio total) rather
  than recomputed from valuations, so the numbers tie out to the client's
  statement exactly.

## 3. Currency of record

- Portfolio total value and all EUR-denominated figures use the
  `Valuation + accr. interest (EUR)` column.
- `base_currency` (used to label every aggregate monetary figure — total
  value, KPIs, sleeve values, rate-impact) is **fixed at EUR**, matching
  this column and the file's own stated total (e.g. "Total: 21'300'578
  EUR"). This is deliberately NOT the same as `dominant_holding_currency`
  (the currency with the largest weighted valuation, used only for
  informational purposes) — conflating the two would label EUR-denominated
  aggregates with whatever currency happens to dominate the book, which is
  wrong regardless of which currency wins. A single holding's own value is
  shown in that holding's own `Currency` column via `value_qc` — see §9.

## 4. Fixed income sub-buckets (Cash / Fixed Income line-items table)

The raw export has no "Govies vs Corporate vs High Yield vs EM Debt" column,
so this is a rule-based classifier applied to `Composite Rating`, `Duration`,
and `Geographical breakdown`. It mirrors the bucket names in the Syz TAA
preference matrix (see `slide_recipe.md` §Slide 5-6) so the line-items table
groups holdings the same way the house view talks about them:

1. **EM Debt** — `Geographical breakdown` maps to region `MEA EM` (see
   `geography_map.py`) AND the row sits in the `Bonds` section. EM
   geography takes priority over rating.
2. **Govies 1-10 (local)** / **Govies 10+ (local)** — `Composite Rating` in
   `{AAA, AA+, AA, AA-}` (investment-grade sovereign/quasi-sovereign band)
   AND `Sector` is blank or `Government`. Split by `Duration`: `<= 10` is
   "1-10", `> 10` is "10+".
3. **Corporate IG (local)** — `Composite Rating` in
   `{A+, A, A-, BBB+, BBB, BBB-}`.
4. **High Yield (local or global hdg)** — `Composite Rating` in
   `{BB+, BB, BB-, B+, B, B-, ...}` or `Not Rated` (private/structured
   credit defaults to High Yield rather than Govies/IG, being the more
   conservative assumption for an unrated instrument).
5. Any `Bonds`-section row that is itself a fund/ETF (`Description`
   contains "FUND", "FD", "ETF", "SICAV", "CERTIFICATE") is still bucketed
   by its own rating/geography above — funds are not a separate bucket in
   the line-items table, only in the "across investment vehicles" donut
   (§6).

Note: the `Portfolio` sheet carries **no `Sector` column** (only the
separate `Fixed Income` proposal tab does), so step 2's IG/Govies split
uses `Composite Rating` + `Duration` only — there is no "and Sector is
blank or Government" clause in the actual code, only rating and duration.

Line-items table sub-grouping for the `Equities` section (slide 10) is
separate from this fixed-income logic: each equity row is bucketed into
`US / Eurozone / UK / Switzerland / Japan / EM / Global / thematic` via
`geography_map.py::equity_region_of()`, applied to `Geographical
breakdown`.

Desks with an authoritative issuer-level mapping should not rely on this
heuristic: pass `--bucket-overrides overrides.csv` (columns `isin,bucket`)
to `parse_portfolio.py` and any ISIN present there skips the rule above.

## 5. Asset-class donut (Portfolio Overview, slide "Allocation by asset class")

One slice per `Portfolio` sheet section header (after the §1a
precious-metal reclassification), weight = sum of that section's
`Weight (%)`. Sections with 0% total are omitted from the chart (not shown
as a 0% slice).

## 6. Investment-vehicle donut (per-sleeve "across investment vehicles")

Within a sleeve, each holding is classified as:

- **Direct line** — `Description` does not contain any fund/structured
  keyword below.
- **Fund** — `Description` contains "FUND", "FD", "SICAV", "ETF", "ETC",
  or the row's `Composite Rating` is blank and `Coupon (%)` is blank (a
  common signature of collective vehicles in this export).
- **Structured / AMC** — `Description` contains "CERTIFICATE", "CERT",
  "AMC", "TRACKER", or the ISIN prefix is `XS`/`CH` combined with a
  `Coupon (%)` that is blank and a bullet/barrier-style name pattern
  (e.g. "DCI-", "RAF ", "CPN ").

This is applied in the order above (first match wins) and is a heuristic
for the same reason as §4 — the export has no vehicle-type column.

**The blank-rating/blank-coupon signature is a signature of _securities_**
— collective vehicles that carry neither a rating nor a coupon. Three kinds
of holding carry neither for an entirely different reason, and are
classified **Direct line** without consulting it:

- a bank account (the `Cash` section);
- a §1a precious-metal account — the metal itself, held in account form;
- anything left in the `Commodities` sleeve after the keyword rules above,
  which is physical metal: bars and coins (`GOLD KG`, `OR KRUGERRAND`).
  Funds and ETCs in that sleeve carry "FUND"/"ETF"/"ETC" in their
  description and are caught as **Fund** before this rule is reached.

## 7. Geographic & sector exposure (portfolio-wide bar charts)

- Geography: each holding's `Geographical breakdown` value is mapped to one
  of five regions via `geography_map.py::region_of()`, then weighted by
  `Weight (%)`.
- Sector: **the `Portfolio` sheet has no `Sector` column** — this is a real
  limitation of the custodian export, not an oversight (the reference deck
  discloses the same gap in its own footnote: "these fields are not in the
  client file"). `sector_exposure_pct` is therefore `null` in `parsed.json`
  unless the caller supplies `--sector-overrides overrides.csv` (columns
  `isin,sector`, a desk-maintained issuer→GICS-style mapping). When `null`,
  `build_proposal.py` renders the sector panel with a "sector data not
  available for this custodian file" note instead of a chart — it never
  fabricates a single fake "Diversified 100%" slice.

## 8. Equity breakdown (geography / sector / market cap)

- Geography uses the finer six-bucket `equity_region_of()` map (`US /
  Eurozone / UK / Switzerland / Japan / EM / Global / thematic`), scoped to
  the `Equities` section only, weights renormalized to sum to 100% of the
  equity sleeve (not the whole portfolio). This is the same taxonomy used
  by the slide 10 line-items table, so the two slides agree with each
  other.
- Sector: same `--sector-overrides` mechanism and `null`-when-absent
  behavior as §7, renormalized to the equity sleeve.
- Market cap is not a column in the export — every equity line is
  classified `Large cap` unless the desk supplies
  `--market-cap-overrides overrides.csv` (columns `isin,market_cap`);
  without overrides the market-cap donut is `null` in `parsed.json` and the
  build script leaves the market-cap panel on the slide with a "data not
  available" placeholder rather than fabricating a split.

## 9. Concentration & top holdings

- Sort all included holdings by `Weight (%)` descending.
- `top_5` / `top_10` / `top_20` = cumulative weight of the first 5/10/20
  rows.
- `positions_above_5pct` = count of rows with `Weight (%) > 5`.

## 10. Liquidity profile

Each holding is classified by **how it is actually realised** — not by its
asset class, and not by its §6 vehicle alone. Five buckets, first match
wins (`classify_liquidity()`):

- **Illiquid (lock-up)** — `Private Assets` section, or `Description`
  contains "COMMIT" (already excluded per §2 if zero-valued, but a partly
  drawn commitment line still counts here).
- **Cash & metal accounts** — the `Cash` section, plus §1a precious-metal
  accounts. An account balance settles like money whether it is denominated
  in a currency or in ounces: a metal account is the *currency* form of the
  metal, not a bar in a vault, so it belongs here rather than under
  Physical. It is still a commodity in the §5 asset-class donut — the two
  views answer different questions.
- **Physical assets** — holdings in the `Commodities` sleeve
  classified "Direct line" under §6: physical bars and coins. Sellable at a
  screen price, but neither a listed security nor a fund with a daily NAV,
  and settlement is a delivery — so neither of those two buckets fits.
- **Daily-liquid fund** — classified "Fund" under the §6 vehicle rule.
- **Listed** — everything else (direct bonds, direct equities, structured
  products, ETFs/ETCs).

`liquid_share_pct` is everything that is not **Illiquid (lock-up)** — the
four other buckets are all realisable at short notice, by different routes.

## 11. Income & interest-rate sensitivity

- Running yield (portfolio-level KPI) = `sum(Weight% * Yield%)` over rows
  where `Yield (%)` is populated; rows without a yield figure (mostly
  equities) contribute 0 to this sum and are excluded from the weighting
  base printed in the caption, matching the reference deck's caveat text
  pattern ("N of M fixed-income line(s) report no modified duration...").
- Income by sleeve (EUR) = `sum(Valuation EUR * Yield% or Coupon%)` per
  asset-class section, using `Yield (%)` where present else `Coupon (%)`.
- FI duration (portfolio KPI) = valuation-weighted average `Duration`
  across `Bonds`-section rows that report a duration.
- Interest-rate sensitivity ("Impact +100bp") = `-1 * FI duration * 1% *
  total Fixed Income valuation`.
- Duration sub-sleeve table groups `Bonds`-section rows by a short
  descriptive label derived from `Sector` + instrument type keywords
  (mirrors the reference deck's "Corporate bond (auto)", "Corporate bond
  (oil)" style labels) — see `parse_portfolio.py::describe_bond_subsleeve()`
  for the exact keyword list; this label is cosmetic only and does not
  feed any other calculation.

## 12. Risk-profile dial (slide 8)

See `slide8_profile_dial.md` for the full profile-band table. In short:
"growth assets" = Equities + Alternatives + Commodities + Structured
Products + Private Assets; "defensive assets" = Cash + Fixed Income. The
slide shows the client's **actual** computed growth-asset % positioned
against the chosen risk profile's band — the % is never a fixed model
number, it is this portfolio's real figure.

## 13. Rounding

Percentages are computed at full float precision internally, then rounded
to 6 decimal places (`parse_portfolio.py::pct()`) before landing in
`parsed.json` — 6dp, not fewer, specifically so a genuinely nonzero but
very small weight (a residual currency balance worth ~0.00003% of the
book, say) survives as a small positive number rather than collapsing to
exactly `0.0` before the display layer ever sees it. Display-time
formatting then rounds again to 1 decimal place. Donut/bar chart category
labels bake the rounded percentage into the label text (matching the
reference deck's `"Fixed Income  55%"` style), so a chart rebuilt from
`parsed.json` must round the same way `update_chart.py::pct_label_*()`
does, or labels and slice sizes will disagree by a fraction of a point.

**`<0.1%` display rule.** Anywhere a per-holding or per-category
percentage is shown at 1-decimal precision (`fmt_pct()` in
`build_proposal.py` and `build_line_items_tables.py`,
`pct_label_one_decimal()` in `update_chart.py`), a value that is
genuinely greater than zero but would round to `0.0%` displays as `<0.1%`
instead — `"INCOME TO BE RECEIVED IN GBP  0.0%"` reads as an error or
missing data next to a real holding, not as a very small real position.
An actual zero (e.g. a sleeve with no income at all) still shows `0.0%`
as-is; only `0 < value < 0.1` gets the `<0.1%` treatment. This rule only
works because of the 6dp rounding above — if `pct()` rounded to 1dp
before storing, the distinction between "genuinely zero" and "just very
small" would already be gone by the time the display layer runs.
