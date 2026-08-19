# Slide 8: risk-profile dial

Slide 8 ("Portfolio strategies across the (risk-return) trade-off") makes
two claims that must be consistent with each other and with the rest of
the deck:

1. A characteristics table (investment horizon, liquidity needs, etc.) for
   the risk profile the user selected as one of the three inputs.
2. A one-line statement: "Proposed profile for this portfolio: {profile} -
   {X}% in growth assets" — where **X is this portfolio's actual computed
   growth-asset percentage**, not a fixed model number for that profile.

## Why the growth-asset % is computed, not looked up

The six risk profiles are ordered points on a growth/defensive spectrum,
not fixed target weights — client portfolios built to the same profile
still vary. `parsed.json: risk_profile.growth_assets_pct` (assumptions.md
§12) is:

```
growth_assets_pct   = sum weight of {Equities, Alternatives, Commodities,
                                      Structured Products, Private Assets}
defensive_assets_pct = sum weight of {Cash, Fixed Income}
```

computed from the client's actual holdings. The band table below is used
only to sanity-check that the computed % is roughly where the selected
profile says it should be — it is never substituted for the real number on
the slide.

## Profile bands (sanity-check reference, not a target-weight table)

| Profile | Growth-asset band | Typical horizon | Typical liquidity need |
|---|---|---|---|
| Fixed Income | 0-10% | 1-3 years | High liquidity, capital preservation priority |
| Conservative | 10-25% | 3-5 years | Regular distribution, stable capital |
| Moderate | 25-45% | 5-10 years | Regular distribution but stable capital |
| Balanced | 45-65% | 5-10 years | Balanced between growth and distribution |
| Growth | 65-85% | 7-12 years | Growth priority, distributions opportunistic |
| Equity | 85-100% | 10+ years | Full growth priority, minimal liquidity needs |

These bands and the horizon/liquidity text come from
`RISK_PROFILE_CHARACTERISTICS` / `RISK_PROFILE_GROWTH_BAND` in
`scripts/build_proposal.py` — edit them there if your desk's bands differ;
this file documents the numbers in code, it doesn't independently define
them.

## What to do if the computed % falls outside its own profile's band

This is a real, useful signal — it means the client picked (or was
assigned) a risk profile that their actual proposed portfolio doesn't
match. `build_proposal.py` does not currently flag this automatically; if
you're extending it, the natural place is `fill_profile_dial()` in
`build_proposal.py`, printing a warning to the console (not fabricating a
"corrected" profile on the slide — that decision belongs to the advisor).

## Table fields that are NOT computed

Two rows in the slide 8 characteristics table are not derived from
`parsed.json` at all, because none of the three inputs (Excel, PDF, risk
profile) contains the data:

- **Financial profile** ("Professional") — a client-classification field
  from KYC, not from the portfolio.
- **Geographic exposure** ("Global") — the reference deck's own copy for
  this field describes the *offer* (Syz's own global reach), not the
  client's portfolio's geographic exposure (that's slide 12's
  `geographic_exposure_pct` chart, a different thing despite the similar
  label). Both are left as the template's static text.
