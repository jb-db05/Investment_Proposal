# Investment Proposal Builder

An Anthropic Claude skill that builds a fully populated, Syz-branded
Investment Proposal PowerPoint deck from a client's portfolio Excel export,
a weekly market-update PDF, and a chosen risk profile.

The skill itself — documentation, scripts, template, reference material,
and a worked example — lives in [`investment-proposal-builder/`](investment-proposal-builder/).
Start there: [`investment-proposal-builder/SKILL.md`](investment-proposal-builder/SKILL.md).

## Layout

```
investment-proposal-builder/
├── SKILL.md            — entry point: workflow, usage, extension points
├── assets/             — the branded PowerPoint template
├── scripts/            — parser, chart updater, table builder, orchestrator
├── references/         — every calculation rule, slide-by-slide data map, risk-profile bands
└── work/                — worked example (sample Excel, PDF, generated deck)
```

## Quick start

```bash
cd investment-proposal-builder
python3 scripts/parse_portfolio.py work/Portfolio_Excel_Balanced.xlsx \
    --profile Balanced --client-name "Jane Doe" -o parsed.json

python3 scripts/build_proposal.py \
    --excel work/Portfolio_Excel_Balanced.xlsx --profile Balanced \
    --client-name "Jane Doe" --market-update work/market_update.json \
    -o "Investment Proposal - Jane Doe.pptx"
```

See `SKILL.md` for the full workflow, including how the market-update PDF
is turned into `market_update.json` (a judgment call made by Claude, not a
deterministic parser) and the validation/QA steps to run before delivering
a deck to a client.
