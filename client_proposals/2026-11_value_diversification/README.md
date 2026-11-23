# Portfolio Review & Diversification Proposal — November 2026

Client-specific deck built on the Syz template (`investment-proposal-builder/assets/template.pptx`).

## Inputs (not committed — confidential / large)
- Client portfolio export (`Book2.xlsx`, single sheet, custodian layout)
- `Bank_Syz_Outlook_2027_full_v8` (Outlook 2027, CIO office)
- `20261123_NOV_ResearchForumDeck-Final` (Research Forum, 23 Nov 2026)

## Client instructions applied
- Treated as **commodities**: XAG and XPT accounts, BGF World Mining, VanEck Gold Miners ETF,
  WisdomTree M&R Earths ETF, DCI XAG-USD note.
- Value approach kept; no new AI exposure; gold kept, other metals sold; energy **not** used as a diversifier
  (client works in the energy sector).

## Method notes
- YTD performance = custodian `TW Perf. (EUR, %)` (time-weighted, EUR; from purchase date for 2026 purchases).
- Region / sector / theme tags are assigned at issuer or fund-mandate level in `scripts/analyse.py` (`META`),
  because the custodian file has no sector field.
- Proposal weights and trades are defined in `scripts/build_deck.py` (`SELL`, `BUY`).

## Rebuild
```bash
python3 scripts/analyse.py Book2.xlsx            # writes analysis.json
python3 scripts/build_deck.py analysis.json ../../investment-proposal-builder/assets/template.pptx out.pptx
```
