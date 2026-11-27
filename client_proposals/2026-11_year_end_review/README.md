# Year-End Portfolio Review 2026 (portfolio as of 26 Nov 2026)

Review only — no switches proposed. Built on the Syz template.

Inputs (not committed): updated client export (`Book_26.12.2026.xlsx`, statement dated 26 Nov 2026),
Syz Outlook 2027, Research Forum 23 Nov 2026, Syz Investment Conclusions Nov 2026, Earnings Pulse 23 Nov 2026.

- Commodities per client instruction: XAG/XPT accounts, BGF World Mining, VanEck Gold Miners, M&R Earths ETF, DCI XAG-USD notes.
- Estimated YTD return: buy-and-hold estimate on current holdings, start weight = weight / (1 + TW EUR return).
  Indicative only — excludes flows and positions sold during the year.
- Region / sector / theme tags: `scripts/analyse.py` (`META`).

```bash
python3 scripts/analyse.py Book_26.12.2026.xlsx && mv analysis.json analysis_new.json
python3 scripts/build_review.py analysis_new.json ../../investment-proposal-builder/assets/template.pptx out.pptx
```
