# Portfolio analysis and equity proposal: Port_10_Nov_26

As at 10 November 2026. House view from the Weekly Investment Meeting of 9 November 2026.
Portfolio total: USD 1,112,557 across 54 positions. Risk profile assumed: **Growth** (75–80% in growth assets).

## 1. Diagnosis

| | Today |
|---|---|
| Asset mix | Equities 79.8%, bonds 12.1%, other 6.8% (Bitcoin ETF, MyMap, global dividend fund), cash 1.3% |
| Equity lines | 37: 24 single stocks, 11 funds/ETFs, 2 certificates |
| Lines below 2% of the portfolio | 20 of 37 |
| Single stocks | 49.8% of the portfolio, half of it in names under 2% |
| AI-exposed equities | 28.5% of the portfolio, **36% of the equity sleeve**: Amazon, Alphabet, Microsoft, Meta, Broadcom, Nvidia, Tesla, ams-OSRAM, the cyber-security ETF, Alibaba, the uranium ETF |
| China | 11.7% of the portfolio across six lines (Alibaba, BYD, NIO, Prosus, MSCI China ETF, SG China certificate), plus China inside the EM ETF |
| Gaps | Healthcare 1.1% (two small lines), Switzerland 0%, India 0%, broad US exposure only through mega-caps |
| Speculative tail | Satellogic, Planet Labs, Circle, HelloFresh, GRAIL, uranium and rare-earth ETFs: 7.7% in total, most of it down 30–90% |

Why this matters now: the 5 November sell-off was concentrated in AI (Nasdaq 100 −4.9% on the week, semis −12.3% in one day), while Europe, India, pharma and staples held up. The house view stays neutral on equities and every region, overweight cash, gold and commodities, and Syz Symphony cut its equity signal to 75%. The current book is a set of overlapping thematic bets rather than a diversified equity allocation.

## 2. Proposal (equities)

**Sell 21 lines and trim Amazon** (USD 374k, 33.6% of the portfolio):

- **China, six lines into one ETF**: sell NIO, BYD, Alibaba, Prosus and the SG China Next Winners certificate; top up the iShares MSCI China ETF to 3%.
- **AI and tech concentration**: sell Tesla, Nvidia, ams-OSRAM and the L&G Cyber Security ETF (take the +48% gain); trim Amazon from 6.0% to 4.0%. Keep Alphabet, Microsoft, Meta, Broadcom and Amazon as the core AI holdings, about 16.6% of the portfolio.
- **Speculative and thematic satellites**: sell Satellogic, Planet Labs, Circle, the Global X Uranium ETF and the VanEck Rare Earth ETF. SpaceX stays as the single space holding.
- **Small lines folded into funds**: sell HelloFresh, GRAIL, Illumina, DSV, Brookfield, YPF and the second Japan vehicle (TOPIX ETF).

**Buy 6 new funds and top up 5 existing ones** (USD 316k):

| Position | Target | Why |
|---|---|---|
| iShares Nasdaq US Biotech UCITS ETF (IE00BYXG2H39) | 3.5% | Biotech: AI-insulated growth, replacing the Illumina/GRAIL single lines |
| Xtrackers MSCI World Health Care UCITS ETF (IE00BM67HK77) | 3.0% | Pharma and medtech: the defensive half of the healthcare allocation |
| Xtrackers S&P 500 Equal Weight UCITS ETF (IE00BLNMYC90) | 5.0% | US breadth: ex-Mag-7 EPS +27% vs Mag-7 +14% in Q3 |
| iShares Core MSCI Europe UCITS ETF (IE00B4K48X80) | 4.0% | Low AI exposure, Q3 EPS +28% (consider a USD-hedged class) |
| iShares Core SPI ETF (CH0237935652) | 2.5% | Swiss pharma, staples and insurers; SMI +15% YTD |
| iShares MSCI India UCITS ETF (IE00BZCQB185) | 3.0% | The "anti-AI trade": Nifty +8.5% in the week Topix fell 6.5% |
| Top-ups | | MSCI World Value to 4.5%, Vanguard EM to 4.0%, MSCI China to 3.0%, iMGP Euro Select to 3.0%, Energy Select to 2.0% |
| Cash | 1.3% → 6.6% | House overweight on cash; redeploy in stages around Nvidia results (25 Nov) and 2027 capex guidance (late Jan) |

## 3. Result

| Equity sleeve | Current | Proposed |
|---|---|---|
| Equity weight | 79.8% | 74.5% |
| Lines | 37 | 22 |
| Single stocks | 24 | 8 |
| Lines below 2% | 20 | 1 (SpaceX, 1.9%) |
| AI-exposed share of equities | 35.8% | 22.2% |
| AI-insulated (% of portfolio) | 10.4% | 20.3% |
| Health care (% of portfolio) | 1.1% | 6.5% |
| China (% of portfolio) | 11.7% | 3.0% |
| Europe + Switzerland (% of portfolio) | 11.5% | 15.8% |

Unrealised P/L crystallised on sales: −USD 101.3k (gains USD 22.2k, losses −USD 123.5k), based on the custodian's cost basis. Most of the sales harvest losses, which may offset other gains, depending on the client's tax residence.

## 4. Points to confirm before presenting

- **Risk profile**: Growth was assumed from the 75–80% growth-asset weight. Change `--profile` in `build.sh` if the KYC profile differs.
- **Client and RM names**: the cover reads "Prepared for Client" and "Relationship Manager: [Name]".
- **Fixed income is unchanged.** The house view is underweight fixed income and overweight gold. The portfolio holds 0% gold and about 2.6% in Argentine sovereign bonds (four lines, two of them under 0.2%), which are worth a separate review.
- **Share classes and listing**: the new funds are UCITS, so they avoid US estate-tax exposure for a non-US client. Check availability on the platform, and consider a USD-hedged class for the EUR and CHF funds.
- **Classifications** (sector, region, AI camp) are advisor tags in `rebalance_spec.json`. Diversified funds are not looked through.
- **Slides 5–6** (economic scenario, investment views) are frozen by design and still show the August scenario weights (50/30/20). The 9 November meeting has 40/40/20.

## Files

- `Investment Proposal - 10 Nov 2026.pptx`: the client deck, 33 slides. The new slides 9–13 show the diagnosis, current vs proposed, sales and purchases.
- `Proposed_Portfolio_10_Nov_26.xlsx`: the proposed portfolio in the custodian layout, plus `Trades` and `Before vs After` sheets.
- `rebalance_spec.json`: every decision and tag (edit this file and rerun `./build.sh`).
- `market_update.json`: market slides 3–4, taken from the 9 Nov PDF.
