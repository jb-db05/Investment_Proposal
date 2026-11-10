#!/usr/bin/env bash
# Rebuild every deliverable in this folder from its inputs:
#   Port_10_Nov_26.xlsx (current holdings), rebalance_spec.json (advisor decisions),
#   market_update.json (from the 9 Nov Weekly Investment Meeting PDF).
set -euo pipefail
cd "$(dirname "$0")"
B=../../investment-proposal-builder/scripts
python3 "$B/build_rebalance.py" Port_10_Nov_26.xlsx rebalance_spec.json \
    --out-excel Proposed_Portfolio_10_Nov_26.xlsx --out-json rebalance.json \
    --out-sector-csv sector_overrides.csv --out-mcap-csv market_cap_overrides.csv
python3 "$B/build_proposal.py" --excel Proposed_Portfolio_10_Nov_26.xlsx --profile Growth \
    --client-name "Client" --market-update market_update.json --valuation-date "10 November 2026" \
    --sector-overrides sector_overrides.csv --market-cap-overrides market_cap_overrides.csv \
    --rebalance rebalance.json --keep-parsed-json parsed.json \
    -o "Investment Proposal - 10 Nov 2026.pptx"
