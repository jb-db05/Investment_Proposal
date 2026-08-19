"""
Country -> region lookup for the "Geographic exposure" chart on the Portfolio
Overview slide.

The only geographic signal in a standard custodian export is the
"Geographical breakdown" column (issuer domicile / fund registration
country), not verified economic look-through exposure. That is a known
limitation, not a bug: mapping domicile to region is the best available
approximation without fabricating data the file doesn't contain, and it is
what the reference deck (work16.pptx, slide 12) does too.

Five regions match the reference deck's own taxonomy:
    US        - United States / North America
    Europe DM - developed Western Europe
    Japan     - Japan
    MEA EM    - Middle East, Africa, Latin America, emerging Asia, emerging
                Europe
    Global    - multi-country funds/ETFs, fund domiciles that carry no
                economic signal (Luxembourg, Ireland, Cayman Islands as
                registration jurisdictions), stateless/precious-metal/
                index lines, or anything not in the map below

Edit this file's REGION_MAP to match your desk's own taxonomy — nothing
else in the skill depends on these exact five names except the slide 12
chart label built from region_of().
"""

from __future__ import annotations

REGION_MAP: dict[str, str] = {
    # --- US / North America ---
    "United States": "US",
    "United States of America": "US",
    "Canada": "US",
    # --- Japan ---
    "Japan": "Japan",
    # --- Europe DM ---
    "United Kingdom": "Europe DM",
    "France": "Europe DM",
    "Germany": "Europe DM",
    "Switzerland": "Europe DM",
    "Netherlands": "Europe DM",
    "Italy": "Europe DM",
    "Sweden": "Europe DM",
    "Spain": "Europe DM",
    "Belgium": "Europe DM",
    "Austria": "Europe DM",
    "Denmark": "Europe DM",
    "Norway": "Europe DM",
    "Finland": "Europe DM",
    "Portugal": "Europe DM",
    "Ireland": "Europe DM",
    "Luxembourg": "Europe DM",
    "European Union": "Europe DM",
    "Misc.Europe": "Europe DM",
    "Guernsey": "Europe DM",
    "Jersey": "Europe DM",
    "Isle of Man": "Europe DM",
    # --- MEA EM / broader emerging markets ---
    "Poland": "MEA EM",
    "Czech Republic": "MEA EM",
    "Hungary": "MEA EM",
    "Turkey": "MEA EM",
    "South Africa": "MEA EM",
    "Brazil": "MEA EM",
    "Mexico": "MEA EM",
    "China": "MEA EM",
    "India": "MEA EM",
    "Indonesia": "MEA EM",
    "Saudi Arabia": "MEA EM",
    "United Arab Emirates": "MEA EM",
    "Qatar": "MEA EM",
    # --- Global / no usable economic signal from domicile alone ---
    "Miscellaneous": "Global",
    "Stateless": "Global",
    "Several countries": "Global",
    "Various countries": "Global",
    "Indices": "Global",
    "Currency and precious metal": "Global",
    "Cayman Islands": "Global",
}

DEFAULT_REGION = "Global"


def region_of(country: str | None) -> str:
    """Map a raw 'Geographical breakdown' value to one of the five report regions."""
    if not country:
        return DEFAULT_REGION
    return REGION_MAP.get(country.strip(), DEFAULT_REGION)


# Finer six-bucket taxonomy used only for the Equities sleeve: the line-items
# table (slide 10) and the equity breakdown "by geography" donut (slide 16)
# both split equities more finely than the five portfolio-wide regions above.
EQUITY_REGION_MAP: dict[str, str] = {
    "United States": "US",
    "United States of America": "US",
    "Canada": "US",
    "United Kingdom": "UK",
    "Switzerland": "Switzerland",
    "Japan": "Japan",
    "China": "EM",
    "India": "EM",
    "Brazil": "EM",
    "Mexico": "EM",
    "South Africa": "EM",
    "Turkey": "EM",
    "Poland": "EM",
    "Indonesia": "EM",
    "France": "Eurozone",
    "Germany": "Eurozone",
    "Netherlands": "Eurozone",
    "Italy": "Eurozone",
    "Spain": "Eurozone",
    "Luxembourg": "Eurozone",
    "Ireland": "Eurozone",
    "Belgium": "Eurozone",
    "Austria": "Eurozone",
    "Finland": "Eurozone",
    "Portugal": "Eurozone",
    "European Union": "Eurozone",
    "Sweden": "Global / thematic",
    "Misc.Europe": "Global / thematic",
    "Miscellaneous": "Global / thematic",
    "Stateless": "Global / thematic",
    "Several countries": "Global / thematic",
    "Various countries": "Global / thematic",
    "Indices": "Global / thematic",
    "Currency and precious metal": "Global / thematic",
    "Cayman Islands": "Global / thematic",
    "Guernsey": "Global / thematic",
    "Jersey": "Global / thematic",
}

DEFAULT_EQUITY_REGION = "Global / thematic"


def equity_region_of(country: str | None) -> str:
    """Map a raw 'Geographical breakdown' value to the six equity-sleeve buckets."""
    if not country:
        return DEFAULT_EQUITY_REGION
    return EQUITY_REGION_MAP.get(country.strip(), DEFAULT_EQUITY_REGION)
