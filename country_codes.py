"""Map FedEx zone-chart country names to ISO 3166-1 alpha-2 codes."""

from __future__ import annotations

import re

try:
    import pycountry
except ImportError:
    pycountry = None  # type: ignore[assignment]

# FedEx-specific spellings and territories not matched by pycountry.lookup
FEDEX_COUNTRY_TO_ISO: dict[str, str] = {
    "Bahama": "BS",
    "Phillipines": "PH",
    "United Kingdom (Great Britain)": "GB",
    "Hong Kong SAR, China": "HK",
    "Macau SAR, China": "MO",
    "Taiwan": "TW",
    "Russian Federation": "RU",
    "Côte D'ivoire (Ivory Coast)": "CI",
    "Cote D'ivoire (Ivory Coast)": "CI",
    "Democratic Republic of the Congo": "CD",
    "Congo": "CG",
    "South Korea": "KR",
    "U.S. Virgin Islands": "VI",
    "British Virgin Islands": "VG",
    "Turks & Caicos Islands": "TC",
    "St. Kitts and Nevis": "KN",
    "St. Maarten": "SX",
    "St. Martin": "MF",
    "St. Vincent & the Grenadines": "VC",
    "Trinidad & Tobago": "TT",
    "Antigua & Barbuda": "AG",
    "Bonaire Sint Eustatius and Saba": "BQ",
    "Netherlands Antilles": "CW",
    "Monserrat": "MS",
    "Faeroe Islands": "FO",
    "Syrian Arab Republic": "SY",
    "Palestinian Territory": "PS",
    "United Republic of Tanzania": "TZ",
    "Republic of Moldova": "MD",
    "Macedonia": "MK",
    "Myanmar": "MM",
    "Laos": "LA",
    "Vietnam": "VN",
    "Brunei": "BN",
    "Cape Verde": "CV",
    "East Timor": "TL",
    "Swaziland": "SZ",
    "Wallis & Futuna": "WF",
    "French Guiana": "GF",
    "Réunion": "RE",
    "Reunion": "RE",
    "Slovakia": "SK",
    "Slovenia": "SI",
    "Bosnia-Herzegovina": "BA",
    "Czech Republic": "CZ",
    "Vatican City": "VA",
    "United Arab Emirates": "AE",
    "Iran": "IR",
    "Micronesia": "FM",
    "Northern Mariana Islands": "MP",
    "Cook Islands": "CK",
    "French Polynesia": "PF",
    "New Caledonia": "NC",
    "American Samoa": "AS",
    "Guam": "GU",
    "Puerto Rico": "PR",
    "China": "CN",
    "Mexico": "MX",
    "Canada": "CA",
    "Canada (Rest of Country)": "CA",
    "Curacao": "CW",
    "Turkey": "TR",
}


def country_name_to_iso2(name: str) -> str:
    if name in FEDEX_COUNTRY_TO_ISO:
        return FEDEX_COUNTRY_TO_ISO[name]

    if name.startswith("Canada "):
        return "CA"

    if pycountry is not None:
        try:
            return pycountry.countries.lookup(name).alpha_2
        except LookupError:
            pass
        normalized = name.strip().lower()
        for country in pycountry.countries:
            if country.name.lower() == normalized:
                return country.alpha_2
            if hasattr(country, "official_name") and country.official_name:
                if country.official_name.lower() == normalized:
                    return country.alpha_2

    # Last resort: strip parentheticals and retry
    simplified = re.sub(r"\s*\([^)]*\)", "", name).strip()
    if simplified != name and simplified in FEDEX_COUNTRY_TO_ISO:
        return FEDEX_COUNTRY_TO_ISO[simplified]
    if pycountry is not None and simplified:
        try:
            return pycountry.countries.lookup(simplified).alpha_2
        except LookupError:
            pass

    return name  # fallback: keep original if unknown
