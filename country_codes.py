"""Map FedEx zone-chart country names to ISO 3166-1 alpha-2 codes."""

from __future__ import annotations

import re

try:
    import pycountry
except ImportError:
    pycountry = None  # type: ignore[assignment]

# Built-in English country-name → ISO2 map so ISO exports work without pycountry.
# FedEx-specific spellings / territories are included here as well.
COUNTRY_NAME_TO_ISO2: dict[str, str] = {
    "Afghanistan": "AF",
    "Albania": "AL",
    "Algeria": "DZ",
    "American Samoa": "AS",
    "Andorra": "AD",
    "Angola": "AO",
    "Anguilla": "AI",
    "Antigua & Barbuda": "AG",
    "Argentina": "AR",
    "Armenia": "AM",
    "Aruba": "AW",
    "Australia": "AU",
    "Austria": "AT",
    "Azerbaijan": "AZ",
    "Bahama": "BS",
    "Bahamas": "BS",
    "Bahrain": "BH",
    "Bangladesh": "BD",
    "Barbados": "BB",
    "Belarus": "BY",
    "Belgium": "BE",
    "Belize": "BZ",
    "Benin": "BJ",
    "Bermuda": "BM",
    "Bhutan": "BT",
    "Bolivia": "BO",
    "Bonaire Sint Eustatius and Saba": "BQ",
    "Bosnia-Herzegovina": "BA",
    "Bosnia and Herzegovina": "BA",
    "Botswana": "BW",
    "Brazil": "BR",
    "British Virgin Islands": "VG",
    "Brunei": "BN",
    "Bulgaria": "BG",
    "Burkina Faso": "BF",
    "Burundi": "BI",
    "Cambodia": "KH",
    "Cameroon": "CM",
    "Canada": "CA",
    "Canada (Rest of Country)": "CA",
    "Cape Verde": "CV",
    "Cayman Islands": "KY",
    "Central African Republic": "CF",
    "Chad": "TD",
    "Chile": "CL",
    "China": "CN",
    "Colombia": "CO",
    "Congo": "CG",
    "Cook Islands": "CK",
    "Costa Rica": "CR",
    "Cote D'ivoire (Ivory Coast)": "CI",
    "Côte D'ivoire (Ivory Coast)": "CI",
    "Croatia": "HR",
    "Curacao": "CW",
    "Cyprus": "CY",
    "Czech Republic": "CZ",
    "Czechia": "CZ",
    "Democratic Republic of the Congo": "CD",
    "Denmark": "DK",
    "Djibouti": "DJ",
    "Dominica": "DM",
    "Dominican Republic": "DO",
    "East Timor": "TL",
    "Ecuador": "EC",
    "Egypt": "EG",
    "El Salvador": "SV",
    "Equatorial Guinea": "GQ",
    "Eritrea": "ER",
    "Estonia": "EE",
    "Eswatini": "SZ",
    "Ethiopia": "ET",
    "Faeroe Islands": "FO",
    "Faroe Islands": "FO",
    "Fiji": "FJ",
    "Finland": "FI",
    "France": "FR",
    "French Guiana": "GF",
    "French Polynesia": "PF",
    "Gabon": "GA",
    "Gambia": "GM",
    "Georgia": "GE",
    "Germany": "DE",
    "Ghana": "GH",
    "Gibraltar": "GI",
    "Greece": "GR",
    "Greenland": "GL",
    "Grenada": "GD",
    "Guadeloupe": "GP",
    "Guam": "GU",
    "Guatemala": "GT",
    "Guinea": "GN",
    "Guinea-Bissau": "GW",
    "Guyana": "GY",
    "Haiti": "HT",
    "Honduras": "HN",
    "Hong Kong SAR, China": "HK",
    "Hungary": "HU",
    "Iceland": "IS",
    "India": "IN",
    "Indonesia": "ID",
    "Iran": "IR",
    "Iraq": "IQ",
    "Ireland": "IE",
    "Israel": "IL",
    "Italy": "IT",
    "Jamaica": "JM",
    "Japan": "JP",
    "Jordan": "JO",
    "Kazakhstan": "KZ",
    "Kenya": "KE",
    "Kiribati": "KI",
    "Kuwait": "KW",
    "Kyrgyzstan": "KG",
    "Laos": "LA",
    "Latvia": "LV",
    "Lebanon": "LB",
    "Lesotho": "LS",
    "Liberia": "LR",
    "Libya": "LY",
    "Liechtenstein": "LI",
    "Lithuania": "LT",
    "Luxembourg": "LU",
    "Macau SAR, China": "MO",
    "Macedonia": "MK",
    "Madagascar": "MG",
    "Malawi": "MW",
    "Malaysia": "MY",
    "Maldives": "MV",
    "Mali": "ML",
    "Malta": "MT",
    "Marshall Islands": "MH",
    "Martinique": "MQ",
    "Mauritania": "MR",
    "Mauritius": "MU",
    "Mexico": "MX",
    "Micronesia": "FM",
    "Monaco": "MC",
    "Mongolia": "MN",
    "Monserrat": "MS",
    "Montenegro": "ME",
    "Montserrat": "MS",
    "Morocco": "MA",
    "Mozambique": "MZ",
    "Myanmar": "MM",
    "Namibia": "NA",
    "Nauru": "NR",
    "Nepal": "NP",
    "Netherlands": "NL",
    "Netherlands Antilles": "CW",
    "New Caledonia": "NC",
    "New Zealand": "NZ",
    "Nicaragua": "NI",
    "Niger": "NE",
    "Nigeria": "NG",
    "Niue": "NU",
    "Northern Mariana Islands": "MP",
    "Norway": "NO",
    "Oman": "OM",
    "Pakistan": "PK",
    "Palau": "PW",
    "Palestinian Territory": "PS",
    "Panama": "PA",
    "Papua New Guinea": "PG",
    "Paraguay": "PY",
    "Peru": "PE",
    "Phillipines": "PH",
    "Philippines": "PH",
    "Poland": "PL",
    "Portugal": "PT",
    "Puerto Rico": "PR",
    "Qatar": "QA",
    "Republic of Moldova": "MD",
    "Reunion": "RE",
    "Réunion": "RE",
    "Romania": "RO",
    "Russian Federation": "RU",
    "Rwanda": "RW",
    "Saint Lucia": "LC",
    "Samoa": "WS",
    "San Marino": "SM",
    "Saudi Arabia": "SA",
    "Senegal": "SN",
    "Serbia": "RS",
    "Seychelles": "SC",
    "Sierra Leone": "SL",
    "Singapore": "SG",
    "Slovakia": "SK",
    "Slovenia": "SI",
    "Solomon Islands": "SB",
    "Somalia": "SO",
    "South Africa": "ZA",
    "South Korea": "KR",
    "Spain": "ES",
    "Sri Lanka": "LK",
    "St. Kitts and Nevis": "KN",
    "St. Maarten": "SX",
    "St. Martin": "MF",
    "St. Vincent & the Grenadines": "VC",
    "Sudan": "SD",
    "Suriname": "SR",
    "Swaziland": "SZ",
    "Sweden": "SE",
    "Switzerland": "CH",
    "Syrian Arab Republic": "SY",
    "Taiwan": "TW",
    "Thailand": "TH",
    "Togo": "TG",
    "Tonga": "TO",
    "Trinidad & Tobago": "TT",
    "Tunisia": "TN",
    "Turkey": "TR",
    "Turkmenistan": "TM",
    "Turks & Caicos Islands": "TC",
    "Tuvalu": "TV",
    "U.S. Virgin Islands": "VI",
    "U.S.A.": "US",
    "USA": "US",
    "United States": "US",
    "United States of America": "US",
    "Uganda": "UG",
    "Ukraine": "UA",
    "United Arab Emirates": "AE",
    "United Kingdom": "GB",
    "United Kingdom (Great Britain)": "GB",
    "United Republic of Tanzania": "TZ",
    "Uruguay": "UY",
    "Uzbekistan": "UZ",
    "Vanuatu": "VU",
    "Vatican City": "VA",
    "Venezuela": "VE",
    "Vietnam": "VN",
    "Wallis & Futuna": "WF",
    "Yemen": "YE",
    "Zambia": "ZM",
    "Zimbabwe": "ZW",
}

# Backward-compatible alias used elsewhere in the project.
FEDEX_COUNTRY_TO_ISO = COUNTRY_NAME_TO_ISO2

_UK_PREFIX_RE = re.compile(
    r"^United Kingdom(?:\s*\(Great Britain\))?",
    re.IGNORECASE,
)


def country_name_to_iso2(name: str) -> str:
    text = (name or "").strip()
    if not text:
        return text

    if text in COUNTRY_NAME_TO_ISO2:
        return COUNTRY_NAME_TO_ISO2[text]

    if text.startswith("Canada "):
        return "CA"

    if _UK_PREFIX_RE.match(text):
        return "GB"

    if pycountry is not None:
        try:
            return pycountry.countries.lookup(text).alpha_2
        except LookupError:
            pass
        normalized = text.lower()
        for country in pycountry.countries:
            if country.name.lower() == normalized:
                return country.alpha_2
            if getattr(country, "official_name", None):
                if country.official_name.lower() == normalized:
                    return country.alpha_2

    # Last resort: strip parentheticals and retry.
    simplified = re.sub(r"\s*\([^)]*\)", "", text).strip()
    if simplified and simplified != text:
        if simplified in COUNTRY_NAME_TO_ISO2:
            return COUNTRY_NAME_TO_ISO2[simplified]
        if pycountry is not None:
            try:
                return pycountry.countries.lookup(simplified).alpha_2
            except LookupError:
                pass

    return text
