"""Service names, country codes, and region labels for FedEx rate exports."""

from __future__ import annotations

import re
from typing import Any

from country_codes import country_name_to_iso2

POSTAL_RANGE_RE = re.compile(
    r"^(.+?)\s+(\d[\d,]*)\s*-\s*(\d[\d,]*)\s*$"
)
FILENAME_COUNTRY_RE = re.compile(r"International([A-Z]{2})", re.IGNORECASE)

SERVICE_CODE_RULES: list[tuple[str, str]] = [
    ("Priority Express", "IPE"),
    ("Priority Freight", "IPF"),
    # Regional Economy Freight before Economy Freight (substring overlap).
    ("Regional Economy Freight", "REF"),
    ("Economy Freight", "IEF"),
    ("Deferred Freight", "FDF"),
    ("Regional Economy", "RE"),
    ("International Economy", "IE"),
    ("International Priority", "IP"),
]


def country_from_source(data: dict[str, Any]) -> str:
    source = data.get("source_file", "")
    match = FILENAME_COUNTRY_RE.search(source)
    if match:
        return match.group(1).upper()
    return ""


def is_export_block(block: dict[str, Any]) -> bool:
    hierarchy = block.get("hierarchy") or []
    service = block.get("service") or ""
    if hierarchy:
        last = hierarchy[-1].lower()
        if "export" in last and "import" not in last:
            return True
        if "import" in last:
            return False
    return "import" not in service.lower()


def shorten_service_name(service: str) -> str:
    name = service.removeprefix("FedEx ").strip()
    name = re.sub(r"\s*\([^)]+\)", "", name).strip()
    name = re.sub(r"\s+(Export|Import)\s*$", "", name, flags=re.IGNORECASE).strip()
    return name


def service_rate_code(service: str) -> str:
    match = re.search(r"\(([A-Z]{2,3})\)", service)
    if match:
        return match.group(1)
    for pattern, code in SERVICE_CODE_RULES:
        if pattern in service:
            return code
    return ""


def direction_label(block: dict[str, Any]) -> str:
    return "Export" if is_export_block(block) else "Import"


def zone_region_label(block: dict[str, Any], zone: str) -> str:
    code = service_rate_code(block.get("service") or "")
    return f"{direction_label(block)} {code} Zone {zone}"


def zone_chart_region_label(chart_name: str, rate_code: str, zone: str) -> str:
    """Region label for zone chart text export (matches XLSX country region column)."""
    return f"{chart_direction(chart_name)} {rate_code} Zone {zone}"


def parse_postal_entry(country_label: str) -> tuple[str, str, str] | None:
    match = POSTAL_RANGE_RE.match(country_label.strip())
    if not match:
        return None
    return match.group(1).strip(), match.group(2), match.group(3)


def normalize_zone_country_name(country_label: str) -> str | None:
    """Return a display country name for zone charts, or None to skip the row."""
    name = country_label.strip()
    if not name or name == "-":
        return None
    if parse_postal_entry(name):
        return None
    if "\n" in name:
        first_line = name.split("\n", 1)[0].strip()
        return first_line or None
    return name


# Postal labels use country name (e.g. "France Zone S"), not service/direction.
# Country-region labels look like "Export RE Zone S" / "Import IEF Zone M".
COUNTRY_REGION_LABEL_RE = re.compile(
    r"^(Export|Import) [A-Z]{2,3} Zone [A-Z0-9]+$"
)
POSTAL_ZONE_LABEL_RE = re.compile(
    r"^(?!Export |Import ).+ Zone [A-Z0-9]+$"
)


def postal_zone_label(country_label: str, zone_letter: str) -> str:
    """Build a postal-zone display label, e.g. 'France Zone S'."""
    parsed = parse_postal_entry(country_label)
    if not parsed or not zone_letter:
        return ""
    return f"{parsed[0]} Zone {zone_letter}"


def postal_zone_labels(
    countries: list[dict[str, Any]], rate_code: str
) -> dict[str, str]:
    """Map zone-chart postal rows to 'Country Zone X' labels (e.g. France Zone S)."""
    result: dict[str, str] = {}
    for entry in countries:
        country_label = entry.get("country", "")
        zone_letter = entry.get("zones", {}).get(rate_code)
        if not zone_letter:
            continue
        label = postal_zone_label(country_label, zone_letter)
        if label:
            result[country_label] = label
    return result


def is_postal_zone_region(region: str) -> bool:
    text = region.strip()
    if COUNTRY_REGION_LABEL_RE.match(text):
        return False
    return bool(POSTAL_ZONE_LABEL_RE.match(text))


def zone_letter_from_postal_region(region: str) -> str | None:
    if not is_postal_zone_region(region):
        return None
    match = re.search(r" Zone ([A-Z0-9]+)$", region.strip())
    return match.group(1) if match else None


def chart_direction(chart_name: str) -> str:
    if "import" in chart_name.lower():
        return "Import"
    return "Export"


def package_cost_key(service: str) -> str:
    """Map package-only services to dedicated transport cost column keys."""
    if "Freight" in service:
        return "package_freight"
    if "Regional Economy" in service:
        return "package_regional_economy"
    if "International Economy" in service:
        return "package_economy"
    return "package"


TRANSPORT_COST_TITLES: dict[str, str] = {
    "envelope": "Transport cost (Envelope)",
    "pak": "Transport cost (Pak)",
    "package": "Transport cost (Package)",
    "package_economy": "Transport cost (Package Economy)",
    "package_regional_economy": "Transport cost (Regional Economy)",
    "package_freight": "Transport cost (Freight)",
}


def transport_cost_title(cost_key: str) -> str:
    return TRANSPORT_COST_TITLES.get(
        cost_key, f"Transport cost ({cost_key})"
    )


def iso_from_postal_zone_label(postal_zone: str) -> str:
    match = re.match(r"^(.+?) Zone [A-Z0-9]+$", postal_zone.strip())
    if not match:
        return ""
    country = match.group(1)
    if re.fullmatch(r"[A-Z]{2}", country):
        return country
    return country_name_to_iso2(country)


def shipment_fields_for_region(
    region: str,
    block: dict[str, Any],
    zone_charts: list[dict[str, Any]],
    home_country: str,
) -> dict[str, str]:
    export = is_export_block(block)

    if is_postal_zone_region(region):
        if export:
            return {
                "origin_country": home_country,
                "origin_country_region": "",
                "origin_postal_zone": "",
                "destination_country": "",
                "destination_country_region": "",
                "destination_postal_zone": region,
            }
        return {
            "origin_country": "",
            "origin_country_region": "",
            "origin_postal_zone": region,
            "destination_country": home_country,
            "destination_country_region": "",
            "destination_postal_zone": "",
        }

    if export:
        return {
            "origin_country": home_country,
            "origin_country_region": "",
            "origin_postal_zone": "",
            "destination_country": "",
            "destination_country_region": region,
            "destination_postal_zone": "",
        }

    return {
        "origin_country": "",
        "origin_country_region": region,
        "origin_postal_zone": "",
        "destination_country": home_country,
        "destination_country_region": "",
        "destination_postal_zone": "",
    }


def chart_for_block(block: dict[str, Any], zone_charts: list[dict[str, Any]]) -> dict[str, Any] | None:
    direction = direction_label(block)
    for chart in zone_charts:
        name = chart.get("name", "")
        if chart_direction(name) == direction:
            return chart
    return None
