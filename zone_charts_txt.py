"""Export zone chart data from FedEx rates JSON to text files."""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any, Literal

from country_codes import country_name_to_iso2
from rate_metadata import (
    normalize_zone_country_name,
    parse_postal_entry,
    postal_zone_label,
    zone_chart_region_label,
)

LabelMode = Literal["names", "iso2"]


def _filter_countries_for_display(countries: list[str], mode: LabelMode) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for country in countries:
        normalized = normalize_zone_country_name(country)
        if not normalized:
            continue
        label = (
            country_name_to_iso2(normalized) if mode == "iso2" else normalized
        )
        if label in seen:
            continue
        seen.add(label)
        result.append(label)
    return sorted(result)


def _format_country_list(countries: list[str], mode: LabelMode) -> str:
    return ", ".join(_filter_countries_for_display(countries, mode))


def _postal_range_sort_key(range_str: str) -> int:
    return int(range_str.split("-", 1)[0])


def format_zone_charts_text(
    data: dict[str, Any],
    *,
    label_mode: LabelMode = "names",
    include_header: bool = True,
) -> str:
    lines: list[str] = []

    for chart in data.get("zone_charts", []):
        chart_name = chart.get("name", "Zone Chart")
        if include_header:
            lines.append(chart_name)
            lines.append("")

        for rate_service in chart.get("rate_services", []):
            rate_name = rate_service.get("rate_name", "")
            zones = rate_service.get("zones", [])
            zones_sorted = sorted(zones, key=lambda z: str(z.get("zone", "")))

            for zone_entry in zones_sorted:
                zone_code = zone_entry.get("zone", "")
                countries = zone_entry.get("countries", [])
                country_list = _format_country_list(countries, label_mode)
                region = zone_chart_region_label(chart_name, rate_name, zone_code)
                lines.append(f"{region}: {country_list}")

            lines.append("")

        if include_header:
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def format_zone_charts_text_with_duplicate_iso(data: dict[str, Any]) -> str:
    """Country names first, then the same zones with ISO 3166-1 alpha-2 codes."""
    parts = [
        format_zone_charts_text(data, label_mode="names", include_header=True),
        format_zone_charts_text(data, label_mode="iso2", include_header=False),
    ]
    return "\n".join(parts).rstrip() + "\n"


def format_postal_zones_text(data: dict[str, Any]) -> str:
    """Export postal-code zone mappings in Name / Country / Postal Code layout."""
    lines: list[str] = []

    for chart in data.get("zone_charts", []):
        chart_name = chart.get("name", "Zone Chart")
        countries = chart.get("countries") or []

        lines.append(chart_name)
        lines.append("")
        lines.append("Name")
        lines.append("Country")
        lines.append("Postal Code")
        lines.append("")

        groups: dict[str, dict[str, set[str]]] = defaultdict(
            lambda: defaultdict(set)
        )

        for entry in countries:
            country_label = entry.get("country", "")
            parsed = parse_postal_entry(country_label)
            if not parsed:
                continue
            country_name, start, end = parsed
            range_str = f"{start}-{end}"
            country_iso = country_name_to_iso2(country_name)

            for zone_letter in entry.get("zones", {}).values():
                zone_label = postal_zone_label(country_label, zone_letter)
                if zone_label:
                    groups[zone_label][country_iso].add(range_str)

        for zone_label in sorted(groups.keys()):
            for country_iso in sorted(groups[zone_label].keys()):
                ranges = sorted(
                    groups[zone_label][country_iso],
                    key=_postal_range_sort_key,
                )
                lines.append(zone_label)
                lines.append(country_iso)
                lines.append("; ".join(ranges))
                lines.append("")

        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def write_postal_zones_data(
    data: dict[str, Any],
    output_path: str | Path,
) -> Path:
    output_path = Path(output_path)
    text = format_postal_zones_text(data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return output_path


def write_postal_zones_txt(
    json_path: str | Path,
    output_path: str | Path,
) -> Path:
    with Path(json_path).open(encoding="utf-8") as f:
        data = json.load(f)
    return write_postal_zones_data(data, output_path)


def write_zone_charts_data(
    data: dict[str, Any],
    output_path: str | Path,
    *,
    include_iso_duplicate: bool = True,
) -> Path:
    output_path = Path(output_path)
    if include_iso_duplicate:
        text = format_zone_charts_text_with_duplicate_iso(data)
    else:
        text = format_zone_charts_text(data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return output_path


def write_zone_charts_iso_data(
    data: dict[str, Any],
    output_path: str | Path,
) -> Path:
    output_path = Path(output_path)
    text = format_zone_charts_text(data, label_mode="iso2", include_header=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    return output_path


def write_zone_charts_txt(
    json_path: str | Path,
    output_path: str | Path,
    *,
    include_iso_duplicate: bool = True,
) -> Path:
    with Path(json_path).open(encoding="utf-8") as f:
        data = json.load(f)
    return write_zone_charts_data(data, output_path, include_iso_duplicate=include_iso_duplicate)


def write_zone_charts_iso_txt(
    json_path: str | Path,
    output_path: str | Path,
) -> Path:
    with Path(json_path).open(encoding="utf-8") as f:
        data = json.load(f)
    return write_zone_charts_iso_data(data, output_path)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Export FedEx zone charts from JSON to text file(s)."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="output/fedex_rates.json",
        help="Input JSON path",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="output/fedex_zones.txt",
        help="Output text file",
    )
    parser.add_argument(
        "--no-iso-duplicate",
        action="store_true",
        help="Omit the duplicate section with ISO 3166-1 alpha-2 country codes",
    )
    parser.add_argument(
        "--iso-only",
        default="",
        help="Also write ISO-only file at this path (optional)",
    )
    args = parser.parse_args()

    out = write_zone_charts_txt(
        args.input,
        args.output,
        include_iso_duplicate=not args.no_iso_duplicate,
    )
    print(f"Wrote {out}")

    if args.iso_only:
        iso_out = write_zone_charts_iso_txt(args.input, args.iso_only)
        print(f"Wrote {iso_out}")


if __name__ == "__main__":
    main()
