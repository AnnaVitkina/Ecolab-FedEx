"""
Convert parsed FedEx rates JSON into a matrix-style XLSX workbook.

Unified export/import table with shared Transport cost columns.
"""

from __future__ import annotations

import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from rate_metadata import (
    chart_for_block,
    country_from_source,
    direction_label,
    is_postal_zone_region,
    package_cost_key,
    parse_postal_entry,
    postal_zone_label,
    service_rate_code,
    shipment_fields_for_region,
    shorten_service_name,
    transport_cost_title,
    zone_letter_from_postal_region,
    zone_region_label,
)

SHIPMENT_HEADERS = [
    "Origin Country",
    "Origin Country Region",
    "Origin Postal Code Zone",
    "Destination Country",
    "Destination Country Region",
    "Destination Postal Code Zone",
    "Service",
]

COST_KEY_ORDER = [
    "envelope",
    "pak",
    "package",
    "package_economy",
    "package_regional_economy",
    "package_freight",
]
WEIGHT_RANGE_RE = re.compile(
    r"^\s*([\d,]+(?:\.\d+)?)\s*-\s*([\d,]+(?:\.\d+)?)\s*$"
)


def _parse_number(text: str) -> float:
    text = text.strip()
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    return float(text)


def format_weight_label(
    entry: dict[str, Any],
    pricing_model: str,
    *,
    open_ended_replacement: dict[str, str] | None = None,
) -> str:
    rate_type = (entry.get("rate_type") or "").lower()
    weight = (entry.get("weight") or "").strip()

    if rate_type == "envelope" and not weight:
        return "Flat"

    if weight and WEIGHT_RANGE_RE.match(weight):
        if open_ended_replacement and weight in open_ended_replacement:
            return open_ended_replacement[weight]
        parts = re.split(r"\s*-\s*", weight)
        lower = _parse_number(parts[0])
        upper = _parse_number(parts[1])
        if upper >= 99999:
            if lower == int(lower):
                return f">={int(lower)}"
            return f">={lower}"
        if upper == int(upper):
            return str(int(upper))
        return str(upper)

    if weight:
        value = _parse_number(weight)
        if value == int(value):
            return str(int(value))
        return str(value)

    if pricing_model in ("per_pound", "per_kilogram"):
        return "Per kg"
    return ""


def open_ended_weight_labels(entries: list[dict[str, Any]]) -> dict[str, str]:
    weighted: list[tuple[str, str]] = []
    open_ranges: list[str] = []

    for entry in entries:
        weight = (entry.get("weight") or "").strip()
        if not weight:
            continue
        if WEIGHT_RANGE_RE.match(weight):
            parts = re.split(r"\s*-\s*", weight)
            upper = _parse_number(parts[1])
            if upper >= 99999:
                open_ranges.append(weight)
            else:
                weighted.append((weight, weight))
        else:
            weighted.append((weight, weight))

    replacements: dict[str, str] = {}
    if not open_ranges:
        return replacements

    if len(entries) == 1 and len(open_ranges) == 1:
        parts = re.split(r"\s*-\s*", open_ranges[0])
        lower = _parse_number(parts[0])
        if lower == int(lower):
            replacements[open_ranges[0]] = f">={int(lower)}"
        else:
            replacements[open_ranges[0]] = f">={lower}"
        return replacements

    numeric_weights: list[tuple[float, str]] = []
    for entry in entries:
        weight = (entry.get("weight") or "").strip()
        if not weight or not WEIGHT_RANGE_RE.match(weight):
            if weight:
                try:
                    numeric_weights.append((_parse_number(weight), weight))
                except ValueError:
                    pass
            continue
        parts = re.split(r"\s*-\s*", weight)
        upper = _parse_number(parts[1])
        if upper < 99999:
            numeric_weights.append((upper, weight))

    last_numeric = max((value for value, _ in numeric_weights), default=None)

    for open_weight in open_ranges:
        parts = re.split(r"\s*-\s*", open_weight)
        lower = _parse_number(parts[0])
        if last_numeric is not None:
            threshold = last_numeric
        else:
            threshold = lower - 0.5 if lower > 0 else lower
        if threshold == int(threshold):
            replacements[open_weight] = f">{int(threshold)}"
        else:
            replacements[open_weight] = f">{threshold}"

    return replacements


def weight_sort_key(label: str) -> tuple[int, float, str]:
    lower = label.lower()
    if lower == "flat":
        return (0, 0.0, label)
    if lower == "per kg":
        return (4, 99999.0, label)
    match = re.match(r">=(\d+(?:\.\d+)?)", label)
    if match:
        return (3, float(match.group(1)), label)
    match = re.match(r">(\d+(?:\.\d+)?)", label)
    if match:
        return (3, float(match.group(1)), label)
    try:
        return (1, float(label), label)
    except ValueError:
        return (2, 0.0, label)


def pricing_label_for_group(
    weight_labels: list[str], pricing_model: str
) -> str:
    if pricing_model in ("per_pound", "per_kilogram"):
        return "Rate by Weight/chargeable kg"
    has_brackets = any(label.lower() not in ("flat", "per kg") for label in weight_labels)
    if has_brackets:
        return "Rate by Weight/chargeable kg"
    return "Rate by: per shipment"


def value_header_for_group(weight_labels: list[str]) -> str:
    if len(weight_labels) == 1 and weight_labels[0].lower() == "flat":
        return "Flat"
    return "p/unit"


def value_headers_for_labels(weight_labels: list[str]) -> list[str]:
    return [
        "Flat" if label.lower() == "flat" else "p/unit" for label in weight_labels
    ]


def resolve_cost_key(block: dict[str, Any], rate_type: str) -> str:
    if rate_type == "envelope":
        return "envelope"
    if rate_type == "pak":
        return "pak"
    return package_cost_key(block.get("service") or "")


@dataclass
class CostColumnGroup:
    cost_key: str
    pricing_model: str
    title: str
    pricing_label: str
    value_header: str
    currency: str
    weight_labels: list[str]
    prices: dict[str, dict[str, float]] = field(default_factory=dict)

    @property
    def width(self) -> int:
        return 1 + len(self.weight_labels)

    def price_for(self, row_key: str, weight: str) -> float | None:
        return self.prices.get(row_key, {}).get(weight)


@dataclass
class MatrixRow:
    origin_country: str
    origin_country_region: str
    origin_postal_zone: str
    destination_country: str
    destination_country_region: str
    destination_postal_zone: str
    service: str
    currency: str
    row_key: str


def row_key_for(service: str, shipment: dict[str, str]) -> str:
    return "|".join(
        [
            service,
            shipment.get("origin_country_region", ""),
            shipment.get("origin_postal_zone", ""),
            shipment.get("destination_country_region", ""),
            shipment.get("destination_postal_zone", ""),
        ]
    )


def _zone_letter_sort_key(letter: str) -> tuple[int, str]:
    return (len(letter), letter)


def _region_zone_letter(region: str) -> str:
    if is_postal_zone_region(region):
        return zone_letter_from_postal_region(region) or ""
    match = re.search(r" Zone ([A-Z0-9]+)$", region)
    return match.group(1) if match else ""


def _row_primary_region(row: MatrixRow) -> str:
    if row.origin_country:
        return row.destination_country_region or row.destination_postal_zone
    return row.origin_country_region or row.origin_postal_zone


def _row_is_postal(row: MatrixRow) -> bool:
    if row.origin_country:
        return bool(row.destination_postal_zone) and not row.destination_country_region
    return bool(row.origin_postal_zone) and not row.origin_country_region


def _row_sort_key(row: MatrixRow) -> tuple[Any, ...]:
    region = _row_primary_region(row)
    return (
        row.service,
        0 if row.origin_country else 1,
        _zone_letter_sort_key(_region_zone_letter(region)),
        1 if _row_is_postal(row) else 0,
        region,
    )


def build_regions_for_block(
    block: dict[str, Any],
    zone_charts: list[dict[str, Any]],
) -> list[str]:
    chart = chart_for_block(block, zone_charts)
    if not chart:
        zones = block.get("rates", {}).get("zones") or []
        return [zone_region_label(block, zone) for zone in zones]

    code = service_rate_code(block.get("service") or "")
    direction = direction_label(block)
    regions: list[str] = []
    seen: set[str] = set()

    postal_by_zone: dict[str, list[str]] = defaultdict(list)
    postal_seen: dict[str, set[str]] = defaultdict(set)
    for entry in chart.get("countries") or []:
        country_label = entry.get("country", "")
        zones_map = entry.get("zones") or {}
        zone_letter = zones_map.get(code)
        if not zone_letter or not parse_postal_entry(country_label):
            continue
        label = postal_zone_label(country_label, zone_letter)
        if label and label not in postal_seen[zone_letter]:
            postal_seen[zone_letter].add(label)
            postal_by_zone[zone_letter].append(label)

    zone_letters = set(block.get("rates", {}).get("zones") or [])
    for zone in sorted(zone_letters, key=_zone_letter_sort_key):
        label = f"{direction} {code} Zone {zone}"
        if label not in seen:
            regions.append(label)
            seen.add(label)
        for postal_label in sorted(postal_by_zone.get(zone, [])):
            if postal_label not in seen:
                regions.append(postal_label)
                seen.add(postal_label)

    return regions


def zone_letter_for_region(
    block: dict[str, Any],
    region: str,
    zone_charts: list[dict[str, Any]],
) -> str | None:
    code = service_rate_code(block.get("service") or "")
    direction = direction_label(block)

    zone_match = re.match(
        rf"^{re.escape(direction)} {re.escape(code)} Zone ([A-Z0-9]+)$",
        region,
    )
    if zone_match:
        return zone_match.group(1)

    if is_postal_zone_region(region):
        return zone_letter_from_postal_region(region)

    return None


def build_cost_groups_and_rows(
    blocks: list[dict[str, Any]],
    zone_charts: list[dict[str, Any]],
    home_country: str,
) -> tuple[list[CostColumnGroup], list[MatrixRow]]:
    group_specs: dict[str, dict[str, Any]] = {}
    matrix_rows: dict[str, MatrixRow] = {}

    for block in blocks:
        service = shorten_service_name(block.get("service") or "")
        pricing_model = block.get("pricing_model") or "net"
        currency = block.get("currency") or "EUR"
        entries = block.get("rates", {}).get("entries") or []
        if not entries:
            continue

        open_ended = open_ended_weight_labels(entries)
        entries_by_type: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for entry in entries:
            rate_type = (entry.get("rate_type") or "package").lower()
            entries_by_type[rate_type].append(entry)

        regions = build_regions_for_block(block, zone_charts)

        for rate_type, type_entries in entries_by_type.items():
            cost_key = resolve_cost_key(block, rate_type)
            weight_labels: list[str] = []
            label_prices: dict[str, dict[str, float]] = defaultdict(dict)

            for entry in type_entries:
                weight_label = format_weight_label(
                    entry, pricing_model, open_ended_replacement=open_ended
                )
                if not weight_label:
                    continue
                if weight_label not in weight_labels:
                    weight_labels.append(weight_label)

                zone_prices = entry.get("zones") or {}
                for region in regions:
                    zone_letter = zone_letter_for_region(block, region, zone_charts)
                    if zone_letter is None:
                        continue
                    price = zone_prices.get(zone_letter)
                    if price is None:
                        continue
                    label_prices[region][weight_label] = float(price)

            if not weight_labels:
                continue

            weight_labels.sort(key=weight_sort_key)

            if cost_key not in group_specs:
                group_specs[cost_key] = {
                    "cost_key": cost_key,
                    "currency": currency,
                    "weight_labels": [],
                    "prices": defaultdict(dict),
                    "has_per_kg": False,
                    "has_brackets": False,
                }

            spec = group_specs[cost_key]
            if pricing_model in ("per_pound", "per_kilogram"):
                spec["has_per_kg"] = True
            if any(
                label.lower() not in ("flat", "per kg") for label in weight_labels
            ):
                spec["has_brackets"] = True

            for label in weight_labels:
                if label not in spec["weight_labels"]:
                    spec["weight_labels"].append(label)
            spec["weight_labels"].sort(key=weight_sort_key)

            for region in regions:
                shipment = shipment_fields_for_region(
                    region, block, zone_charts, home_country
                )
                rk = row_key_for(service, shipment)
                if rk not in matrix_rows:
                    matrix_rows[rk] = MatrixRow(
                        origin_country=shipment["origin_country"],
                        origin_country_region=shipment["origin_country_region"],
                        origin_postal_zone=shipment["origin_postal_zone"],
                        destination_country=shipment["destination_country"],
                        destination_country_region=shipment[
                            "destination_country_region"
                        ],
                        destination_postal_zone=shipment["destination_postal_zone"],
                        service=service,
                        currency=currency,
                        row_key=rk,
                    )

                for weight_label in weight_labels:
                    price = label_prices[region].get(weight_label)
                    if price is not None:
                        spec["prices"][rk][weight_label] = price

    cost_groups: list[CostColumnGroup] = []
    for cost_key in COST_KEY_ORDER:
        if cost_key not in group_specs:
            continue
        spec = group_specs[cost_key]
        weight_labels = spec["weight_labels"]
        if spec["has_per_kg"] or spec["has_brackets"]:
            pricing_label = "Rate by Weight/chargeable kg"
        else:
            pricing_label = "Rate by: per shipment"

        cost_groups.append(
            CostColumnGroup(
                cost_key=cost_key,
                pricing_model="net",
                title=transport_cost_title(cost_key),
                pricing_label=pricing_label,
                value_header=value_header_for_group(weight_labels),
                currency=spec["currency"],
                weight_labels=weight_labels,
                prices=dict(spec["prices"]),
            )
        )

    rows = sorted(matrix_rows.values(), key=_row_sort_key)
    return cost_groups, rows


class FedExRatesXlsxWriter:
    def __init__(self, data: dict[str, Any]):
        self.data = data
        self.home_country = country_from_source(data)
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Rates"
        self._row = 1
        self._bold = Font(bold=True)
        self._center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def write(self) -> Workbook:
        blocks = self.data.get("rate_blocks", [])
        zone_charts = self.data.get("zone_charts", [])
        cost_groups, rows = build_cost_groups_and_rows(
            blocks, zone_charts, self.home_country
        )
        if cost_groups and rows:
            self._write_matrix(cost_groups, rows)
        self._autosize_columns()
        return self.wb

    def _left_col_count(self) -> int:
        return len(SHIPMENT_HEADERS)

    def _write_matrix(
        self,
        cost_groups: list[CostColumnGroup],
        rows: list[MatrixRow],
    ) -> None:
        left_cols = self._left_col_count()
        total_cols = left_cols + sum(cg.width for cg in cost_groups)

        col = left_cols + 1
        for cg in cost_groups:
            self.ws.cell(row=self._row, column=col, value=cg.title)
            self.ws.cell(row=self._row, column=col).font = self._bold
            self.ws.cell(row=self._row, column=col).alignment = self._center
            if cg.width > 1:
                self.ws.merge_cells(
                    start_row=self._row,
                    start_column=col,
                    end_row=self._row,
                    end_column=col + cg.width - 1,
                )
            col += cg.width
        self._row += 1

        col = left_cols + 1
        for cg in cost_groups:
            self.ws.cell(row=self._row, column=col, value=cg.pricing_label)
            self.ws.cell(row=self._row, column=col).font = self._bold
            self.ws.cell(row=self._row, column=col).alignment = self._center
            if cg.width > 1:
                self.ws.merge_cells(
                    start_row=self._row,
                    start_column=col,
                    end_row=self._row,
                    end_column=col + cg.width - 1,
                )
            col += cg.width
        self._row += 1

        col = left_cols + 1
        for cg in cost_groups:
            if len(cg.weight_labels) == 1:
                self.ws.cell(row=self._row, column=col, value=cg.weight_labels[0])
            else:
                for wi, weight in enumerate(cg.weight_labels):
                    cell = self.ws.cell(
                        row=self._row, column=col + 1 + wi, value=weight
                    )
                    cell.font = self._bold
                    cell.alignment = self._center
            col += cg.width
        self._row += 1

        for col_idx, header in enumerate(SHIPMENT_HEADERS, start=1):
            cell = self.ws.cell(row=self._row, column=col_idx, value=header)
            cell.font = self._bold

        col = left_cols + 1
        for cg in cost_groups:
            self.ws.cell(row=self._row, column=col, value="Currency")
            self.ws.cell(row=self._row, column=col).font = self._bold
            value_headers = value_headers_for_labels(cg.weight_labels)
            for wi, header in enumerate(value_headers):
                cell = self.ws.cell(
                    row=self._row,
                    column=col + 1 + wi,
                    value=header,
                )
                cell.font = self._bold
            col += cg.width
        self._row += 1

        for row in rows:
            shipment = [
                row.origin_country,
                row.origin_country_region,
                row.origin_postal_zone,
                row.destination_country,
                row.destination_country_region,
                row.destination_postal_zone,
                row.service,
            ]
            for col_idx, value in enumerate(shipment, start=1):
                self.ws.cell(row=self._row, column=col_idx, value=value or None)

            col = left_cols + 1
            for cg in cost_groups:
                prices = [
                    cg.price_for(row.row_key, weight) for weight in cg.weight_labels
                ]
                has_price = any(price is not None for price in prices)
                if has_price:
                    self.ws.cell(row=self._row, column=col, value=row.currency)
                for wi, weight in enumerate(cg.weight_labels):
                    price = prices[wi]
                    cell = self.ws.cell(
                        row=self._row,
                        column=col + 1 + wi,
                        value=price if price is not None else None,
                    )
                    if isinstance(price, float):
                        cell.number_format = "0.00"
                col += cg.width
            self._row += 1

        _ = total_cols

    def _autosize_columns(self, max_width: int = 24) -> None:
        for col_cells in self.ws.columns:
            letter = get_column_letter(col_cells[0].column)
            length = max(
                (len(str(cell.value)) for cell in col_cells if cell.value is not None),
                default=8,
            )
            self.ws.column_dimensions[letter].width = min(length + 2, max_width)


def data_to_xlsx(data: dict[str, Any], output_path: str | Path) -> Path:
    output_path = Path(output_path)
    writer = FedExRatesXlsxWriter(data)
    wb = writer.write()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def json_to_xlsx(
    json_path: str | Path,
    output_path: str | Path,
) -> Path:
    json_path = Path(json_path)
    with json_path.open(encoding="utf-8") as f:
        data = json.load(f)
    return data_to_xlsx(data, output_path)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert FedEx rates JSON to matrix-style XLSX."
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
        default="output/fedex_rates_matrix.xlsx",
        help="Output XLSX path",
    )
    args = parser.parse_args()
    out = json_to_xlsx(args.input, args.output)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
