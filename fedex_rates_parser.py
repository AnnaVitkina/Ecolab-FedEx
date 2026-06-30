"""
Parse FedEx courtesy net rate spreadsheets (XML SpreadsheetML .xls) into JSON.

Input files are XML workbooks exported by FedEx (not binary Excel).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

NS = "urn:schemas-microsoft-com:office:spreadsheet"
BOLD_STYLES = {"s2", "s3"}

SKIP_SECTION_TITLES = {
    "Weight",
    "Zone(s)",
    "lbs.",
    "lb.",
    "Kgs",
    "Kg",
    "Customer Name:",
    "Effective Date:",
    "Created On:",
    "Customer Name / Location(s):",
    "Account Number",
    "Country/Territory",
    "Package Type",
    "Earned Discount",
    "Billing Type:",
    "IPE",
    "IP",
    "IE",
    "IPF",
    "IEF",
    "FDF",
    "RE",
    "REF",
}

WEIGHT_UNIT_HEADERS = {"lbs.", "lb.", "kgs", "kg"}
DELIVERY_MODE_RE = re.compile(r"\b(Airport|Door)\b.*\bto\b", re.IGNORECASE)

RATE_TYPE_LABELS = {"envelope", "pak", "package"}
ZONE_HEADER_RE = re.compile(r"^[A-Z]$|^\d{1,2}$")
WEIGHT_RANGE_RE = re.compile(
    r"^\s*([\d,]+(?:\.\d+)?)\s*-\s*([\d,]+(?:\.\d+)?)\s*$"
)
WEIGHT_NUM_RE = re.compile(r"^[\d,]+(?:\.\d+)?$")
LOCATION_APPLICABILITY_RE = re.compile(
    r"Net rates are applicable for location\s+(\d+)\s+only\.?",
    re.IGNORECASE,
)
EFFECTIVE_DATE_NOTE_RE = re.compile(
    r"Net rates are effective as of\s+(.+?)\.?\s*$",
    re.IGNORECASE,
)
ACCOUNT_NUMBER_RE = re.compile(r"^\d+$")


def _tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def _parse_number(text: str) -> float:
    text = text.strip()
    if not text:
        raise ValueError("empty number")
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


def _is_delivery_mode_title(text: str) -> bool:
    return bool(DELIVERY_MODE_RE.search(text))


def _normalize_zone_chart_title(text: str) -> str:
    # "ExportZone Chart" -> "Export Zone Chart"
    return re.sub(r"ExportZone", "Export Zone", text)


@dataclass
class RowData:
    cells: list[str]
    bold_texts: list[str]


class FedExRatesParser:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        tree = ET.parse(self.path)
        self._rows = self._load_rows(tree.getroot())

    def _load_rows(self, root: ET.Element) -> list[RowData]:
        rows: list[RowData] = []
        for row_el in root.findall(f".//{_tag('Row')}"):
            cells: list[str] = []
            bold_texts: list[str] = []
            col = 0
            for cell in row_el.findall(_tag("Cell")):
                index = cell.get(_tag("Index"))
                if index:
                    col = int(index) - 1
                merge = cell.get(_tag("MergeAcross"))
                span = int(merge) + 1 if merge else 1
                data = cell.find(_tag("Data"))
                text = (data.text or "").strip() if data is not None else ""
                style = cell.get(_tag("StyleID"), "")
                while len(cells) <= col:
                    cells.append("")
                cells[col] = text
                if style in BOLD_STYLES and text:
                    bold_texts.append(text)
                for _ in range(1, span):
                    col += 1
                    while len(cells) <= col:
                        cells.append("")
                    cells.append("")
                col += 1
            rows.append(RowData(cells=cells, bold_texts=bold_texts))
        return rows

    def parse(self) -> dict[str, Any]:
        document = self._parse_document()
        hierarchy: list[str] = []
        rate_blocks: list[dict[str, Any]] = []
        zone_charts: list[dict[str, Any]] = []
        current_account: dict[str, str | None] = {
            "customer_name_location": None,
            "account_number": None,
        }
        pending_account_values = False

        i = 0
        while i < len(self._rows):
            row = self._rows[i]

            if pending_account_values:
                values = self._extract_account_values(row)
                if values:
                    current_account = values
                pending_account_values = False
                i += 1
                continue

            if self._is_account_header_row(row):
                pending_account_values = True
                i += 1
                continue

            if self._is_zone_chart_title(row):
                chart_start = i
                chart, i = self._parse_zone_chart(i)
                zone_charts.append(chart)
                for j in range(chart_start + 1, i):
                    section = self._extract_section_title(self._rows[j])
                    if section:
                        hierarchy = self._update_hierarchy(hierarchy, section)
                continue

            section = self._extract_section_title(row)
            if section:
                hierarchy = self._update_hierarchy(hierarchy, section)

            service_info = self._extract_service_header(row)
            if service_info:
                service_name, currency = service_info
                subtype = self._leading_subtype(row, service_name)
                notes = self._extract_service_notes(i)
                if not subtype:
                    subtype = notes.get("subtype")
                pricing_model = self._detect_pricing_model(i)
                table, i = self._parse_rate_table(i)
                rate_blocks.append(
                    {
                        "hierarchy": list(hierarchy),
                        "service": service_name,
                        "subtype": subtype,
                        "currency": currency,
                        "pricing_model": pricing_model,
                        "effective_date_note": notes["effective_date_note"],
                        "location_applicability_note": notes[
                            "location_applicability_note"
                        ],
                        "applicable_locations": notes["applicable_locations"],
                        "customer_name_location": current_account[
                            "customer_name_location"
                        ],
                        "account_number": current_account["account_number"],
                        "rates": table,
                    }
                )
                continue

            i += 1

        return {
            "source_file": self.path.name,
            "document": document,
            "rate_blocks": rate_blocks,
            "zone_charts": zone_charts,
        }

    def _parse_document(self) -> dict[str, Any]:
        meta: dict[str, Any] = {}
        accounts: list[dict[str, str]] = []
        for idx, row in enumerate(self._rows):
            texts = [c for c in row.cells if c]
            joined = " ".join(texts)
            if not texts:
                continue
            if texts[0] == "FedEx Ratesheets":
                meta["title"] = texts[0]
            elif len(texts) >= 2 and texts[0] == "Customer Name:":
                meta["customer_name"] = texts[1]
            elif len(texts) >= 2 and texts[0] == "Effective Date:":
                meta["effective_date"] = texts[1]
            elif len(texts) >= 2 and texts[0] == "Created On:":
                meta["created_on"] = texts[1]
            elif self._is_account_header_row(row) and idx + 1 < len(self._rows):
                values = self._extract_account_values(self._rows[idx + 1])
                if values:
                    accounts.append(values)

        if accounts:
            meta["accounts"] = accounts
            meta["account_number"] = accounts[0]["account_number"]
            meta["customer_name_location"] = accounts[0]["customer_name_location"]
        return meta

    def _is_account_header_row(self, row: RowData) -> bool:
        texts = [c for c in row.cells if c]
        return (
            "Customer Name / Location(s):" in texts and "Account Number" in texts
        )

    def _extract_account_values(self, row: RowData) -> dict[str, str] | None:
        texts = [c for c in row.cells if c]
        if len(texts) < 2:
            return None
        customer_name_location, account_number = texts[0], texts[1]
        if not ACCOUNT_NUMBER_RE.match(account_number):
            return None
        return {
            "customer_name_location": customer_name_location,
            "account_number": account_number,
        }

    def _is_zone_chart_title(self, row: RowData) -> bool:
        joined = "".join(row.cells)
        return "Zone Chart" in joined or "ZoneChart" in joined.replace(" ", "")

    def _extract_section_title(self, row: RowData) -> str | None:
        if not row.bold_texts:
            return None
        # Single bold section header (ignore table header rows)
        candidates = [t for t in row.bold_texts if t not in SKIP_SECTION_TITLES]
        if not candidates:
            return None
        if "Weight" in row.bold_texts and "Zone(s)" in row.bold_texts:
            return None
        if all(ZONE_HEADER_RE.match(t) for t in candidates):
            return None
        # Service rows handled elsewhere
        if any("Currency:" in c for c in row.cells):
            return None
        if len(candidates) == 1:
            return candidates[0]
        # Multi-bold rows like customer header
        if "Customer Name / Location(s):" in candidates:
            return None
        if any(_is_delivery_mode_title(t) for t in candidates):
            return None
        return candidates[0]

    def _update_hierarchy(self, hierarchy: list[str], title: str) -> list[str]:
        if title in ("FedEx ExpressNet Rates", "FedEx GroundNet Rates"):
            return [title]
        if title in ("U.S.A. Express Rates", "U.S.A. Ground Rates", "Belgium Express Rates"):
            base = hierarchy[:1] if hierarchy else []
            return base + [title]
        # Replace same-level geographic grouping
        geographic_prefixes = (
            "United States ",
            "FedEx International",
            "Belgium FedEx International",
        )
        if any(title.startswith(p) for p in geographic_prefixes):
            base = hierarchy[:2] if len(hierarchy) >= 2 else hierarchy[:]
            return base + [title]
        # Freight / delivery sub-headers under international
        if hierarchy and hierarchy[0] == "FedEx ExpressNet Rates":
            if len(hierarchy) >= 3 and hierarchy[2].startswith("United States FedEx International"):
                return hierarchy[:3] + [title]
        # Ground product family
        if hierarchy and hierarchy[0] == "FedEx GroundNet Rates":
            if len(hierarchy) >= 2:
                return hierarchy[:2] + [title]
        return hierarchy + [title]

    def _extract_service_header(self, row: RowData) -> tuple[str, str] | None:
        texts = [c for c in row.cells if c]
        if len(texts) < 2:
            return None
        currency = next((t for t in texts if t.startswith("Currency:")), None)
        if not currency:
            return None
        service = texts[0]
        if service.startswith("Currency:"):
            return None
        return service, currency.replace("Currency:", "").strip()

    def _leading_subtype(self, row: RowData, service_name: str) -> str | None:
        for text in row.bold_texts:
            if text != service_name and text not in SKIP_SECTION_TITLES:
                return text
        return None

    def _extract_service_notes(self, start_index: int) -> dict[str, Any]:
        """Read per-service disclaimer rows between the header and rate table."""
        effective_date_note: str | None = None
        location_applicability_note: str | None = None
        applicable_locations: list[str] = []
        subtype: str | None = None

        for j in range(start_index + 1, min(start_index + 20, len(self._rows))):
            row = self._rows[j]
            texts = [c for c in row.cells if c]
            if not texts:
                continue

            joined = " ".join(texts)
            if "Per-pound rates" in joined or "Per-kilogram rates" in joined:
                break
            if "Weight" in texts and "Zone(s)" in texts:
                break
            if self._extract_service_header(row):
                break

            for text in texts:
                if not subtype and _is_delivery_mode_title(text):
                    subtype = text

                if not effective_date_note:
                    if EFFECTIVE_DATE_NOTE_RE.search(text) or text.startswith(
                        "Net rates are effective as of"
                    ):
                        effective_date_note = text

                loc_match = LOCATION_APPLICABILITY_RE.search(text)
                if loc_match:
                    if not location_applicability_note:
                        location_applicability_note = text
                    code = loc_match.group(1)
                    if code not in applicable_locations:
                        applicable_locations.append(code)

        return {
            "effective_date_note": effective_date_note,
            "location_applicability_note": location_applicability_note,
            "applicable_locations": applicable_locations,
            "subtype": subtype,
        }

    def _detect_pricing_model(self, start_index: int) -> str:
        for j in range(start_index, min(start_index + 12, len(self._rows))):
            joined = " ".join(c for c in self._rows[j].cells if c)
            if "Per-pound rates" in joined:
                return "per_pound"
            if "Per-kilogram rates" in joined:
                return "per_kilogram"
        return "net"

    def _parse_rate_table(self, start_index: int) -> tuple[dict[str, Any], int]:
        i = start_index + 1
        zones: list[str] = []
        rates: dict[str, Any] = {
            "zones": [],
            "entries": [],
        }
        current_rate_type: str | None = None

        while i < len(self._rows):
            row = self._rows[i]
            cells = row.cells
            texts = [c for c in cells if c]

            if self._is_zone_chart_title(row):
                break

            if self._extract_service_header(row):
                break

            section = self._extract_section_title(row)
            if section and section not in RATE_TYPE_LABELS:
                if not _is_delivery_mode_title(section):
                    break

            if "Weight" in texts and "Zone(s)" in texts:
                i += 1
                if i < len(self._rows):
                    zone_row = self._rows[i]
                    zones = self._parse_zone_headers(zone_row.cells)
                    rates["zones"] = zones
                    i += 1
                continue

            if not zones:
                i += 1
                continue

            if not texts:
                i += 1
                continue

            first_cell = cells[0] if cells else ""
            first = first_cell.lower()
            if first in RATE_TYPE_LABELS:
                current_rate_type = first
                i += 1
                continue

            zone_values = self._aligned_zone_values(cells, zones)
            if zone_values is None:
                i += 1
                continue

            if self._is_per_pound_weight_row(first_cell):
                weight_info = self._parse_weight_range(first_cell)
                rates["entries"].append(
                    {
                        "rate_type": current_rate_type or "package",
                        "weight": weight_info["label"],
                        "weight_min": weight_info["min"],
                        "weight_max": weight_info["max"],
                        "zones": self._zip_zones(zones, zone_values),
                    }
                )
                i += 1
                continue

            if current_rate_type in ("envelope", "pak") and not first_cell:
                rates["entries"].append(
                    {
                        "rate_type": current_rate_type,
                        "weight": "",
                        "zones": self._zip_zones(zones, zone_values),
                    }
                )
                i += 1
                continue

            if first_cell and WEIGHT_NUM_RE.match(first_cell):
                rates["entries"].append(
                    {
                        "rate_type": current_rate_type or "package",
                        "weight": first_cell,
                        "zones": self._zip_zones(zones, zone_values),
                    }
                )
                i += 1
                continue

            i += 1

        return rates, i

    def _parse_zone_headers(self, cells: list[str]) -> list[str]:
        texts = [c for c in cells if c]
        if texts and texts[0].lower() in WEIGHT_UNIT_HEADERS:
            return texts[1:]
        return texts

    def _aligned_zone_values(
        self, cells: list[str], zones: list[str]
    ) -> list[str] | None:
        """Return rate values aligned to zone columns (weight in column 0)."""
        if len(cells) < len(zones) + 1:
            return None
        values = cells[1 : 1 + len(zones)]
        if not any(values):
            return None
        parsed = 0
        for value in values:
            if not value:
                continue
            try:
                _parse_number(value)
                parsed += 1
            except ValueError:
                return None
        if parsed == 0:
            return None
        return values

    def _zip_zones(self, zones: list[str], values: list[str]) -> dict[str, float]:
        result: dict[str, float] = {}
        for zone, value in zip(zones, values):
            if not zone or not value:
                continue
            try:
                result[str(zone)] = _parse_number(value)
            except ValueError:
                continue
        return result

    def _is_per_pound_weight_row(self, text: str) -> bool:
        return bool(WEIGHT_RANGE_RE.match(text))

    def _parse_weight_range(self, text: str) -> dict[str, Any]:
        m = WEIGHT_RANGE_RE.match(text)
        if not m:
            return {"label": text, "min": None, "max": None}
        min_v = _parse_number(m.group(1))
        max_v = _parse_number(m.group(2))
        return {"label": text.strip(), "min": min_v, "max": max_v}

    def _parse_zone_chart(self, start_index: int) -> tuple[dict[str, Any], int]:
        title_row = self._rows[start_index]
        title = _normalize_zone_chart_title(
            next((c for c in title_row.cells if c), "Zone Chart")
        )

        i = start_index + 1
        service_codes: list[str] = []
        service_columns: dict[str, int] = {}
        countries: list[dict[str, str]] = []

        while i < len(self._rows):
            row = self._rows[i]
            texts = [c for c in row.cells if c]
            if not texts:
                i += 1
                continue

            if texts[0] == "Country/Territory":
                service_codes = []
                service_columns = {}
                for col_idx, cell_text in enumerate(row.cells):
                    if cell_text and cell_text != "Country/Territory":
                        service_codes.append(cell_text)
                        service_columns[cell_text] = col_idx
                i += 1
                continue

            if self._extract_service_header(row) or self._is_zone_chart_title(row):
                break
            if texts[0] in ("FedEx GroundNet Rates", "FedEx ExpressNet Rates"):
                break

            country = (row.cells[0] or "").strip()
            if not country or country.startswith('"'):
                i += 1
                continue

            zone_map = {
                code: row.cells[col]
                for code, col in service_columns.items()
                if col < len(row.cells) and row.cells[col]
            }
            if zone_map:
                countries.append({"country": country, "zones": zone_map})
            i += 1

        rate_services = []
        for code in service_codes:
            zone_to_countries: dict[str, list[str]] = {}
            for entry in countries:
                zone = entry["zones"].get(code)
                if not zone:
                    continue
                zone_to_countries.setdefault(zone, []).append(entry["country"])
            rate_services.append(
                {
                    "rate_name": code,
                    "zones": [
                        {"zone": zone, "countries": sorted(names)}
                        for zone, names in sorted(zone_to_countries.items())
                    ],
                }
            )

        return (
            {
                "name": title,
                "countries": countries,
                "rate_services": rate_services,
            },
            i,
        )


def parse_fedex_rates(path: str | Path) -> dict[str, Any]:
    return FedExRatesParser(path).parse()


def write_json(data: dict[str, Any], output_path: str | Path, indent: int = 2) -> None:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=indent, ensure_ascii=False)


def nest_rate_blocks(flat: dict[str, Any]) -> dict[str, Any]:
    """Group flat rate_blocks into a hierarchy keyed by section and service names."""
    root: dict[str, Any] = {}
    for block in flat.get("rate_blocks", []):
        node = root
        for level in block.get("hierarchy", []):
            node = node.setdefault(level, {})
        service = block["service"]
        svc_node = node.setdefault(service, {})
        model = block["pricing_model"]
        payload = {
            "currency": block.get("currency"),
            "subtype": block.get("subtype"),
            "customer_name_location": block.get("customer_name_location"),
            "account_number": block.get("account_number"),
            "effective_date_note": block.get("effective_date_note"),
            "location_applicability_note": block.get("location_applicability_note"),
            "applicable_locations": block.get("applicable_locations"),
            "zones": block["rates"]["zones"],
            "entries": block["rates"]["entries"],
        }
        svc_node[model] = payload
    return {
        "source_file": flat.get("source_file"),
        "document": flat.get("document"),
        "rates": root,
        "zone_charts": flat.get("zone_charts"),
    }


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert FedEx XML rate sheet (.xls) to structured JSON."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="input/Ratecard2026_InternationalBE ECOLAB EAN201177953.xls",
        help="Path to FedEx .xls (XML spreadsheet)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="output/fedex_rates.json",
        help="Output JSON path",
    )
    parser.add_argument(
        "--nested",
        action="store_true",
        help="Also write a hierarchy-keyed JSON (adds _nested suffix)",
    )
    args = parser.parse_args()
    data = parse_fedex_rates(args.input)
    write_json(data, args.output)
    if args.nested:
        nested_path = Path(args.output).with_stem(
            Path(args.output).stem + "_nested"
        )
        write_json(nest_rate_blocks(data), nested_path)
        print(f"Wrote {nested_path}")
    print(f"Wrote {args.output}")
    print(f"  rate blocks: {len(data['rate_blocks'])}")
    print(f"  zone charts: {len(data['zone_charts'])}")


if __name__ == "__main__":
    main()
