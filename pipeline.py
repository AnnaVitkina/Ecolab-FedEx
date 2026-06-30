"""
End-to-end FedEx rate sheet processing pipeline.

Designed for Google Colab:

    import os
    exec(open('/content/Ecolab-FedEx/pipeline.py').read())

Parses the input .xls in memory and writes deliverable files to the output folder:
  - matrix XLSX (rates)
  - zone chart text exports
"""

from __future__ import annotations

import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path

CODE_DIR = Path("/content/Ecolab-FedEx")
DRIVE_BASE = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Ecolab/RMT_FedEx"
)
try:
    _LOCAL_BASE = Path(__file__).resolve().parent
except NameError:
    _LOCAL_BASE = CODE_DIR
_IN_COLAB = Path("/content").is_dir()
_ACTIVE_BASE = DRIVE_BASE if _IN_COLAB else _LOCAL_BASE

DEFAULT_INPUT_DIR = _ACTIVE_BASE / "input"
DEFAULT_PROCESSING_DIR = _ACTIVE_BASE / "processing"
DEFAULT_OUTPUT_DIR = _ACTIVE_BASE / "output"

_CODE_DIR = CODE_DIR if CODE_DIR.is_dir() else _LOCAL_BASE
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

from fedex_rates_parser import parse_fedex_rates
from json_to_xlsx import data_to_xlsx
from zone_charts_txt import (
    write_postal_zones_data,
    write_zone_charts_data,
    write_zone_charts_iso_data,
)


@dataclass
class PipelineOutputs:
    input_file: Path
    processing_dir: Path
    output_dir: Path
    xlsx_path: Path | None = None
    zones_txt_path: Path | None = None
    zones_iso_txt_path: Path | None = None
    postal_zones_txt_path: Path | None = None
    rate_blocks: int = 0
    zone_charts: int = 0
    output_files: list[Path] = field(default_factory=list)


def ensure_pipeline_dirs() -> None:
    for path in (DEFAULT_INPUT_DIR, DEFAULT_PROCESSING_DIR, DEFAULT_OUTPUT_DIR):
        path.mkdir(parents=True, exist_ok=True)


def discover_input_files(input_dir: Path = DEFAULT_INPUT_DIR) -> list[Path]:
    if not input_dir.is_dir():
        return []
    return sorted(input_dir.glob("*.xls"))


def select_input_file(input_dir: Path = DEFAULT_INPUT_DIR) -> Path:
    """Prompt the user to choose a FedEx .xls file from input/ or enter a path."""
    files = discover_input_files(input_dir)

    print("\nFedEx Rate Sheet Pipeline")
    print("=" * 40)
    print(f"Input:      {DEFAULT_INPUT_DIR}")
    print(f"Processing: {DEFAULT_PROCESSING_DIR}")
    print(f"Output:     {DEFAULT_OUTPUT_DIR}")

    if files:
        print(f"\nFiles in {input_dir}:\n")
        for index, path in enumerate(files, start=1):
            print(f"  {index}. {path.name}")
        print("  0. Enter a custom file path")
        print()

        while True:
            raw = input(f"Select file to process [1-{len(files)}, default 1]: ").strip()
            if not raw:
                return files[0]
            if raw == "0":
                break
            if raw.isdigit():
                choice = int(raw)
                if 1 <= choice <= len(files):
                    return files[choice - 1]
            print("Invalid selection. Try again.")
    else:
        print(f"\nNo .xls files found in {input_dir}.")
        print("Enter the full path to a FedEx rate sheet (.xls).\n")

    while True:
        raw = input("Path to .xls file: ").strip().strip('"')
        if not raw:
            print("A file path is required.")
            continue
        path = Path(raw)
        if path.is_file():
            return path
        print(f"File not found: {path}")


def processing_dir_for(input_path: Path, base_dir: Path = DEFAULT_PROCESSING_DIR) -> Path:
    return base_dir / input_path.stem


def output_dir_for(input_path: Path, base_dir: Path = DEFAULT_OUTPUT_DIR) -> Path:
    return base_dir / input_path.stem


def stage_input_file(input_path: Path, processing_dir: Path) -> Path:
    """Copy the source rate sheet into the processing workspace."""
    processing_dir.mkdir(parents=True, exist_ok=True)
    staged_path = processing_dir / input_path.name
    shutil.copy2(input_path, staged_path)
    return staged_path


def remove_stale_outputs(work_dir: Path) -> None:
    """Remove artifacts left from older pipeline runs."""
    if not work_dir.is_dir():
        return
    for path in work_dir.glob("*.json"):
        path.unlink(missing_ok=True)


def run_pipeline(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    processing_dir: str | Path | None = None,
    *,
    write_xlsx: bool = True,
    write_zones: bool = True,
) -> PipelineOutputs:
    ensure_pipeline_dirs()

    input_path = Path(input_path)
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    work_dir = Path(processing_dir) if processing_dir else processing_dir_for(input_path)
    out_dir = Path(output_dir) if output_dir else output_dir_for(input_path)
    work_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    remove_stale_outputs(work_dir)
    remove_stale_outputs(out_dir)

    staged_input = stage_input_file(input_path, work_dir)

    stem = input_path.stem
    xlsx_path = out_dir / f"{stem}_matrix.xlsx"
    zones_txt_path = out_dir / f"{stem}_zones.txt"
    zones_iso_txt_path = out_dir / f"{stem}_zones_iso.txt"
    postal_zones_txt_path = out_dir / f"{stem}_postal_zones.txt"

    print(f"\nProcessing: {input_path.name}")
    print(f"Staged in:  {work_dir}")
    print(f"Output dir: {out_dir}\n")

    print("[1/4] Parsing rate sheet...")
    data = parse_fedex_rates(staged_input)
    print(f"      {len(data['rate_blocks'])} rate blocks, {len(data['zone_charts'])} zone charts")

    output_files: list[Path] = []

    if write_xlsx:
        print("[2/4] Exporting matrix XLSX...")
        data_to_xlsx(data, xlsx_path)
        output_files.append(xlsx_path)
        print(f"      -> {xlsx_path.name}")
    else:
        print("[2/4] Skipping matrix XLSX")
        xlsx_path = None

    if write_zones:
        print("[3/4] Exporting zone charts...")
        write_zone_charts_data(data, zones_txt_path, include_iso_duplicate=True)
        write_zone_charts_iso_data(data, zones_iso_txt_path)
        output_files.extend([zones_txt_path, zones_iso_txt_path])
        print(f"      -> {zones_txt_path.name}")
        print(f"      -> {zones_iso_txt_path.name}")
        print("[4/4] Exporting postal code zones...")
        write_postal_zones_data(data, postal_zones_txt_path)
        output_files.append(postal_zones_txt_path)
        print(f"      -> {postal_zones_txt_path.name}")
    else:
        print("[3/4] Skipping zone chart export")
        print("[4/4] Skipping postal zone export")
        zones_txt_path = None
        zones_iso_txt_path = None
        postal_zones_txt_path = None

    return PipelineOutputs(
        input_file=input_path,
        processing_dir=work_dir,
        output_dir=out_dir,
        xlsx_path=xlsx_path,
        zones_txt_path=zones_txt_path,
        zones_iso_txt_path=zones_iso_txt_path,
        postal_zones_txt_path=postal_zones_txt_path,
        rate_blocks=len(data["rate_blocks"]),
        zone_charts=len(data["zone_charts"]),
        output_files=output_files,
    )


def _default_cli_argv() -> list[str]:
    if _IN_COLAB:
        return []
    try:
        from IPython import get_ipython

        shell = get_ipython()
        if shell is not None and shell.__class__.__name__ in (
            "ZMQInteractiveShell",
            "Shell",
        ):
            return []
    except ImportError:
        pass
    return sys.argv[1:]


def print_summary(result: PipelineOutputs) -> None:
    print("\n" + "=" * 40)
    print("Done!")
    print(f"  Source:      {result.input_file.name}")
    print(f"  Processing:  {result.processing_dir}")
    print(f"  Rate blocks: {result.rate_blocks}")
    print(f"  Zone charts: {result.zone_charts}")
    print(f"\nOutputs in {result.output_dir}:")
    for path in result.output_files:
        print(f"  - {path.name}")


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Run the full FedEx rate sheet pipeline (XLSX + zone text exports)."
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="Path to FedEx .xls file (omit to choose interactively from the input folder)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR}/<input-stem>/)",
    )
    parser.add_argument(
        "--processing-dir",
        help=f"Processing workspace (default: {DEFAULT_PROCESSING_DIR}/<input-stem>/)",
    )
    parser.add_argument(
        "--input-dir",
        default=str(DEFAULT_INPUT_DIR),
        help="Directory to scan when selecting a file interactively",
    )
    parser.add_argument(
        "--no-xlsx",
        action="store_true",
        help="Skip matrix XLSX output",
    )
    parser.add_argument(
        "--no-zones",
        action="store_true",
        help="Skip zone chart text export",
    )
    args = parser.parse_args(argv if argv is not None else _default_cli_argv())

    try:
        input_path = Path(args.input) if args.input else select_input_file(Path(args.input_dir))
        result = run_pipeline(
            input_path,
            args.output_dir,
            args.processing_dir,
            write_xlsx=not args.no_xlsx,
            write_zones=not args.no_zones,
        )
        print_summary(result)
        return 0
    except (FileNotFoundError, KeyboardInterrupt) as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"\nPipeline failed: {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
