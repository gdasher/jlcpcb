#!/usr/bin/env python3
"""
Transform Autodesk Fusion 360 PCB assembly exports (BOM + CPL CSVs)
to the format required by JLCPCB (BOM + CPL xlsx).

Usage:
    python3 fusion360_to_jlcpcb.py BOM.csv FRONT_CPL.csv BACK_CPL.csv [options]

Arguments:
    BOM.csv         Fusion 360 BOM export (CSV)
    FRONT_CPL.csv   Fusion 360 Pick-and-Place front-side CSV
    BACK_CPL.csv    Fusion 360 Pick-and-Place back-side CSV

Options:
    --bom-out PATH  Output path for BOM xlsx (default: ./bom.xlsx)
    --cpl-out PATH  Output path for CPL xlsx (default: ./cpl.xlsx)
    -h, --help      Show this help message
"""

import argparse
import csv
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# BOM helpers
# ---------------------------------------------------------------------------

_GENERIC_DESCRIPTIONS = {
    "capacitor - generic",
    "resistor fixed - generic",
    "transistor",
    "diode rectifier - generic",
    "voltage regulator",
    "n-channel mosfet - generic",
    "diode",
}


def _is_generic(description: str) -> bool:
    return description.strip().lower() in _GENERIC_DESCRIPTIONS


def _choose_comment(value: str, device: str, description: str) -> str:
    """
    Heuristic for the JLCPCB 'Comment' field:
      1. If Value differs from Device, Value is a component value (e.g. '10k') — use it.
      2. Otherwise, if Description is non-empty and not generic, use Description.
      3. Fall back to Device (which equals Value in this case).
    """
    if value != device:
        return value
    if description and not _is_generic(description):
        return description
    return device


def _format_designators(parts: str) -> str:
    """Normalise designator list: strip whitespace around commas."""
    return ", ".join(d.strip() for d in parts.split(","))


def read_fusion_bom(path: Path) -> list[dict]:
    """Parse a Fusion 360 BOM CSV, returning rows as dicts."""
    for encoding in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            with open(path, newline="", encoding=encoding) as f:
                reader = csv.DictReader(f)
                return list(reader)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Could not decode {path} with any supported encoding")


def build_jlcpcb_bom_rows(bom_rows: list[dict]) -> list[tuple]:
    """Convert Fusion 360 BOM rows to JLCPCB BOM tuples (Comment, Designator, Footprint, Part#)."""
    result = []
    for row in bom_rows:
        value = (row.get("Value") or "").strip()
        device = (row.get("Device") or "").strip()
        description = (row.get("Description") or "").strip()
        parts = (row.get("Parts") or "").strip()
        package = (row.get("Package") or "").strip()

        if not parts:
            continue  # skip rows with no designators

        comment = _choose_comment(value, device, description)
        designator = _format_designators(parts)
        result.append((comment, designator, package, ""))
    return result


# ---------------------------------------------------------------------------
# CPL helpers
# ---------------------------------------------------------------------------

# Fusion 360 CPL has no header; columns are:
#   Designator, X (mm), Y (mm), Rotation, Value, Package
_CPL_COLS = ("Designator", "X", "Y", "Rotation", "Value", "Package")


def read_fusion_cpl(path: Path, layer: str) -> list[tuple]:
    """
    Parse a Fusion 360 CPL CSV.
    Returns list of (Designator, Mid X, Mid Y, Layer, Rotation) tuples.
    """
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for line in reader:
            if not line or not line[0].strip():
                continue
            # Skip any accidental header row (Fusion sometimes adds one)
            if line[0].strip().lower() in ("designator", "refdes", "ref"):
                continue
            if len(line) < 4:
                continue
            designator = line[0].strip()
            mid_x = f"{float(line[1].strip()):.4f}mm"
            mid_y = f"{float(line[2].strip()):.4f}mm"
            rotation = float(line[3].strip())
            rows.append((designator, mid_x, mid_y, layer, rotation))
    return rows


# ---------------------------------------------------------------------------
# xlsx writing
# ---------------------------------------------------------------------------

_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center")


def _write_header(ws, headers: list[str]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN


def _autofit_columns(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


def write_bom_xlsx(rows: list[tuple], out_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["Comment", "Designator", "Footprint", "JLCPCB Part #\uff08optional\uff09"]
    _write_header(ws, headers)

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    ws.freeze_panes = "A2"
    _autofit_columns(ws)
    wb.save(out_path)
    print(f"  BOM written -> {out_path}  ({len(rows)} components)")


def write_cpl_xlsx(rows: list[tuple], out_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["Designator", "Mid X", "Mid Y", "Layer", "Rotation"]
    _write_header(ws, headers)

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    ws.freeze_panes = "A2"
    _autofit_columns(ws)
    wb.save(out_path)
    print(f"  CPL written -> {out_path}  ({len(rows)} placements)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Fusion 360 BOM/CPL CSVs to JLCPCB xlsx format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("bom_csv", metavar="BOM.csv", help="Fusion 360 BOM CSV")
    parser.add_argument("front_cpl_csv", metavar="FRONT_CPL.csv", help="Fusion 360 front-side CPL CSV")
    parser.add_argument("back_cpl_csv", metavar="BACK_CPL.csv", help="Fusion 360 back-side CPL CSV")
    parser.add_argument("--bom-out", metavar="PATH", default="bom.xlsx",
                        help="Output path for BOM xlsx (default: ./bom.xlsx)")
    parser.add_argument("--cpl-out", metavar="PATH", default="cpl.xlsx",
                        help="Output path for CPL xlsx (default: ./cpl.xlsx)")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    bom_path = Path(args.bom_csv)
    front_path = Path(args.front_cpl_csv)
    back_path = Path(args.back_cpl_csv)
    bom_out = Path(args.bom_out)
    cpl_out = Path(args.cpl_out)

    for p in (bom_path, front_path, back_path):
        if not p.exists():
            print(f"Error: file not found: {p}", file=sys.stderr)
            sys.exit(1)

    print("Reading Fusion 360 exports...")
    bom_data = read_fusion_bom(bom_path)
    front_cpl = read_fusion_cpl(front_path, layer="Top")
    back_cpl = read_fusion_cpl(back_path, layer="Bottom")

    print(f"  BOM: {len(bom_data)} line items")
    print(f"  Front CPL: {len(front_cpl)} placements")
    print(f"  Back CPL:  {len(back_cpl)} placements")

    bom_rows = build_jlcpcb_bom_rows(bom_data)
    cpl_rows = front_cpl + back_cpl

    print("Writing JLCPCB files...")
    write_bom_xlsx(bom_rows, bom_out)
    write_cpl_xlsx(cpl_rows, cpl_out)

    print("Done.")


if __name__ == "__main__":
    main()
