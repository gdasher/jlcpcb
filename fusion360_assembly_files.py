#!/usr/bin/env python3
"""
Write Fusion 360-format assembly files (BOM CSV + front/back pick-and-place)
from a board dump made by fusion360_dump_board.py.

The output matches what Fusion's CAM "Export file" produces, so it can be fed
straight to fusion360_to_jlcpcb.py:

    <name>.csv               BOM: Qty, Value, Device, Package, Parts, Description, <attributes>
    PnP_<name>_front.csv     Designator, X, Y, Rotation, Value, Package (top side)
    PnP_<name>_back.csv      same, bottom side

Coordinates are the element origin, rounded the way Fusion's CAM output does
(internal units * 3.125e-6, half-up on the double); this reproduces Fusion's
files digit for digit.

Usage:
    python3 fusion360_assembly_files.py DUMP.json [--name NAME] [--out DIR]
"""

import argparse
import csv
import html
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

_genericDevicesets = ("R", "C", "L")


def _Natural(text: str) -> list:
  return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", text)]


def _Device(part: dict) -> str:
  """Fusion's BOM Device column: '<set>_<device>' for the generic R/C/L library, else concatenated."""
  deviceset, device = part.get("deviceset", ""), part.get("device", "")
  if deviceset in _genericDevicesets and device:
    return f"{deviceset}_{device}"
  return deviceset + device


def _Description(part: dict) -> str:
  """First line of the deviceset description with the HTML stripped, as Fusion's BOM shows it."""
  firstLine = (part.get("desc") or "").split("\n")[0]
  return html.unescape(re.sub(r"<[^>]*>", "", firstLine)).strip()


def _Mm(value: float) -> str:
  return str(Decimal(round(value * 320000) * 3.125e-6).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _BomAttributes(element: dict) -> dict:
  """Attributes as Fusion's BOM groups and prints them: no NAME, no empties, and VALUE
  only when it differs from the part value."""
  return {k: v for k, v in element["attrs"].items()
          if k != "NAME" and v and not (k == "VALUE" and v == element["value"])}


def WriteBom(parts: dict, elements: dict, path: Path) -> int:
  attrNames = sorted({k for e in elements.values() for k in e["attrs"] if k != "NAME"})
  groups = {}
  for ref, element in elements.items():
    if not element.get("populate", True):
      continue
    part = parts.get(ref, {})
    attrs = tuple(sorted(_BomAttributes(element).items()))
    key = (element["value"], _Device(part), element["pkg"], _Description(part), attrs)
    groups.setdefault(key, []).append(ref)

  with open(path, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Qty", "Value", "Device", "Package", "Parts", "Description"] + attrNames)
    for key in sorted(groups, key=lambda k: (k[0], k[1], k[2])):
      value, device, package, description, attrs = key
      refs = sorted(groups[key], key=_Natural)
      attrMap = dict(attrs)
      writer.writerow([len(refs), value, device, package, ", ".join(refs), description]
                      + [attrMap.get(k, "") for k in attrNames])
  return len(groups)


def WritePickAndPlace(elements: dict, path: Path, bottom: bool) -> int:
  count = 0
  with open(path, "w", newline="", encoding="utf-8") as f:
    for ref in sorted(elements):
      element = elements[ref]
      if bool(element["m"]) != bottom or not element.get("populate", True):
        continue
      f.write(f"{ref},{_Mm(element['x'])},{_Mm(element['y'])},{_Mm(element['a'])},"
              f"{element['value']},{element['pkg']}\n")
      count += 1
  return count


def ParseArgs() -> argparse.Namespace:
  parser = argparse.ArgumentParser(
    description="Write Fusion 360-format BOM/PnP files from a fusion360_dump_board.py dump.",
    formatter_class=argparse.RawDescriptionHelpFormatter,
    epilog=__doc__,
  )
  parser.add_argument("dump", metavar="DUMP.json", help="Board dump from fusion360_dump_board.py")
  parser.add_argument("--name", default="board", help="Board name used in the file names (default: board)")
  parser.add_argument("--out", default=".", help="Output directory (default: .)")
  return parser.parse_args()


def Main() -> None:
  args = ParseArgs()
  data = json.loads(Path(args.dump).read_text())
  out = Path(args.out)
  out.mkdir(parents=True, exist_ok=True)
  lines = WriteBom(data["parts"], data["els"], out / f"{args.name}.csv")
  front = WritePickAndPlace(data["els"], out / f"PnP_{args.name}_front.csv", bottom=False)
  back = WritePickAndPlace(data["els"], out / f"PnP_{args.name}_back.csv", bottom=True)
  print(f"BOM: {lines} lines; PnP: {front} front, {back} back -> {out}")


if __name__ == "__main__":
  Main()
