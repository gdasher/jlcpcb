"""
Fusion 360 script: dump the open electronics design's board elements and
schematic parts to JSON for fusion360_assembly_files.py.

Run it inside Fusion (Utilities > Scripts and Add-Ins, or through the Fusion
MCP) with the electronics design open. It reads the first open Board and
Schematic products and writes OUTPUT_PATH.

Fusion's CAM "Export file" writes the BOM and pick-and-place files itself;
this dump is the fallback when that is not available (e.g. Electron.mfgexport
from a script skips the assembly outputs).
"""

import json
import os

import adsk.core
import adsk.electron

OUTPUT_PATH = os.path.join(os.path.expanduser("~"), "fusion_board_dump.json")
_unitsPerMm = 320000  # Fusion electronics internal units


def _Name(obj) -> str:
  return obj if isinstance(obj, str) else getattr(obj, "name", "")


def _FindProduct(app, objectType: str):
  for i in range(app.documents.count):
    for product in app.documents.item(i).products:
      if product.objectType == objectType:
        return product
  return None


def _Parts(schematic) -> dict:
  parts = {}
  for i in range(schematic.parts.count):
    part = schematic.parts.item(i)
    entry = {}
    for key, getter in (("deviceset", lambda: _Name(part.deviceset)),
                        ("device", lambda: _Name(part.device)),
                        ("desc", lambda: part.deviceset.description)):
      try:
        entry[key] = getter()
      except Exception:
        pass
    parts[part.name] = entry
  return parts


def _Elements(board) -> dict:
  elements = {}
  for i in range(board.elements.count):
    element = board.elements.item(i)
    try:
      attrs = element.attributes
      attributes = {attrs.item(k).name: attrs.item(k).value for k in range(attrs.count)}
    except Exception:
      attributes = {}
    try:
      populate = element.populate
    except Exception:
      populate = True
    elements[element.name] = {
      "x": element.x / _unitsPerMm, "y": element.y / _unitsPerMm,
      "a": element.angle, "m": element.mirror,
      "value": element.value, "pkg": element.package.name,
      "attrs": attributes, "populate": populate,
    }
  return elements


def run(context):
  app = adsk.core.Application.get()
  board = _FindProduct(app, "adsk::electron::Board")
  schematic = _FindProduct(app, "adsk::electron::Schematic")
  if board is None:
    raise RuntimeError("No open electronics Board")
  data = {"parts": _Parts(schematic) if schematic else {}, "els": _Elements(board)}
  with open(OUTPUT_PATH, "w") as f:
    json.dump(data, f, indent=1)
  print(f"Wrote {OUTPUT_PATH}: {len(data['els'])} elements, {len(data['parts'])} parts")
