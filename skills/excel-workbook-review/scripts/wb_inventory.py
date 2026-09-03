#!/usr/bin/env python3
"""Structural inventory of an Excel workbook. Pure stdlib - no openpyxl needed.

Usage: python3 tools/wb_inventory.py <workbook.xlsx|xlsm> [--json]

Reads the OOXML package directly and reports the things a reviewer needs to
know before touching anything: sheets and their visibility, formula density,
volatile and fragile functions, cached error values, external links, defined
names, protection, and whether a VBA project is present.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"}
FRAGILE = {"INDIRECT", "OFFSET", "VLOOKUP", "HLOOKUP", "GETPIVOTDATA"}
FUNC_RE = re.compile(r"([A-Z][A-Z0-9._]*)\s*\(")


def tag(el):
    return el.tag.rsplit("}", 1)[-1]


def find_all(root, name):
    return [e for e in root.iter() if tag(e) == name]


def attr(el, name):
    for k, v in el.attrib.items():
        if k.rsplit("}", 1)[-1] == name:
            return v
    return None


def sheet_targets(z):
    """Map sheet name -> part path, via workbook.xml + its rels."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    by_id = {attr(r, "Id"): attr(r, "Target") for r in find_all(rels, "Relationship")}
    out = []
    for s in find_all(wb, "sheet"):
        target = by_id.get(attr(s, "id"), "")
        if target and not target.startswith("/"):
            target = "xl/" + target.lstrip("./")
        out.append(
            {
                "name": attr(s, "name"),
                "state": attr(s, "state") or "visible",
                "part": target.lstrip("/"),
            }
        )
    names = [
        {
            "name": attr(n, "name"),
            "refers_to": (n.text or "").strip(),
            "hidden": (attr(n, "hidden") == "1"),
        }
        for n in find_all(wb, "definedName")
    ]
    return out, names


def scan_sheet(z, part):
    try:
        xml = z.read(part)
    except KeyError:
        return None
    root = ET.fromstring(xml)
    formulas = find_all(root, "f")
    funcs = Counter()
    ext_refs = 0
    for f in formulas:
        text = (f.text or "").upper()
        if "[" in text:
            ext_refs += 1
        for m in FUNC_RE.finditer(text):
            funcs[m.group(1)] += 1
    error_cells = []
    for c in find_all(root, "c"):
        if attr(c, "t") == "e":
            value = None
            for kid in c:
                if tag(kid) == "v":
                    value = kid.text
                    break
            error_cells.append({"ref": attr(c, "r"), "error": value})
    protection = bool(find_all(root, "sheetProtection"))
    dim = find_all(root, "dimension")
    return {
        "part": part,
        "used_range": attr(dim[0], "ref") if dim else None,
        "formula_count": len(formulas),
        "external_formula_refs": ext_refs,
        "volatile_functions": {k: v for k, v in funcs.items() if k in VOLATILE},
        "fragile_functions": {k: v for k, v in funcs.items() if k in FRAGILE},
        "top_functions": funcs.most_common(10),
        "error_cells": error_cells[:50],
        "error_cell_count": len(error_cells),
        "merged_cells": len(find_all(root, "mergeCell")),
        "conditional_formats": len(find_all(root, "conditionalFormatting")),
        "data_validations": len(find_all(root, "dataValidation")),
        "protected": protection,
    }


def main(argv):
    if len(argv) < 2:
        print(__doc__.strip())
        return 2
    path = argv[1]
    with zipfile.ZipFile(path) as z:
        parts = set(z.namelist())
        sheets, names = sheet_targets(z)
        report = {
            "workbook": path,
            "has_vba_project": "xl/vbaProject.bin" in parts,
            "has_external_links": any(p.startswith("xl/externalLinks/") for p in parts),
            "external_link_parts": sorted(
                p for p in parts if p.startswith("xl/externalLinks/") and p.endswith(".xml")
            ),
            "has_pivot_caches": any(p.startswith("xl/pivotCache/") for p in parts),
            "query_tables": sorted(p for p in parts if "queryTable" in p or "connections" in p),
            "defined_names": names,
            "sheets": [],
        }
        for s in sheets:
            info = scan_sheet(z, s["part"]) or {}
            info.update({"name": s["name"], "state": s["state"]})
            report["sheets"].append(info)
    report["totals"] = {
        "sheets": len(report["sheets"]),
        "hidden_sheets": sum(1 for s in report["sheets"] if s.get("state") != "visible"),
        "formulas": sum(s.get("formula_count", 0) for s in report["sheets"]),
        "error_cells": sum(s.get("error_cell_count", 0) for s in report["sheets"]),
    }
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
