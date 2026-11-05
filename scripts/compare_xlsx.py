#!/usr/bin/env python3
"""Compare Excel sheet structure and literal/formula cells without modifying either file.

Formula cached results are intentionally ignored because Excel recalculation may be stale.
"""
import argparse
import json
import sys
from xlsx_core import cell_map, workbook_inventory


def compare(old, new, limit=30):
    left, right = cell_map(old), cell_map(new)
    removed_keys, added_keys = set(left) - set(right), set(right) - set(left)
    changed = [(key, left[key], right[key])
               for key in set(left) & set(right) if left[key] != right[key]]
    a, b = workbook_inventory(old), workbook_inventory(new)
    important = ("charts", "tables", "images", "macros", "external_links", "defined_names")
    changed_counts = {k: {"old": a[k], "new": b[k]} for k in important if a[k] != b[k]}
    by_sheet_a = {s["name"]: s for s in a["sheets"]}
    by_sheet_b = {s["name"]: s for s in b["sheets"]}
    sheet_change = {name: {"old": by_sheet_a.get(name), "new": by_sheet_b.get(name)}
                    for name in sorted(set(by_sheet_a) | set(by_sheet_b))
                    if by_sheet_a.get(name) != by_sheet_b.get(name)}
    return {
        "old_file": str(old), "new_file": str(new),
        "removed_cells": len(removed_keys), "added_cells": len(added_keys),
        "changed_cells": len(changed),
        "removed_examples": [str(k) for k in sorted(removed_keys)[:limit]],
        "added_examples": [str(k) for k in sorted(added_keys)[:limit]],
        "changed_examples": [
            {"sheet": k[0], "cell": k[1], "old": str(a0)[:120], "new": str(b0)[:120]}
            for k, a0, b0 in sorted(changed)[:limit]],
        "object_count_changes": changed_counts, "sheet_changes": sheet_change
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("old")
    parser.add_argument("new")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--max-details", type=int, default=30)
    parser.add_argument("--fail-on-difference", action="store_true")
    args = parser.parse_args(argv)
    try:
        result = compare(args.old, args.new, max(0, args.max_details))
    except Exception as exc:
        print("ERROR: " + type(exc).__name__ + ": " + str(exc), file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print("Removed cells: {removed_cells}; added: {added_cells}; changed: {changed_cells}".format(**result))
        for k, v in result["object_count_changes"].items():
            print("Object change:", k, v)
        for ex in result["changed_examples"]:
            print("Changed:", ex["sheet"], ex["cell"], ex["old"], "->", ex["new"])
        for name in result["sheet_changes"]:
            print("Changed sheet metadata:", name)
    differences = bool(result["removed_cells"] or result["added_cells"] or result["changed_cells"]
                       or result["object_count_changes"] or result["sheet_changes"])
    return int(args.fail_on_difference and differences)


if __name__ == "__main__":
    sys.exit(main())
