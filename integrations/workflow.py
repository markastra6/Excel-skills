#!/usr/bin/env python3
"""Small allowlisted cross-app workflow runner; dry-run unless --apply.

Supported operations: local XLSX auditing, workbook inventory, GitHub issue
export to CSV and optional CSV import to Excel on Windows. This is NOT Composio
and does not copy Composio's cloud connections or credentials.
"""
import argparse
import glob
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "integrations"))
from audit_xlsx import audit
from xlsx_core import workbook_inventory
from github_issues import fetch_issues, export_csv, check_repo

ALLOWED = {"github_issues_to_csv", "audit_xlsx", "inventory_xlsx", "csv_to_xlsx"}


def checked_path(name, root=ROOT):
    """Paths are restricted to repo workspace, excluding traversal and symlink escapes."""
    if not isinstance(name, str) or not name.strip():
        raise ValueError("Missing path")
    p = (Path(root).resolve() / name).resolve()
    if not p.is_relative_to(Path(root).resolve()):
        raise ValueError("Paths must be inside the repository workspace")
    return p


def validate_plan(data, root=ROOT):
    if not isinstance(data, dict) or data.get("version") != 1:
        raise ValueError("Workflow version must be 1")
    steps = data.get("steps")
    if not isinstance(steps, list) or not 1 <= len(steps) <= 100:
        raise ValueError("Workflow needs 1-100 steps")
    plan = []
    for index, step in enumerate(steps, 1):
        if not isinstance(step, dict) or step.get("type") not in ALLOWED:
            raise ValueError("Unsupported step " + str(index))
        kind = step["type"]
        if kind == "github_issues_to_csv":
            repo = check_repo(step.get("repo", ""))
            state = step.get("state", "open")
            if state not in ("open", "closed", "all"):
                raise ValueError("Invalid issue state")
            output = checked_path(step.get("output"), root)
            if output.suffix.lower() != ".csv":
                raise ValueError("GitHub export output must be CSV")
            plan.append((kind, {"repo": repo, "state": state, "output": output}))
        elif kind in {"audit_xlsx", "inventory_xlsx"}:
            patterns = step.get("files")
            if not isinstance(patterns, list) or not patterns:
                raise ValueError("File patterns must be a non-empty list")
            matches = []
            for pattern in patterns:
                p = checked_path(pattern, root)
                for f in sorted(glob.glob(str(p), recursive=False)):
                    matched = checked_path(str(Path(f).resolve().relative_to(Path(root).resolve())), root)
                    if matched.is_file() and matched.suffix.lower() == ".xlsx":
                        matches.append(matched)
            matches = sorted(set(matches))
            if not matches and not step.get("allow_empty", False):
                raise ValueError("No XLSX files matched: " + str(patterns))
            plan.append((kind, {"files": matches}))
        else:
            source = checked_path(step.get("source"), root)
            destination = checked_path(step.get("destination"), root)
            if source.suffix.lower() != ".csv" or destination.suffix.lower() != ".xlsx":
                raise ValueError("CSV import needs .csv source and .xlsx destination")
            if source == destination:
                raise ValueError("Source and destination cannot match")
            plan.append((kind, {"source": source, "destination": destination}))
    return plan


def run_plan(plan, apply=False):
    for i, (kind, data) in enumerate(plan, 1):
        print(str(i) + ". " + kind + " " + str({k: str(v) for k, v in data.items()}))
        if not apply:
            continue
        if kind == "github_issues_to_csv":
            issues = fetch_issues(data["repo"], data["state"], os.environ.get("GITHUB_TOKEN"))
            export_csv(data["output"], issues)
            print("   exported issues:", len(issues))
        elif kind == "audit_xlsx":
            for file in data["files"]:
                info = audit(file)
                print("   " + ("PASS" if not info["errors"] else "FAIL") + " " + str(file))
                if info["errors"]:
                    raise RuntimeError("Workbook audit errors: " + "; ".join(info["errors"][:8]))
        elif kind == "inventory_xlsx":
            for file in data["files"]:
                item = workbook_inventory(file)
                print("   sheets=" + str(len(item["sheets"])) + " charts=" + str(item["charts"]))
        elif kind == "csv_to_xlsx":
            if os.name != "nt":
                raise RuntimeError("CSV-to-XLSX import requires Windows and desktop Excel")
            if not data["source"].is_file():
                raise FileNotFoundError(data["source"])
            cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File",
                   str(ROOT / "desktop" / "ExcelSkills.ps1"), "-Action", "ImportCsv",
                   "-Source", str(data["source"]), "-Destination", str(data["destination"])]
            subprocess.run(cmd, check=True)
    if not apply:
        print("Dry run only. Add --apply to execute.")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workflow", type=Path)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args(argv)
    try:
        config = json.loads(args.workflow.read_text(encoding="utf-8"))
        plan = validate_plan(config)
        run_plan(plan, args.apply)
        return 0
    except (ValueError, RuntimeError, OSError, json.JSONDecodeError,
            subprocess.CalledProcessError) as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
