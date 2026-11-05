#!/usr/bin/env python3
"""Export GitHub issues to Excel-safe UTF-8 CSV, using only Python stdlib.

No GitHub token is required for public repositories. For private repositories,
set GITHUB_TOKEN in the process environment; never put tokens in config files.
Does not execute arbitrary URLs or run external programs.
"""
import argparse
import csv
import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen

REPO_RE = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
COLUMNS = ["number", "title", "state", "created_at", "updated_at",
           "assignees", "labels", "url"]
RETRY_CODES = {429, 500, 502, 503, 504}


def safe_text(value):
    """Keep data literal when CSV is opened in Excel (CSV injection defense)."""
    s = "" if value is None else str(value)
    return "'" + s if s.lstrip().startswith(("=", "+", "-", "@")) or s.startswith(("\t", "\r", "\n")) else s


def check_repo(name):
    if not REPO_RE.fullmatch(name) or ".." in name or name.startswith((".", "-")):
        raise ValueError("Repository must be in owner/name form with safe characters")
    return name


def request_json(url, token=None, opener=urlopen, pause=time.sleep):
    if urlparse(url).netloc != "api.github.com" or urlparse(url).scheme != "https":
        raise ValueError("GitHub API URL not allowed")
    headers = {"Accept": "application/vnd.github+json",
               "User-Agent": "ExcelSkills-issue-export/1.0",
               "X-GitHub-Api-Version": "2022-11-28"}
    if token:
        headers["Authorization"] = "Bearer " + token
    req = Request(url, headers=headers, method="GET")
    for attempt in range(4):
        try:
            with opener(req, timeout=30) as response:
                payload = json.loads(response.read().decode("utf-8"))
                link = response.headers.get("Link", "")
                return payload, link
        except HTTPError as exc:
            if exc.code not in RETRY_CODES or attempt == 3:
                raise RuntimeError("GitHub API error HTTP " + str(exc.code)) from exc
            pause(2 ** attempt)
        except URLError as exc:
            if attempt == 3:
                raise RuntimeError("GitHub API connection error") from exc
            pause(2 ** attempt)
    raise RuntimeError("GitHub API failed after retries")


def next_link(header):
    for item in header.split(","):
        match = re.search(r'<([^>]+)>;\s*rel="next"', item.strip())
        if match:
            url = match.group(1)
            parsed = urlparse(url)
            if parsed.scheme != "https" or parsed.netloc != "api.github.com":
                raise ValueError("Unexpected pagination host")
            return url
    return None


def fetch_issues(repo, state="open", token=None, max_pages=100, opener=urlopen):
    repo = check_repo(repo)
    if state not in {"open", "closed", "all"}:
        raise ValueError("state must be open, closed or all")
    if not 1 <= max_pages <= 1000:
        raise ValueError("max_pages must be 1 to 1000")
    owner, name = repo.split("/")
    url = ("https://api.github.com/repos/" + quote(owner) + "/" + quote(name)
           + "/issues?state=" + state + "&per_page=100")
    issues, seen = [], set()
    for _ in range(max_pages):
        payload, link = request_json(url, token=token, opener=opener)
        if not isinstance(payload, list):
            raise ValueError("Unexpected GitHub API response")
        for issue in payload:
            if "pull_request" in issue:
                continue
            number = issue.get("number")
            if not isinstance(number, int) or number in seen:
                continue
            seen.add(number)
            issues.append({
                "number": number,
                "title": safe_text(issue.get("title", "")),
                "state": safe_text(issue.get("state", "")),
                "created_at": safe_text(issue.get("created_at", "")),
                "updated_at": safe_text(issue.get("updated_at", "")),
                "assignees": safe_text(", ".join(a.get("login", "") for a in issue.get("assignees", []) if isinstance(a, dict))),
                "labels": safe_text(", ".join(a.get("name", "") for a in issue.get("labels", []) if isinstance(a, dict))),
                "url": safe_text(issue.get("html_url", "")),
            })
        url = next_link(link)
        if not url:
            return issues
    if url:
        raise RuntimeError("Pagination limit reached; export would be incomplete")
    return issues


def export_csv(destination, items):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".issues-", suffix=".csv", dir=destination.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8-sig", newline="") as output:
            writer = csv.DictWriter(output, fieldnames=COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            for item in items:
                writer.writerow({key: item.get(key, "") for key in COLUMNS})
        os.replace(tmp, destination)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return destination


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", required=True, help="owner/name")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--state", choices=("open", "closed", "all"), default="open")
    parser.add_argument("--max-pages", type=int, default=100)
    parser.add_argument("--apply", action="store_true", help="Actually retrieve and write CSV")
    args = parser.parse_args(argv)
    try:
        check_repo(args.repo)
        if not args.apply:
            print("DRY RUN: export", args.repo, "issues to", args.output)
            return 0
        items = fetch_issues(args.repo, args.state, os.environ.get("GITHUB_TOKEN"), args.max_pages)
        export_csv(args.output, items)
        print("Exported", len(items), "issues to", args.output)
        return 0
    except (ValueError, RuntimeError, OSError) as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
