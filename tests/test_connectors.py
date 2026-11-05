"""Tests use mock HTTP responses; no live network or Excel installation."""
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "integrations"))
from github_issues import check_repo, fetch_issues, export_csv, safe_text, next_link
from workflow import validate_plan, run_plan


class Response:
    def __init__(self, rows, link=""):
        self.rows = rows
        self.headers = {"Link": link}

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return json.dumps(self.rows).encode("utf-8")


class ConnectorTests(unittest.TestCase):
    def test_csv_formula_injection(self):
        for prefix in ("=", "+", "-", "@", "\t"):
            self.assertTrue(safe_text(prefix + "1").startswith("'"))
        self.assertEqual(safe_text("Ordinary"), "Ordinary")

    def test_owner_repo_validation(self):
        for invalid in ("../../secrets", "owner/", "https://github.com/a/b", "owner/.."):
            with self.assertRaises(ValueError):
                check_repo(invalid)

    def test_pagination_filters_pull_requests(self):
        links = iter([
            Response([{"number": 1, "title": "=1+1", "state": "open"}],
                     '<https://api.github.com/repos/a/b/issues?state=open&page=2>; rel="next"'),
            Response([{"number": 2, "title": "PR", "pull_request": {}},
                      {"number": 3, "title": "Issue", "state": "open"}])
        ])
        seen = []
        def open_fake(req, timeout):
            seen.append(req.full_url)
            return next(links)
        rows = fetch_issues("a/b", opener=open_fake)
        self.assertEqual([x["number"] for x in rows], [1, 3])
        self.assertEqual(rows[0]["title"], "'=1+1")
        self.assertEqual(len(seen), 2)

    def test_reject_malicious_pagination(self):
        with self.assertRaises(ValueError):
            next_link('<https://evil.test/steal>; rel="next"')

    def test_write_utf8_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "issues.csv"
            export_csv(path, [{"number": 1, "title": "தமிழ்"}])
            raw = path.read_bytes()
            self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))
            self.assertIn("தமிழ்", raw.decode("utf-8-sig"))

    def test_workflow_dry_run_has_no_side_effects(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            plan = validate_plan({"version": 1, "steps": [
                {"type": "github_issues_to_csv", "repo": "owner/repo",
                 "output": "reports/test.csv"}]}, root)
            with patch("workflow.fetch_issues", side_effect=AssertionError("Network attempted")):
                run_plan(plan, apply=False)
            self.assertFalse((root / "reports/test.csv").exists())

    def test_workflow_blocks_parent_traversal(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                validate_plan({"version": 1, "steps": [
                    {"type": "github_issues_to_csv", "repo": "owner/repo",
                     "output": "../secret.csv"}]}, Path(tmp))

    def test_workflow_unknown_action(self):
        with self.assertRaises(ValueError):
            validate_plan({"version": 1, "steps": [{"type": "shell", "cmd": "echo"}]})


if __name__ == "__main__":
    unittest.main()
