---
name: local-app-integrations
description: Run safe local Excel workflows and GitHub data exports without relying on Composio's cloud integration service.
---

# Local App Integrations

This is a small integration **alternative for supported workflows**, not a clone of Composio. It does not reproduce hosted authentication, app access, event delivery, hundreds of integrations or cloud automation.

## When to use
- Exporting GitHub issue data to Excel-compatible CSV.
- Running a repeatable sequence of Excel audits and local actions.
- Importing CSV into desktop Excel on Windows.
- Automating a GitHub repo's issue export using GitHub Actions.

## Commands
```bash
python integrations/github_issues.py --repo markastra6/Excel-skills --output reports/issues.csv
python integrations/github_issues.py --repo markastra6/Excel-skills --output reports/issues.csv --apply
python integrations/workflow.py examples/github_issues_workflow.json
python integrations/workflow.py examples/github_issues_workflow.json --apply
```

## Safety
- Dry-run first; no shell commands or arbitrary API hosts are allowed.
- Set `GITHUB_TOKEN` as an environment variable for private GitHub access. Never commit credentials.
- The default export excludes pull requests, omits issue bodies and protects against spreadsheet formula injection.
- The repository is public; exported CSV files belong in the gitignored `reports/` directory.
- External services beyond GitHub need their own reviewed adapters and authorization.
