# PR SLA Monthly Report

Auto-generates a PowerPoint deck listing every team member's **completed PRs** in a given month across configured Azure DevOps repos, with:

- ADO work item IDs (hyperlinked to the work item)
- PR URL (hyperlinked)
- Last commit date on the PR (`Dev Complete/PR Submitted Date`)
- PR completed date (`Approved Date`)
- Business days between the two (`Days To Process`)
- Empty `Comments` column for manual notes

One slide per repo. Layout mirrors the existing `Template` sheet in `config/team-config.xlsx`.

---

## Configuration

Edit `config/team-config.xlsx`:

| Sheet | Purpose |
|-------|---------|
| **Repos** | One ADO repo URL per row under header `ADO Repo`. Supports both `dev.azure.com/<org>/<project>/_git/<repo>` and `<org>.visualstudio.com/<project>/_git/<repo>`. |
| **Team** | `Name` + `V-id` (UPN/email) per team member. The script resolves the UPN to an ADO identity GUID and queries PRs where that user is the **creator**. |
| **Template** | Reference only — shows the target table layout. Ignored at runtime. |

---

## Local run

```bash
cd ~/Personal/pr-sla-report
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

export ADO_PAT="<your PAT — see below>"

# Default: previous calendar month
python src/generate_report.py

# Specific month
python src/generate_report.py --month 2026-05

# Custom date range
python src/generate_report.py --from 2026-04-01 --to 2026-04-30
```

Output lands in `output/PR_SLA_<YYYY-MM>.pptx`.

---

## Personal Access Token (PAT)

Create at [https://dev.azure.com/MicrosoftIT/_usersSettings/tokens](https://dev.azure.com/MicrosoftIT/_usersSettings/tokens):

| Scope | Level |
|-------|-------|
| **Code** | Read |
| **Work Items** | Read |
| **Identity** | Read |

Set `Organization = All accessible organizations` if your repos span multiple (e.g. `MicrosoftIT` + `microsoftit`).

Expire in 90 days, then rotate.

---

## GitHub Actions monthly automation

1. Push this repo to GitHub (public or private).
2. Repo → **Settings → Secrets and variables → Actions → New repository secret**:
   - Name: `ADO_PAT`
   - Value: your PAT
3. The workflow at [.github/workflows/monthly-report.yml](.github/workflows/monthly-report.yml) runs on the **1st of every month at 02:00 UTC** and uploads the generated `.pptx` as a workflow artifact (`pr-sla-report`, retained 90 days).
4. To regenerate any past month: **Actions → Monthly PR SLA Report → Run workflow**, optionally typing a `YYYY-MM` override.

To get the PPT delivered (email / Teams / Slack) instead of just stored as an artifact, add a step to the workflow — see commented examples in the file.

---

## How the math works

- **Last commit date**: pulled from `GET /pullrequests/{id}/commits`, picking the most recent `committer.date` (falls back to `author.date`).
- **Completed date**: the PR's `closedDate` returned by ADO when `status=completed`.
- **Business days**: `numpy.busday_count(last_commit, completed)` — exclusive of start, inclusive of end. Same-day = `0`. **Weekends are excluded; US holidays are NOT** (add a `holidays=` arg in `business_days_between` if you want them).

---

## Project layout

```
pr-sla-report/
├── .github/workflows/monthly-report.yml   # cron + manual dispatch
├── config/
│   └── team-config.xlsx                   # repos + team + template
├── output/                                # generated decks (gitignored)
├── src/
│   └── generate_report.py                 # main script
├── requirements.txt
└── README.md
```
