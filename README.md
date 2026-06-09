# PR SLA Monthly Report

Auto-generate a PowerPoint deck listing every team member's **completed pull requests** across multiple Azure DevOps repos for a chosen month, with computed SLA metrics (last commit date, last approval date, business days between them) and hyperlinks to every PR and linked work item.

One slide per repo. Layout mirrors the `Template` sheet in your Excel config (teal → blue gradient header, 6 columns).

---

## Table of Contents

- [What it does](#what-it-does)
- [Features](#features)
- [Output columns — full specification](#output-columns--full-specification)
- [Prerequisites — software](#prerequisites--software)
- [First-time setup on any machine](#first-time-setup-on-any-machine)
- [Configuring repos + team (`team-config.xlsx`)](#configuring-repos--team-team-configxlsx)
- [Personal Access Token (PAT) — required scopes](#personal-access-token-pat--required-scopes)
- [Running the script locally](#running-the-script-locally)
- [Sorting options](#sorting-options)
- [Automating with GitHub Actions (recommended)](#automating-with-github-actions-recommended)
- [Alternative: macOS launchd (local cron)](#alternative-macos-launchd-local-cron)
- [How each value is computed](#how-each-value-is-computed)
- [Project layout](#project-layout)
- [Troubleshooting](#troubleshooting)
- [Updating the script](#updating-the-script)
- [FAQ](#faq)

---

## What it does

Every month it asks Azure DevOps:

> "For each person on my team, which PRs did they author that were **completed** during this month, in each of these repos?"

For every such PR it pulls:

1. The **last commit** that landed on the PR branch → "Dev Complete / PR Submitted Date".
2. The timestamp of the **last reviewer-approval vote** on the PR → "Approved Date".
3. The IDs of every **linked work item** (ADO tasks/bugs/user stories).

Then it computes **business days** between `Dev Complete` and `Approved` (weekends excluded; same-day = `0`), assembles one slide per repo with a hyperlinked table, and writes a `.pptx`.

You run it once a month (or let GitHub Actions run it automatically on the 1st).

---

## Features

| Feature | Notes |
|---|---|
| **Multi-repo** | Add as many repos as you want — one slide per repo. Supports both `dev.azure.com/<org>/<project>/_git/<repo>` and legacy `<org>.visualstudio.com/<project>/_git/<repo>` URLs. |
| **Per-person filter** | Configurable team list in Excel (Name + UPN). Resolved to ADO identity GUID per org and cached. |
| **Real approval timestamps** | Reads PR threads and picks the latest `VoteUpdate` with vote ≥ 5 (Approved or Approved-with-suggestions). Not just "PR closed date". |
| **Business-day math** | `numpy.busday_count`; weekends excluded; US holidays NOT excluded by default (configurable in code). |
| **Hyperlinked output** | Every ADO work item ID is a click-through link to the work item; the PR URL is a click-through link to the PR. |
| **Multiple ADO IDs per PR** | Rendered space-separated in one cell (matches your existing template). |
| **Configurable sort** | `--sort-by approved` (default), `dev-complete`, `days`, or `author`; `--desc` flips order. |
| **Three date selection modes** | Default: previous calendar month. Or `--month YYYY-MM`. Or custom `--from / --to`. |
| **Throttle-aware** | ADO `429` and `5xx` responses are retried with exponential backoff. |
| **CI-ready** | Included GitHub Actions workflow runs monthly and uploads the deck as an artifact for 90 days. |

---

## Output columns — full specification

| # | Column header (in PPT) | Source | Format | Hyperlink? |
|---|---|---|---|---|
| 1 | **ADO Link** | All work-item IDs linked to the PR via the ADO `pullRequests/{id}/workitems` endpoint. | Space-separated list of numeric IDs (e.g. `15717821  15732018  15706485`). | ✅ Each ID links to `https://dev.azure.com/<org>/<project>/_workitems/edit/<id>`. |
| 2 | **PR** | The PR's canonical web URL. | Plain URL text. | ✅ The whole cell is a link to `https://dev.azure.com/<org>/<project>/_git/<repo>/pullrequest/<id>`. |
| 3 | **Dev Complete / PR Submitted Date** | The **publish date of the latest commit** on the PR source branch (`committer.date`, falling back to `author.date`). | `MM/DD/YYYY` | — |
| 4 | **Approved Date** | The **publish date of the latest reviewer approval vote** on the PR (vote == `10` Approved, or vote == `5` Approved-with-suggestions). | `MM/DD/YYYY` | — |
| 5 | **Days To Process** | `numpy.busday_count(Dev Complete, Approved)` — **weekdays only**; same-day = `0`; never negative. | Integer, centered. | — |
| 6 | **Comments** | Always empty in generated output — reserved for you to type your notes in PowerPoint. | — | — |

Slide title is rendered in red, format: `<repo-name> | PR SLA (<Month YYYY>)`.

---

## Prerequisites — software

You need **two things** on the machine that will generate the deck:

| Tool | Minimum version | Purpose | Install |
|---|---|---|---|
| **Python** | 3.10+ (3.12 recommended) | Runs the script | macOS: `brew install python@3.12` <br> Windows: [python.org installer](https://www.python.org/downloads/) <br> Linux: `apt install python3.12 python3.12-venv` |
| **Git** | any | Clone this repo + version control | macOS: `brew install git` <br> Windows: [git-scm.com](https://git-scm.com/) <br> Linux: `apt install git` |

That's it — every other dependency is pulled from `requirements.txt` into an isolated virtual environment.

**No PowerPoint installation needed.** The `.pptx` file is generated standalone by the `python-pptx` library; open the result in PowerPoint, Keynote, Google Slides, or LibreOffice.

---

## First-time setup on any machine

```bash
# 1. Clone the repo
git clone <your-fork-url> ~/Personal/pr-sla-report
cd ~/Personal/pr-sla-report

# 2. Create an isolated Python environment
python3 -m venv .venv
source .venv/bin/activate          # macOS/Linux
# .venv\Scripts\activate           # Windows PowerShell

# 3. Install dependencies (~30 MB)
pip install -r requirements.txt

# 4. Set your ADO PAT (see next section for how to create one)
export ADO_PAT="<paste-your-token>"
# Persist it: append `export ADO_PAT=...` to ~/.zshrc or ~/.bashrc

# 5. Run a test for last month
python src/generate_report.py

# 6. Open the result
open output/PR_SLA_*.pptx          # macOS
# start output\PR_SLA_*.pptx       # Windows
# xdg-open output/PR_SLA_*.pptx    # Linux
```

---

## Configuring repos + team (`team-config.xlsx`)

Open [config/team-config.xlsx](config/team-config.xlsx) — it has **three sheets**:

### Sheet 1: `Repos`

One ADO repo URL per row under the header `ADO Repo`. Any of these URL shapes work:

```
https://dev.azure.com/<org>/<project>/_git/<repo>
https://<org>.visualstudio.com/<project>/_git/<repo>
```

Currently configured:
- `MicrosoftIT/OneITVSO/ConsumerStore-WebComponents`
- `microsoftit/OneITVSO/CE-SITES-aemfnd-MT-AEM`

Add or remove rows to change which repos appear as slides.

### Sheet 2: `Team`

Two columns: `Name` (display only) and `V-id` (must be the person's UPN / Microsoft login email, e.g. `v-aklagad@microsoft.com`).

The script resolves each UPN to an ADO identity GUID once per run (and caches it). Only PRs where this identity is the **creator** are included.

### Sheet 3: `Template`

Reference only — shows the target table shape and screenshot the script reproduces. Ignored at runtime; safe to edit or delete.

---

## Personal Access Token (PAT) — required scopes

Create at **https://dev.azure.com/MicrosoftIT/_usersSettings/tokens** → **+ New Token**.

| Field | Value |
|---|---|
| Name | `pr-sla-report` |
| Organization | **All accessible organizations** (your 2 repos live under both `MicrosoftIT` and `microsoftit`) |
| Expiration | 90 days max (then rotate) |
| Scopes | **Custom defined** — uncheck everything, then check only: |

| Scope group | Permission | Why |
|---|---|---|
| **Code** | Read | List + read completed PRs, their commits, and their threads |
| **Work Items** | Read | Fetch the work-item IDs linked to each PR |
| **Identity** | Read | Resolve each `v-xxx@microsoft.com` UPN to an ADO identity GUID |

Click **Create**. **Copy the token immediately** — ADO never shows it again. If you lose it, regenerate.

---

## Running the script locally

```bash
# Previous calendar month (default)
python src/generate_report.py

# A specific month
python src/generate_report.py --month 2026-05

# Arbitrary date range
python src/generate_report.py --from 2026-04-01 --to 2026-04-30

# Choose a sort
python src/generate_report.py --sort-by dev-complete
python src/generate_report.py --sort-by approved --desc
python src/generate_report.py --sort-by days
python src/generate_report.py --sort-by author

# Custom output path
python src/generate_report.py --output /tmp/may.pptx

# Quieter logging
LOG_LEVEL=WARNING python src/generate_report.py

# Help
python src/generate_report.py --help
```

The default output path is `output/PR_SLA_<YYYY-MM>.pptx` (or `<from>_<to>.pptx` for custom ranges).

---

## Sorting options

| `--sort-by` | What gets used as the key |
|---|---|
| `approved` *(default)* | Approved Date → Dev Complete Date → PR ID |
| `dev-complete` | Dev Complete Date → Approved Date → PR ID |
| `days` | Days To Process → Approved Date → PR ID |
| `author` | Author name (case-insensitive) → Dev Complete Date → PR ID |

Add `--desc` to reverse (e.g. show slowest PRs first with `--sort-by days --desc`).

---

## Automating with GitHub Actions (recommended)

The repo ships with [.github/workflows/monthly-report.yml](.github/workflows/monthly-report.yml) which:

- Runs on a cron schedule: **1st of every month at 02:00 UTC** (≈ 7:30 AM IST, 7 PM PT the day before).
- Can also be triggered manually with an optional `YYYY-MM` override.
- Uploads the generated `.pptx` as a workflow **artifact** (retained 90 days).

### One-time setup

```bash
# Push the repo to GitHub (private recommended)
gh repo create pr-sla-report --private --source=. --push
```

Then in the GitHub web UI:

1. **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `ADO_PAT`
   - Value: the PAT you created above
2. **Actions** tab → enable workflows if prompted.

### Each month

You don't have to do anything — the deck is generated automatically and is downloadable from:

> **Actions → Monthly PR SLA Report → latest run → Artifacts → `pr-sla-report`**

To re-run for any past month: **Run workflow → Override month: `2026-05`**.

---

## Alternative: macOS launchd (local cron)

If you'd rather run it on your own laptop instead of GitHub Actions:

```bash
cat > ~/Library/LaunchAgents/com.amandalapu.pr-sla.plist <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>           <string>com.amandalapu.pr-sla</string>
  <key>ProgramArguments</key>
  <array>
    <string>/bin/zsh</string>
    <string>-c</string>
    <string>cd $HOME/Personal/pr-sla-report && .venv/bin/python src/generate_report.py</string>
  </array>
  <key>EnvironmentVariables</key>
  <dict>
    <key>ADO_PAT</key><string>YOUR_PAT_HERE</string>
  </dict>
  <key>StartCalendarInterval</key>
  <dict>
    <key>Day</key><integer>1</integer>
    <key>Hour</key><integer>8</integer>
    <key>Minute</key><integer>0</integer>
  </dict>
  <key>StandardOutPath</key><string>/tmp/pr-sla.log</string>
  <key>StandardErrorPath</key><string>/tmp/pr-sla.err</string>
</dict>
</plist>
EOF

launchctl load ~/Library/LaunchAgents/com.amandalapu.pr-sla.plist
```

Only runs while your Mac is awake. GitHub Actions doesn't have that limitation.

---

## How each value is computed

### Dev Complete / PR Submitted Date

```
GET /git/repositories/{repo}/pullRequests/{pr_id}/commits
```
Take the commit with the maximum `committer.date` (fall back to `author.date` if missing). Truncate to a date.

### Approved Date

```
GET /git/repositories/{repo}/pullRequests/{pr_id}/threads
```

ADO records each vote change as a *system thread* with `properties.CodeReviewThreadType == "VoteUpdate"`. The vote value is in `properties.CodeReviewVoteResult.$value`:

| Vote value | Meaning |
|---|---|
| `10` | Approved |
| `5` | Approved with suggestions |
| `0` | No vote / reset |
| `-5` | Waiting for author |
| `-10` | Rejected |

We filter to threads with vote `≥ 5`, take the latest `comments[].publishedDate` across all such threads, and truncate to a date. If no approval vote ever happened (e.g. the PR was force-merged by an admin), we fall back to the PR's `closedDate` and log a warning.

### Days To Process

```python
numpy.busday_count(dev_complete_date, approved_date)
```

- Weekends are excluded.
- Same-day = `0`.
- Approved before Dev Complete (rare; can happen if someone approved, then more commits came in but no re-approval) → clamped to `0`.
- **US public holidays are NOT excluded.** To exclude them, edit `business_days_between` in [src/generate_report.py](src/generate_report.py) to pass a `holidays=[...]` argument to `numpy.busday_count`.

---

## Project layout

```
pr-sla-report/
├── .github/
│   └── workflows/
│       └── monthly-report.yml      # cron + manual dispatch
├── config/
│   └── team-config.xlsx            # repos + team + template (the only file you edit)
├── output/                         # generated .pptx files (gitignored)
├── src/
│   └── generate_report.py          # the whole script
├── .gitignore
├── CONTEXT.md                      # design notes for AI agents working on this repo
├── README.md                       # this file
└── requirements.txt                # Python deps
```

---

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| `ADO_PAT env var is required.` | `export ADO_PAT=<token>` in your shell, or add to `~/.zshrc`. |
| `requests.exceptions.HTTPError: 401 Unauthorized` | PAT expired or wrong scope. Regenerate with Code:Read + Work Items:Read + Identity:Read. |
| `requests.exceptions.HTTPError: 403 Forbidden` | PAT doesn't have access to that specific org/project. Re-issue with "All accessible organizations". |
| `No ADO identity for v-xxx@microsoft.com in <org>` | That UPN can't be resolved in that ADO org. Double-check spelling, or the person doesn't have an account in that org. |
| `No approval vote found, falling back to closedDate` (warning, not an error) | The PR was merged without a recorded reviewer vote (e.g. force-pushed by an admin). The script uses `closedDate` as the best available timestamp. |
| `429 Too Many Requests` in logs but eventually succeeds | Normal — the client retries with backoff. |
| Slide title or table looks off | Open the `.pptx` in PowerPoint (not Quick Look in Finder, which renders fonts differently). |
| GitHub Action fails with secret not found | Add `ADO_PAT` under Settings → Secrets and variables → Actions. |

---

## Updating the script

```bash
cd ~/Personal/pr-sla-report
git pull
.venv/bin/pip install -r requirements.txt   # in case deps changed
```

---

## FAQ

**Q: Can I add more repos?**
Add a row in `Repos` sheet of `team-config.xlsx`. Push the change to GitHub — the workflow picks it up next run.

**Q: Can I add more team members?**
Add a row in `Team` sheet with their full name and UPN.

**Q: Does it count PRs where the person was a reviewer?**
No — only PRs where the person is the **creator/author**. To change this, modify the `searchCriteria.creatorId` filter in `completed_prs_for_user()`.

**Q: What if a PR was approved multiple times (e.g. someone reset their vote and re-approved)?**
The script uses the **latest** approve-vote timestamp it finds.

**Q: Why business days instead of calendar days?**
Matches the existing SLA definition in your screenshot — weekends shouldn't count against the team's turnaround.

**Q: Can I include US holidays?**
Yes, easily. Edit `business_days_between` in [src/generate_report.py](src/generate_report.py):

```python
import numpy as np
US_HOLIDAYS_2026 = ['2026-01-01', '2026-01-19', '2026-02-16', '2026-05-25', ...]
return int(np.busday_count(start, end, holidays=US_HOLIDAYS_2026))
```

**Q: Can I get the deck emailed to me?**
Yes — add a step to the GitHub Actions workflow using e.g. [dawidd6/action-send-mail](https://github.com/dawidd6/action-send-mail) or a Teams webhook. Ask and I'll wire it up.
