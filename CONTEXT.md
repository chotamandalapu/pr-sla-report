# CONTEXT.md — for AI agents working on `pr-sla-report`

> This file is for any AI assistant (Copilot, Claude, etc.) modifying this repo in future sessions.
> User-facing documentation lives in [README.md](README.md). This file documents **why** the code is shaped the way it is.

---

## 1. What this project is

A small Python CLI that pulls completed pull-request data from Azure DevOps and emits a PowerPoint deck that matches a pre-existing manual monthly report format. It is run **once a month** (by GitHub Actions cron or by the user manually). It has **no UI**, **no web server**, **no database** — just `Excel config` → `ADO REST API` → `.pptx`.

The single source of truth for repos + team membership is [`Monthly Report Data.xlsx`](Monthly%20Report%20Data.xlsx) at the repo root. The script is intentionally driven by that file, not by env vars or CLI flags, because the end user maintains it in Excel as they always have.

---

## 2. Architecture (1-page)

```
Monthly Report Data.xlsx           (user-editable; 3 sheets: Repos, Team, Template)
       │
       ▼
 run_interactive.py  ←── the front door for humans
   ├── 5 prompts (repos confirm / team confirm / PAT / month / sort)
   ├── PAT collected via getpass (hidden input, never written to disk)
   └── calls into
         generate_report.py        ←── also a standalone CLI for cron/CI
           ├── load_config()                 → [RepoRef], [(name, upn)]
           ├── time-window resolution        → previous month / --month / --from --to
           ├── AdoClient
           │     ├── resolve_user_id(upn)    → ADO identity GUID (cached per org)
           │     ├── completed_prs_for_user  → paginated GET pullrequests?status=completed&creatorId=
           │     ├── pr_last_commit_date     → GET pullrequests/{id}/commits  → max(committer.date)
           │     ├── pr_last_approval_date   → GET pullrequests/{id}/threads  → latest VoteUpdate with vote≥5
           │     └── pr_work_item_ids        → GET pullrequests/{id}/workitems
           ├── extract_rows()                → dict[RepoRef, list[PrRow]]
           ├── _sort_key()                   → respects --sort-by + --desc
           └── build_deck()                  → python-pptx, one slide per repo
                                                  │
                                                  ▼
                                     output/PR_SLA_<YYYY-MM>.pptx
```

There are two entry points:

* **`run_interactive.py`** (front door for humans, called by `./run`) — 5 prompts,
  ANSI-styled output, hidden PAT input. Imports from `generate_report` and never
  duplicates business logic.
* **`generate_report.py`** (also a CLI: `python src/generate_report.py --month
  2026-05 --sort-by approved`) — used by GitHub Actions cron because non-
  interactive runs can't stop for prompts.

---

## 3. Key data types

### `RepoRef`
Frozen dataclass. Holds `org / project / repo` plus the original URL. The split into `dev.azure.com` vs `<org>.visualstudio.com` is preserved because work-item links use the modern host but historic data may carry legacy URLs.

### `PrRow`
One row of the output table. Fields:

- `dev_complete_date`  — date of last commit on the PR branch.
- `approved_date`      — date of last reviewer **approval** vote (vote ≥ 5), NOT the PR `closedDate`.
- `business_days`      — `numpy.busday_count(dev_complete_date, approved_date)`.
- `closed_date`        — kept for fallback + audit only; never displayed if `approved_date` was resolvable.

**Why three dates instead of two?** The user originally conflated "completed" with "approved". A PR can be approved on day X but only auto-completed days later (e.g. policy waiting for build). The SLA the team is measured against is **approval turnaround**, so `approved_date` is the metric.

---

## 4. ADO REST surface used

| Endpoint | Used for | Notes |
|---|---|---|
| `GET https://vssps.dev.azure.com/{org}/_apis/identities?searchFilter=General&filterValue={upn}` | UPN → identity GUID | Org-scoped, cached. Some accounts span multiple orgs (MicrosoftIT vs microsoftit lowercased) — resolved separately per org. |
| `GET /git/repositories/{repo}/pullrequests?status=completed&creatorId=&minTime=&maxTime=&queryTimeRangeType=closed` | Find a person's completed PRs in window | `queryTimeRangeType=closed` makes `min/maxTime` apply to `closedDate`. Paginated via `$skip` + `$top=200`. |
| `GET /git/repositories/{repo}/pullRequests/{id}/commits` | Last commit on PR branch | Take `committer.date` preferring max; fall back to `author.date`. |
| `GET /git/repositories/{repo}/pullRequests/{id}/threads` | Approval timestamps | See section 5. |
| `GET /git/repositories/{repo}/pullRequests/{id}/workitems` | Linked work item IDs | Returns just IDs; we don't fetch titles to keep request volume low. |

All endpoints use `api-version=7.1`. Auth is Basic with PAT: `Authorization: Basic base64(":" + PAT)`.

### Throttling

`AdoClient._get()` retries on `429` and `5xx` with `1s → 2s → 4s → 8s` backoff (max 4 attempts). ADO PR endpoints are rate-limited per org per user. Do NOT remove the retry — large team × large repo × deep history can hit `429` cleanly.

---

## 5. Approval timestamps — the non-obvious part

ADO's `pullRequest.reviewers[]` payload includes each reviewer's **current** vote but **not** when they cast it. To get the timestamp, you have to read the PR threads and find the system-generated **vote-update** threads.

### Vote update thread shape

```json
{
  "id": 123,
  "comments": [
    {
      "commentType": "system",
      "content": "voted 10",
      "publishedDate": "2026-05-01T15:42:11Z",
      "author": {...}
    }
  ],
  "properties": {
    "CodeReviewThreadType": { "$type": "System.String", "$value": "VoteUpdate" },
    "CodeReviewVoteResult": { "$type": "System.Int32", "$value": "10" }
  }
}
```

`pr_last_approval_date()` filters threads to those where:
1. `properties.CodeReviewThreadType.$value == "VoteUpdate"`, AND
2. `properties.CodeReviewVoteResult.$value` parses as integer `>= 5`.

Then it takes the maximum `comments[].publishedDate` across all matching threads and truncates to a date.

### Vote scale (ADO)

| Value | Meaning | Counts as approval? |
|---|---|---|
| 10 | Approved | ✅ |
| 5 | Approved with suggestions | ✅ |
| 0 | No vote (reset) | ❌ |
| -5 | Waiting for author | ❌ |
| -10 | Rejected | ❌ |

### Fallback

If no qualifying thread is found, fall back to `pr.closedDate` and emit an INFO log (not a warning — this is normal for force-merged PRs). Never silently drop the row.

---

## 6. Business-day math

`business_days_between(start, end)` returns `int(numpy.busday_count(start, end))`. By NumPy contract this is **exclusive of start, inclusive of end**, so:

- Same calendar day → `0`
- Friday → Monday → `1`
- Negative `end < start` → clamped to `0` (returned before calling NumPy)

**Holidays are intentionally NOT included.** The user opted not to track them in v1 to avoid the bikeshed of "whose holiday calendar". Adding them is a one-line change documented in the README FAQ.

### 6.1 Timezone conversion (v1.1.0+)

ADO returns all timestamps in UTC. Before any timestamp is truncated to a date,
it's converted to the **reporting timezone** stored in the module-level
`_report_tz` (set via `set_report_tz(name)` or `--tz`). Default:
`America/Los_Angeles` (PT — chosen because Microsoft-internal reports usually
frame events in PT).

**Rule of thumb:** any new code that calls `dt.date()` or `.date()` on a parsed
ADO timestamp is a bug. Use `to_report_date(parsed_dt)` instead. The three
existing call sites are:
- `AdoClient.pr_last_commit_date()` → last commit timestamp
- `AdoClient.pr_last_approval_date()` → last approval-vote timestamp
- `extract_rows()` → PR `closedDate`

A history-debugging utility lives at `src/diag_pr.py` — it dumps creator UPN,
commits, threads, and votes for any given PR ID. Use it when a report row
looks wrong:
```
.venv/bin/python src/diag_pr.py 1934106 1872660
```

---

## 7. PowerPoint generation (`python-pptx`)

### Slide dimensions
16:9 widescreen, hardcoded to `13.333 × 7.5 inches`. Matches the user's existing template export.

### Column widths
Defined in `COLUMNS` constant near the top of `render_slide`. Total = 10.8 inches; leaves ~2.5 inches of horizontal padding on the slide.

### Header gradient
Linear interpolation between `HEADER_LEFT` (teal `#18B2A8`) and `HEADER_RIGHT` (blue `#1B76D2`) per cell. `python-pptx` doesn't expose true gradient fills on table cells, so we cheat by assigning each header cell a solid color along the gradient. Visually indistinguishable at slide-display resolution.

### Hyperlinks
- Work-item IDs: each ID is its own `Run` inside a single cell, with `run.hyperlink.address` set. Multiple IDs are joined with a `"  "` separator run (no underline on the separator).
- PR URL: the cell contains a single Run with hyperlink to the PR.

### Why not python-pptx's table style?
We need per-cell fills and per-run hyperlinks; setting `table.first_row` styles wipes our custom fills. So we manually fill every cell.

---

## 8. Sorting

`_sort_key(row, by)` returns a tuple chosen so that the **secondary** sort key gives stable, sensible ordering when the primary ties. e.g. `--sort-by approved` ties broken by `dev_complete_date` then `pr_id`. This stability matters because tests compare row order across runs.

`--desc` is just `reverse=True` passed to `list.sort()`. Don't try to invert specific dimensions selectively — adds bugs without clear benefit.

---

## 9. Things deliberately NOT done

| Not done | Why |
|---|---|
| Caching of ADO responses | Run is monthly + ad-hoc; not worth the cache-invalidation complexity. |
| Multi-month "trend" output | User explicitly wants one deck per month matching their existing format. |
| Reviewer-side filtering ("PRs reviewed by X") | Out of scope; current SLA is creator-centric. |
| Pulling work-item titles | Doubles request count; user just needs ID + hyperlink. |
| HTML / Markdown output | PowerPoint is the required deliverable. |
| Test suite | Project is ~500 LOC with smoke tests in the README. Unit tests would mostly be mocking the ADO client. Add them only if the project grows. |
| Python packaging (pyproject.toml + entrypoint) | One file, one user, one schedule. Overhead doesn't pay off. |

---

## 10. Common modification recipes

### Add a new column to the output

1. Append `("Header label", width_in_inches)` to the `COLUMNS` list in `render_slide`.
2. Add a field to `PrRow`.
3. Populate it in `extract_rows()`.
4. Render it in `render_slide()` after the existing 6 columns.

### Filter by something other than "PR creator"

`completed_prs_for_user()` uses `searchCriteria.creatorId`. Other options ADO accepts:
- `searchCriteria.reviewerId` — PRs reviewed by that identity.
- Drop both → all PRs in the repo for the window (then filter client-side however you want).

### Add US holidays to business-day math

```python
US_HOLIDAYS = np.array(['2026-01-01', '2026-01-19', ...], dtype='datetime64[D]')
def business_days_between(start, end):
    if end < start: return 0
    return int(np.busday_count(start, end, holidays=US_HOLIDAYS))
```

### Send the deck somewhere after generation

Add a step to [.github/workflows/monthly-report.yml](.github/workflows/monthly-report.yml) AFTER the "Generate report" step. The generated `.pptx` is at `output/PR_SLA_*.pptx`. Options:

- Email: [`dawidd6/action-send-mail`](https://github.com/dawidd6/action-send-mail).
- Teams: incoming webhook POST with `file_path` upload.
- Slack: [`MeilCli/slack-upload-file`](https://github.com/MeilCli/slack-upload-file).
- SharePoint: Microsoft Graph API upload.

### Add a new repo

Just edit `Monthly Report Data.xlsx` → `Repos` sheet. No code change.

### Add a new team member

Edit `Monthly Report Data.xlsx` → `Team` sheet. The script resolves their UPN to an identity GUID on next run.

---

## 11. Files and what they own

| File | Owner | Purpose |
|---|---|---|
| `src/generate_report.py` | Everything: config load, ADO client, extraction, PPT build, CLI | Core engine + non-interactive CLI |
| `src/run_interactive.py` | UX | 5-prompt interactive wrapper. Imports from `generate_report`. PAT via getpass. Never writes the PAT to disk. |
| `run` | UX | Bash launcher: auto-creates `.venv`, installs `requirements.txt`, execs the interactive runner. |
| `Monthly Report Data.xlsx` | User | Repos + Team + Template reference |
| `requirements.txt` | Dev deps | `openpyxl`, `python-pptx`, `requests`, `numpy` — pinned to major versions, not exact |
| `.github/workflows/monthly-report.yml` | CI | Monthly cron + manual dispatch + artifact upload (uses generate_report.py CLI directly, NOT the interactive runner) |
| `output/` | Generated | All `.pptx` files; gitignored |
| `.venv/` | Local | Python venv; gitignored |
| `README.md` | User-facing | Setup, usage, features, troubleshooting |
| `CONTEXT.md` | AI/dev-facing | This file |

---

## 12. Tests / smoke checks

There is no `tests/` directory. The contract surface is small and the dominant risk is ADO API drift. Recommended smoke procedure when changing extraction logic:

```bash
# 1. Compile check
python -m py_compile src/generate_report.py

# 2. Config load + math sanity
python -c "
import sys; sys.path.insert(0, 'src')
from generate_report import load_config, CONFIG_XLSX, business_days_between
import datetime as dt
repos, team = load_config(CONFIG_XLSX)
assert len(repos) >= 1
assert len(team) >= 1
assert business_days_between(dt.date(2026,4,30), dt.date(2026,5,1)) == 1
assert business_days_between(dt.date(2026,4,3),  dt.date(2026,4,3))  == 0
print('OK')
"

# 3. Full live run (requires ADO_PAT)
ADO_PAT=<token> python src/generate_report.py --month 2026-05
```

If you change the PPT building code (`build_deck` / `render_slide`), you should diff the resulting `.pptx` visually against a known-good output. There's no automated check.

---

## 13. Security

- The PAT is the only secret. It is read from `ADO_PAT` env var. **Never log it, never echo it, never put it in CLI args**.
- In GitHub Actions, `ADO_PAT` is set from a repo secret. The workflow file references `${{ secrets.ADO_PAT }}` only — GitHub masks the value in logs.
- The `.gitignore` excludes `.env`. If you ever add a `.env` for local dev, do **not** commit it.
- All ADO calls are read-only. The PAT scopes (`Code:Read`, `Work Items:Read`, `Identity:Read`) cannot mutate ADO state even if compromised.
- `Monthly Report Data.xlsx` is committed because UPNs are not sensitive. If your team list ever becomes sensitive, move it to a private location or to GitHub Actions secrets.

---

## 14. Known limitations

| Limitation | Mitigation |
|---|---|
| If a person's UPN doesn't exist in one of the configured orgs, they're skipped for that repo with a warning. | Acceptable — usually means they don't work in that org. |
| `--month` of the *current* month will return only PRs closed up to `today` because of `maxTime`. | Default is previous month; current month is opt-in. |
| Approval timestamps require the PR to have a real reviewer-vote thread. Some PRs (force-merged by admins) have none → falls back to `closedDate`. | Falls back gracefully with INFO log. |
| Holiday-aware business days require code change (not config). | Documented in the FAQ. |
| Single Excel file is the config — no CLI override for repos/team. | Intentional: one source of truth. |
