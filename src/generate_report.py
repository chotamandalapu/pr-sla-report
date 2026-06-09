#!/usr/bin/env python3
"""
PR SLA Monthly Report Generator
--------------------------------
Reads team + repo config from `Monthly Report Data.xlsx` (at the repo root),
queries Azure DevOps for completed PRs per team member in the configured window
(default: previous calendar month), and emits a PowerPoint deck mirroring the
layout in the Template sheet (one slide per repo).

Usage:
    python src/generate_report.py                # previous calendar month
    python src/generate_report.py --month 2026-05
    python src/generate_report.py --from 2026-04-01 --to 2026-04-30

Auth:
    Set env var ADO_PAT to a Personal Access Token with these scopes:
      - Code (Read)
      - Work Items (Read)
      - Identity (Read)        # needed to resolve user UPN -> descriptor
"""

from __future__ import annotations

import argparse
import base64
import calendar
import datetime as dt
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, urlparse

import numpy as np
import openpyxl
import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Paths + logging
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
CONFIG_XLSX = ROOT / "Monthly Report Data.xlsx"
OUTPUT_DIR = ROOT / "output"

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("pr-sla")


# ---------------------------------------------------------------------------
# Data types
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class RepoRef:
    org: str          # e.g. "MicrosoftIT"
    project: str      # e.g. "OneITVSO"
    repo: str         # e.g. "ConsumerStore-WebComponents"
    url: str          # original URL from config
    short_name: str   # for slide title

    @property
    def api_base(self) -> str:
        return f"https://dev.azure.com/{self.org}/{self.project}/_apis"

    @property
    def web_pr_base(self) -> str:
        return (
            f"https://dev.azure.com/{self.org}/{self.project}/_git/"
            f"{self.repo}/pullrequest"
        )

    @property
    def web_workitem_base(self) -> str:
        return (
            f"https://dev.azure.com/{self.org}/{self.project}/_workitems/edit"
        )


@dataclass
class PrRow:
    repo: RepoRef
    pr_id: int
    pr_title: str
    author_name: str
    author_upn: str
    # "Dev Complete / PR Submitted Date" = last commit pushed to the PR branch.
    dev_complete_date: dt.date
    # "Approved Date" = timestamp of the LAST approval vote on the PR
    # (vote == 10 or vote == 5).  Falls back to PR.closedDate if no vote found.
    approved_date: dt.date
    # Business days between dev_complete_date and approved_date (inclusive of end).
    business_days: int
    work_item_ids: list[int]
    closed_date: dt.date  # kept for filtering / audit


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------
REPO_URL_PATTERNS = [
    # https://dev.azure.com/<org>/<project>/_git/<repo>
    re.compile(r"^https?://dev\.azure\.com/([^/]+)/([^/]+)/_git/([^/?#]+)/?"),
    # https://<org>.visualstudio.com/<project>/_git/<repo>
    re.compile(r"^https?://([^.]+)\.visualstudio\.com/([^/]+)/_git/([^/?#]+)/?"),
]


def parse_repo_url(url: str) -> RepoRef:
    url = url.strip().rstrip("/")
    for pat in REPO_URL_PATTERNS:
        m = pat.match(url)
        if m:
            org, project, repo = m.groups()
            return RepoRef(
                org=org,
                project=project,
                repo=repo,
                url=url,
                short_name=repo,
            )
    raise ValueError(f"Cannot parse repo URL: {url!r}")


def load_config(xlsx: Path) -> tuple[list[RepoRef], list[tuple[str, str]]]:
    """Return (repos, [(displayName, upn)])."""
    wb = openpyxl.load_workbook(xlsx, data_only=True)

    if "Repos" not in wb.sheetnames or "Team" not in wb.sheetnames:
        raise RuntimeError(
            f"Config xlsx must contain sheets 'Repos' and 'Team'; "
            f"found: {wb.sheetnames}"
        )

    repos: list[RepoRef] = []
    ws = wb["Repos"]
    for (val,) in ws.iter_rows(values_only=True, min_row=2):
        if not val:
            continue
        repos.append(parse_repo_url(str(val).strip().replace("\u200b", "")))

    team: list[tuple[str, str]] = []
    ws = wb["Team"]
    for row in ws.iter_rows(values_only=True, min_row=2):
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        upn = str(row[1]).strip() if row[1] else ""
        if not upn:
            log.warning("Skipping team row with no UPN: %s", name)
            continue
        team.append((name, upn))

    log.info("Config: %d repo(s), %d team member(s)", len(repos), len(team))
    return repos, team


# ---------------------------------------------------------------------------
# Time window
# ---------------------------------------------------------------------------
def previous_month_window(today: dt.date | None = None) -> tuple[dt.date, dt.date]:
    today = today or dt.date.today()
    first_of_this_month = today.replace(day=1)
    last_of_prev = first_of_this_month - dt.timedelta(days=1)
    first_of_prev = last_of_prev.replace(day=1)
    return first_of_prev, last_of_prev


def parse_month(spec: str) -> tuple[dt.date, dt.date]:
    yyyy, mm = spec.split("-")
    y, m = int(yyyy), int(mm)
    return dt.date(y, m, 1), dt.date(y, m, calendar.monthrange(y, m)[1])


# ---------------------------------------------------------------------------
# Azure DevOps client
# ---------------------------------------------------------------------------
class AdoClient:
    """Minimal ADO REST client with PAT auth, throttling, and retries."""

    def __init__(self, pat: str):
        if not pat:
            raise RuntimeError("ADO_PAT env var is required.")
        token = base64.b64encode(f":{pat}".encode()).decode()
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Basic {token}",
                "Accept": "application/json",
            }
        )
        # Per-org cache: upn -> ADO identity id (GUID)
        self._user_id_cache: dict[tuple[str, str], str | None] = {}

    def _get(
        self, url: str, *, params: dict | None = None, max_retries: int = 4
    ) -> dict:
        backoff = 1.0
        for attempt in range(max_retries):
            r = self.session.get(url, params=params, timeout=30)
            if r.status_code == 429 or r.status_code >= 500:
                log.warning(
                    "ADO %s on %s (attempt %d/%d); retrying in %.1fs",
                    r.status_code,
                    url,
                    attempt + 1,
                    max_retries,
                    backoff,
                )
                time.sleep(backoff)
                backoff *= 2
                continue
            r.raise_for_status()
            return r.json()
        r.raise_for_status()
        return {}

    # --- identity ---------------------------------------------------------
    def resolve_user_id(self, org: str, upn: str) -> str | None:
        """Return ADO identity id (GUID) for a UPN, or None if not found."""
        key = (org.lower(), upn.lower())
        if key in self._user_id_cache:
            return self._user_id_cache[key]
        url = f"https://vssps.dev.azure.com/{org}/_apis/identities"
        try:
            data = self._get(
                url, params={"searchFilter": "General", "filterValue": upn, "api-version": "7.1"}
            )
            ids = data.get("value", [])
            if not ids:
                log.warning("No ADO identity for %s in %s", upn, org)
                self._user_id_cache[key] = None
                return None
            ident_id = ids[0].get("id")
            self._user_id_cache[key] = ident_id
            return ident_id
        except requests.HTTPError as e:
            log.warning("Identity lookup failed for %s: %s", upn, e)
            self._user_id_cache[key] = None
            return None

    # --- PRs --------------------------------------------------------------
    def completed_prs_for_user(
        self,
        repo: RepoRef,
        upn: str,
        since: dt.date,
        until: dt.date,
    ) -> list[dict]:
        """All completed PRs by `upn` in `repo` whose closedDate is in [since, until]."""
        ident = self.resolve_user_id(repo.org, upn)
        if not ident:
            return []

        url = f"{repo.api_base}/git/repositories/{quote(repo.repo)}/pullrequests"
        # ADO returns max 1000 with $top; page if needed.
        all_prs: list[dict] = []
        skip = 0
        page_size = 200
        # minTime/maxTime filter on closedDate when status=completed.
        # ADO wants ISO-8601 with timezone.
        min_iso = dt.datetime.combine(since, dt.time.min).isoformat() + "Z"
        max_iso = dt.datetime.combine(until, dt.time.max).isoformat() + "Z"

        while True:
            params = {
                "searchCriteria.status": "completed",
                "searchCriteria.creatorId": ident,
                "searchCriteria.minTime": min_iso,
                "searchCriteria.maxTime": max_iso,
                "searchCriteria.queryTimeRangeType": "closed",
                "$top": page_size,
                "$skip": skip,
                "api-version": "7.1",
            }
            data = self._get(url, params=params)
            page = data.get("value", [])
            all_prs.extend(page)
            if len(page) < page_size:
                break
            skip += page_size
        return all_prs

    def pr_last_commit_date(self, repo: RepoRef, pr_id: int) -> dt.date | None:
        url = f"{repo.api_base}/git/repositories/{quote(repo.repo)}/pullRequests/{pr_id}/commits"
        data = self._get(url, params={"api-version": "7.1"})
        commits = data.get("value", [])
        if not commits:
            return None
        # Commits are returned newest-first. committer.date is ISO-8601.
        latest = max(
            commits,
            key=lambda c: c.get("committer", {}).get("date") or c.get("author", {}).get("date") or "",
        )
        ts = (
            latest.get("committer", {}).get("date")
            or latest.get("author", {}).get("date")
        )
        if not ts:
            return None
        return _parse_iso(ts).date()

    def pr_work_item_ids(self, repo: RepoRef, pr_id: int) -> list[int]:
        url = f"{repo.api_base}/git/repositories/{quote(repo.repo)}/pullRequests/{pr_id}/workitems"
        try:
            data = self._get(url, params={"api-version": "7.1"})
        except requests.HTTPError as e:
            log.warning("workitems fetch failed for PR %d: %s", pr_id, e)
            return []
        ids: list[int] = []
        for wi in data.get("value", []):
            try:
                ids.append(int(wi.get("id")))
            except (TypeError, ValueError):
                continue
        return sorted(ids)

    def pr_last_approval_date(
        self, repo: RepoRef, pr_id: int
    ) -> dt.date | None:
        """Latest timestamp at which a reviewer voted Approved (10) or
        Approved-with-suggestions (5) on the PR.

        ADO records each vote change as a *system* thread whose
        `properties.CodeReviewThreadType` == 'VoteUpdate' and whose
        `properties.CodeReviewVoteResult` is the vote value as a string.
        The first comment in that thread is the system message; its
        `publishedDate` is the moment the vote happened.
        """
        url = (
            f"{repo.api_base}/git/repositories/{quote(repo.repo)}"
            f"/pullRequests/{pr_id}/threads"
        )
        try:
            data = self._get(url, params={"api-version": "7.1"})
        except requests.HTTPError as e:
            log.warning("threads fetch failed for PR %d: %s", pr_id, e)
            return None

        latest: dt.datetime | None = None
        for thread in data.get("value", []):
            props = thread.get("properties") or {}
            tt = (props.get("CodeReviewThreadType") or {}).get("$value")
            if tt != "VoteUpdate":
                continue
            vote_raw = (props.get("CodeReviewVoteResult") or {}).get("$value")
            try:
                vote = int(vote_raw) if vote_raw is not None else 0
            except (TypeError, ValueError):
                vote = 0
            # 10 = Approved, 5 = Approved with suggestions.
            # Negative or 0 votes (Reset, Waiting, Rejected) are skipped.
            if vote < 5:
                continue
            comments = thread.get("comments") or []
            ts_str = None
            for c in comments:
                ts_str = c.get("publishedDate") or ts_str
            if not ts_str:
                ts_str = thread.get("publishedDate") or thread.get("lastUpdatedDate")
            if not ts_str:
                continue
            ts = _parse_iso(ts_str)
            if latest is None or ts > latest:
                latest = ts
        return latest.date() if latest else None


def _parse_iso(s: str) -> dt.datetime:
    # ADO returns "2026-04-30T18:23:11.0123456Z" — handle the trailing Z.
    s = s.replace("Z", "+00:00")
    # Trim fractional seconds beyond 6 digits (Python can't parse 7).
    s = re.sub(r"\.(\d{6})\d+", r".\1", s)
    return dt.datetime.fromisoformat(s)


# ---------------------------------------------------------------------------
# Business-day math
# ---------------------------------------------------------------------------
def business_days_between(start: dt.date, end: dt.date) -> int:
    """Exclusive of `start`, inclusive of `end`. Same-day = 0."""
    if end < start:
        return 0
    return int(np.busday_count(start, end))


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------
def extract_rows(
    client: AdoClient,
    repos: list[RepoRef],
    team: list[tuple[str, str]],
    since: dt.date,
    until: dt.date,
) -> dict[RepoRef, list[PrRow]]:
    out: dict[RepoRef, list[PrRow]] = {r: [] for r in repos}
    for repo in repos:
        log.info("== Repo: %s (%s) ==", repo.short_name, repo.org)
        for name, upn in team:
            try:
                prs = client.completed_prs_for_user(repo, upn, since, until)
            except requests.HTTPError as e:
                log.warning("PR query failed for %s @ %s: %s", upn, repo.short_name, e)
                continue
            if prs:
                log.info("  %s (%s): %d PR(s)", name, upn, len(prs))
            for pr in prs:
                pr_id = pr["pullRequestId"]
                closed = _parse_iso(pr["closedDate"]).date()
                last_commit = client.pr_last_commit_date(repo, pr_id) or closed
                last_approval = client.pr_last_approval_date(repo, pr_id)
                if last_approval is None:
                    log.info(
                        "  PR %d: no approval vote found, falling back to closedDate %s",
                        pr_id, closed,
                    )
                    last_approval = closed
                wi_ids = client.pr_work_item_ids(repo, pr_id)
                out[repo].append(
                    PrRow(
                        repo=repo,
                        pr_id=pr_id,
                        pr_title=pr.get("title", ""),
                        author_name=name,
                        author_upn=upn,
                        dev_complete_date=last_commit,
                        approved_date=last_approval,
                        business_days=business_days_between(
                            last_commit, last_approval
                        ),
                        work_item_ids=wi_ids,
                        closed_date=closed,
                    )
                )
    return out


# ---------------------------------------------------------------------------
# PowerPoint generation
# ---------------------------------------------------------------------------
HEADER_LEFT = RGBColor(0x18, 0xB2, 0xA8)   # teal
HEADER_RIGHT = RGBColor(0x1B, 0x76, 0xD2)  # blue
HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
TITLE_RED = RGBColor(0xE6, 0x1C, 0x1C)

COLUMNS = [
    ("ADO Link", 1.6),
    ("PR", 3.4),
    ("Dev Complete/PR Submitted Date", 1.6),
    ("Approved Date", 1.3),
    ("Days To Process", 1.1),
    ("Comments", 1.8),
]


def _set_cell_fill(cell, rgb: RGBColor) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb


def _set_cell_text(
    cell,
    text: str,
    *,
    font_size: int = 9,
    bold: bool = False,
    color: RGBColor | None = None,
    hyperlink: str | None = None,
    align=PP_ALIGN.LEFT,
):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if hyperlink:
        run.hyperlink.address = hyperlink


def _build_ado_links_cell(cell, ids: list[int], repo: RepoRef):
    """Render space-separated hyperlinked ADO IDs in a single cell."""
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if not ids:
        run = p.add_run()
        run.text = "-"
        run.font.size = Pt(9)
        return
    for i, wi_id in enumerate(ids):
        if i > 0:
            sep = p.add_run()
            sep.text = "  "
            sep.font.size = Pt(9)
        run = p.add_run()
        run.text = str(wi_id)
        run.font.size = Pt(9)
        run.font.underline = True
        run.font.color.rgb = RGBColor(0x0B, 0x5C, 0xAB)
        run.hyperlink.address = f"{repo.web_workitem_base}/{wi_id}"


def render_slide(prs: Presentation, repo: RepoRef, rows: list[PrRow], month_label: str):
    """One slide per repo, table layout matching the screenshot."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(12.5), Inches(0.6))
    p = tb.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = f"{repo.short_name} | PR SLA ({month_label})"
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = TITLE_RED

    # Table dimensions
    n_cols = len(COLUMNS)
    n_rows = len(rows) + 1  # +1 header
    col_widths_in = [w for _, w in COLUMNS]

    left = Inches(0.4)
    top = Inches(1.05)
    width = Inches(sum(col_widths_in))
    row_height = Inches(0.55)
    height = Inches(0.55 + 0.55 * len(rows))

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    tbl = tbl_shape.table

    # Set column widths
    for i, w in enumerate(col_widths_in):
        tbl.columns[i].width = Inches(w)

    # Header row — gradient via linear interpolation across cells
    for i, (label, _) in enumerate(COLUMNS):
        cell = tbl.cell(0, i)
        # blend teal -> blue across the header
        t = i / max(1, n_cols - 1)
        r = int(HEADER_LEFT[0] + (HEADER_RIGHT[0] - HEADER_LEFT[0]) * t)
        g = int(HEADER_LEFT[1] + (HEADER_RIGHT[1] - HEADER_LEFT[1]) * t)
        b = int(HEADER_LEFT[2] + (HEADER_RIGHT[2] - HEADER_LEFT[2]) * t)
        _set_cell_fill(cell, RGBColor(r, g, b))
        _set_cell_text(
            cell, label,
            font_size=10, bold=True, color=HEADER_TEXT,
        )

    # Body rows
    for ri, row in enumerate(rows, start=1):
        # ADO Link (work item IDs, hyperlinked)
        _build_ado_links_cell(tbl.cell(ri, 0), row.work_item_ids, row.repo)

        # PR URL (display the canonical URL form like the screenshot)
        pr_url = f"{row.repo.web_pr_base}/{row.pr_id}"
        _set_cell_text(tbl.cell(ri, 1), pr_url, font_size=8, hyperlink=pr_url)

        # Dev Complete (= last commit date) and Approved (= last approval vote)
        _set_cell_text(
            tbl.cell(ri, 2), row.dev_complete_date.strftime("%m/%d/%Y"),
            font_size=9,
        )
        _set_cell_text(
            tbl.cell(ri, 3), row.approved_date.strftime("%m/%d/%Y"),
            font_size=9,
        )
        _set_cell_text(
            tbl.cell(ri, 4), str(row.business_days),
            font_size=9, align=PP_ALIGN.CENTER,
        )
        _set_cell_text(tbl.cell(ri, 5), "", font_size=9)

        # Force row height
        tbl.rows[ri].height = row_height


SORT_CHOICES = ("dev-complete", "approved", "days", "author")


def _sort_key(row: PrRow, by: str):
    if by == "dev-complete":
        return (row.dev_complete_date, row.approved_date, row.pr_id)
    if by == "approved":
        return (row.approved_date, row.dev_complete_date, row.pr_id)
    if by == "days":
        return (row.business_days, row.approved_date, row.pr_id)
    # author
    return (row.author_name.lower(), row.dev_complete_date, row.pr_id)


def build_deck(
    data: dict[RepoRef, list[PrRow]],
    month_label: str,
    out_path: Path,
    sort_by: str = "approved",
    sort_desc: bool = False,
):
    prs = Presentation()
    # 16:9 widescreen, 13.33 x 7.5 in
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for repo, rows in data.items():
        rows.sort(key=lambda r: _sort_key(r, sort_by), reverse=sort_desc)
        render_slide(prs, repo, rows, month_label)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    log.info("Wrote %s", out_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Generate monthly PR SLA PowerPoint.")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--month", help="YYYY-MM (defaults to previous calendar month).")
    g.add_argument("--from", dest="from_date", help="Start date YYYY-MM-DD.")
    ap.add_argument("--to", dest="to_date", help="End date YYYY-MM-DD (with --from).")
    ap.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output .pptx path (default: output/PR_SLA_<MONTH>.pptx).",
    )
    ap.add_argument(
        "--sort-by",
        choices=SORT_CHOICES,
        default="approved",
        help=(
            "How to sort rows within each repo slide. "
            "'dev-complete' = by last commit date; "
            "'approved' = by last approval date (default); "
            "'days' = by business days to process; "
            "'author' = alphabetical by name."
        ),
    )
    ap.add_argument(
        "--desc",
        action="store_true",
        help="Sort descending (default: ascending).",
    )
    args = ap.parse_args(argv)

    if args.from_date:
        if not args.to_date:
            ap.error("--from requires --to")
        since = dt.date.fromisoformat(args.from_date)
        until = dt.date.fromisoformat(args.to_date)
        month_label = f"{since.isoformat()} → {until.isoformat()}"
        slug = f"{since.isoformat()}_{until.isoformat()}"
    elif args.month:
        since, until = parse_month(args.month)
        month_label = since.strftime("%B %Y")
        slug = args.month
    else:
        since, until = previous_month_window()
        month_label = since.strftime("%B %Y")
        slug = since.strftime("%Y-%m")

    log.info("Window: %s → %s (%s)", since, until, month_label)

    pat = os.environ.get("ADO_PAT", "")
    repos, team = load_config(CONFIG_XLSX)
    client = AdoClient(pat)

    data = extract_rows(client, repos, team, since, until)

    total = sum(len(v) for v in data.values())
    log.info("Total PRs across all repos: %d", total)

    out_path = args.output or (OUTPUT_DIR / f"PR_SLA_{slug}.pptx")
    build_deck(
        data,
        month_label,
        out_path,
        sort_by=args.sort_by,
        sort_desc=args.desc,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
