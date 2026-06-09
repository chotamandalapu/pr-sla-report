#!/usr/bin/env python3
"""
Diagnose why specific PRs were missed or had wrong dates in the SLA report.

Usage:
    python diag_pr.py 1872660 1934106
"""
from __future__ import annotations
import datetime as dt
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from generate_report import (
    AdoClient, RepoRef, parse_repo_url, load_config, CONFIG_XLSX,
    _parse_iso,
)

REPO_URL = "https://dev.azure.com/MicrosoftIT/OneITVSO/_git/ConsumerStore-WebComponents"


def fmt(dt_: dt.datetime) -> str:
    return dt_.strftime("%Y-%m-%d %H:%M:%S UTC")


def main(pr_ids: list[int]) -> int:
    pat = os.environ.get("ADO_PAT")
    if not pat:
        print("Set ADO_PAT first.")
        return 1
    client = AdoClient(pat)
    repo = parse_repo_url(REPO_URL)
    repos, team = load_config(CONFIG_XLSX)

    team_upns = {u.lower() for _, u in team}
    print(f"Team UPNs ({len(team_upns)}):")
    for n, u in team:
        print(f"   {u:40s}  {n}")
    print()

    for pr_id in pr_ids:
        print(f"================================================================")
        print(f" PR {pr_id}")
        print(f"================================================================")

        # Raw PR details
        url = f"{repo.api_base}/git/repositories/{repo.repo}/pullRequests/{pr_id}"
        try:
            pr = client._get(url, params={"api-version": "7.1"})
        except Exception as e:
            print(f"  ERROR fetching PR: {e}")
            continue

        creator_display = pr.get("createdBy", {}).get("displayName", "?")
        creator_upn     = pr.get("createdBy", {}).get("uniqueName", "?")
        status          = pr.get("status", "?")
        created_at      = pr.get("creationDate")
        closed_at       = pr.get("closedDate")
        title           = pr.get("title", "?")

        in_team = creator_upn.lower() in team_upns

        print(f"  Title:         {title}")
        print(f"  Creator name:  {creator_display}")
        print(f"  Creator UPN:   {creator_upn}    {'✓ IN team' if in_team else '✗ NOT IN team list'}")
        print(f"  Status:        {status}")
        print(f"  creationDate:  {created_at}")
        print(f"  closedDate:    {closed_at}")
        if closed_at:
            closed_dt = _parse_iso(closed_at)
            print(f"     parsed UTC:  {fmt(closed_dt)}  → .date() = {closed_dt.date()}")
            # Compare to a local timezone interpretation
            local = closed_dt.astimezone()
            print(f"     local time:  {local.strftime('%Y-%m-%d %H:%M:%S %Z')}  → date {local.date()}")

        # Commits
        commits_url = (f"{repo.api_base}/git/repositories/{repo.repo}"
                       f"/pullRequests/{pr_id}/commits")
        try:
            data = client._get(commits_url, params={"api-version": "7.1"})
        except Exception as e:
            print(f"  ERROR fetching commits: {e}")
            data = {"value": []}
        commits = data.get("value", [])
        print(f"  Commits:       {len(commits)}")
        if commits:
            latest = max(
                commits,
                key=lambda c: c.get("committer", {}).get("date") or c.get("author", {}).get("date") or ""
            )
            ts = latest.get("committer", {}).get("date") or latest.get("author", {}).get("date")
            if ts:
                cdt = _parse_iso(ts)
                local = cdt.astimezone()
                print(f"  last commit:   {ts}")
                print(f"     parsed UTC:  {fmt(cdt)}  → .date() = {cdt.date()}")
                print(f"     local time:  {local.strftime('%Y-%m-%d %H:%M:%S %Z')}  → date {local.date()}")

        # Approval threads
        threads_url = (f"{repo.api_base}/git/repositories/{repo.repo}"
                       f"/pullRequests/{pr_id}/threads")
        try:
            tdata = client._get(threads_url, params={"api-version": "7.1"})
        except Exception as e:
            print(f"  ERROR fetching threads: {e}")
            tdata = {"value": []}

        approvals = []
        for th in tdata.get("value", []):
            props = th.get("properties") or {}
            tt = (props.get("CodeReviewThreadType") or {}).get("$value")
            if tt != "VoteUpdate":
                continue
            vote_raw = (props.get("CodeReviewVoteResult") or {}).get("$value")
            try:
                vote = int(vote_raw) if vote_raw is not None else 0
            except (TypeError, ValueError):
                vote = 0
            comments = th.get("comments") or []
            ts = comments[-1].get("publishedDate") if comments else None
            approvals.append((vote, ts, th.get("identities")))
        print(f"  Vote threads:  {len(approvals)}")
        for v, ts, _ in approvals:
            cdt = _parse_iso(ts) if ts else None
            local = cdt.astimezone() if cdt else None
            print(f"     vote={v:+d}  at {ts}   (local {local.date() if local else '?'})")
        approve_ts = [ts for v, ts, _ in approvals if v >= 5]
        if approve_ts:
            latest_apr = max(approve_ts)
            apr_dt = _parse_iso(latest_apr)
            print(f"  Latest APPROVE (vote>=5): {latest_apr}")
            print(f"     parsed UTC:  {fmt(apr_dt)}  → .date() = {apr_dt.date()}")
            print(f"     local:       {apr_dt.astimezone().strftime('%Y-%m-%d %H:%M %Z')}  → {apr_dt.astimezone().date()}")
        else:
            print(f"  No approval vote (vote>=5) found — script would fall back to closedDate.")
        print()

    return 0


if __name__ == "__main__":
    sys.exit(main([int(a) for a in sys.argv[1:]] or [1872660, 1934106]))
