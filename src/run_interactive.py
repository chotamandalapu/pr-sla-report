#!/usr/bin/env python3
"""
PR SLA — interactive runner.

Walks the user through the five prompts they asked for, then runs the existing
extraction + PPT build pipeline.  PAT is collected via getpass so it never
appears in terminal history, process listings, or scrollback.

Run with:
    .venv/bin/python src/run_interactive.py
or (after `chmod +x ./run`):
    ./run
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import getpass
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

# Allow `python src/run_interactive.py` from the repo root.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_report import (
    CONFIG_XLSX,
    OUTPUT_DIR,
    SORT_CHOICES,
    AdoClient,
    build_deck,
    configure_holidays,
    extract_rows,
    load_config,
    parse_month,
)


# ---------------------------------------------------------------------------
# Tiny ANSI helpers (no external deps)
# ---------------------------------------------------------------------------
def _supports_color() -> bool:
    return sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


_C = _supports_color()
BOLD   = "\033[1m"  if _C else ""
DIM    = "\033[2m"  if _C else ""
RED    = "\033[31m" if _C else ""
GREEN  = "\033[32m" if _C else ""
YELLOW = "\033[33m" if _C else ""
BLUE   = "\033[34m" if _C else ""
CYAN   = "\033[36m" if _C else ""
RESET  = "\033[0m"  if _C else ""


def hr(char: str = "─", width: int = 72) -> str:
    return DIM + (char * width) + RESET


def banner(title: str) -> None:
    print()
    print(hr("═"))
    print(f"{BOLD}{CYAN}{title}{RESET}")
    print(hr("═"))


def step(n: int, total: int, title: str) -> None:
    print()
    print(f"{BOLD}[{n}/{total}]{RESET} {BOLD}{title}{RESET}")
    print(hr())


def ok(msg: str) -> None:
    print(f"  {GREEN}✓{RESET} {msg}")


def warn(msg: str) -> None:
    print(f"  {YELLOW}!{RESET} {msg}")


def fail(msg: str) -> None:
    print(f"  {RED}✗{RESET} {msg}")


# ---------------------------------------------------------------------------
# Prompt primitives
# ---------------------------------------------------------------------------
def ask_yes_no(prompt: str, default: bool = True) -> bool:
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        try:
            raw = input(f"  {prompt}{suffix}").strip().lower()
        except EOFError:
            print()
            sys.exit(1)
        if not raw:
            return default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        warn("Please answer y or n.")


def ask_text(prompt: str, *, default: str | None = None,
             validator=None, error_msg: str = "Invalid input.") -> str:
    suffix = f" [{default}]: " if default else ": "
    while True:
        try:
            raw = input(f"  {prompt}{suffix}").strip()
        except EOFError:
            print()
            sys.exit(1)
        if not raw and default is not None:
            raw = default
        if not raw:
            warn("Value required.")
            continue
        if validator and not validator(raw):
            warn(error_msg)
            continue
        return raw


def ask_secret(prompt: str) -> str:
    while True:
        try:
            raw = getpass.getpass(f"  {prompt} (input hidden): ").strip()
        except EOFError:
            print()
            sys.exit(1)
        if not raw:
            warn("PAT cannot be empty.")
            continue
        return raw


# ---------------------------------------------------------------------------
# Prompt 0 — verify mandatory prerequisites (Python + Git)
# ---------------------------------------------------------------------------
MIN_PYTHON = (3, 10)


def _detect_python_version() -> str:
    v = sys.version_info
    return f"{v.major}.{v.minor}.{v.micro}"


def _detect_git_version() -> str | None:
    git_path = shutil.which("git")
    if not git_path:
        return None
    try:
        out = subprocess.run(
            [git_path, "--version"],
            capture_output=True, text=True, timeout=5, check=True,
        )
        return out.stdout.strip()
    except (subprocess.SubprocessError, OSError):
        return None


def prompt_prerequisites() -> None:
    """Step 0: confirm Python and Git are installed before doing anything else."""
    step(0, 6, "Mandatory prerequisites")

    py_ver = _detect_python_version()
    py_ok = sys.version_info >= MIN_PYTHON
    git_ver = _detect_git_version()
    git_ok = git_ver is not None

    print(f"  Required on this machine:")
    print(
        f"    • {BOLD}Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+{RESET}  "
        f"{'(detected: ' + py_ver + ')' if py_ok else '(detected: ' + py_ver + ' — TOO OLD)'}"
    )
    print(
        f"    • {BOLD}Git{RESET}             "
        f"{'(detected: ' + git_ver + ')' if git_ok else '(NOT FOUND)'}"
    )
    print()
    print(f"  {DIM}If either is missing, install with:{RESET}")
    print(f"    {DIM}macOS:   brew install python@3.12 git{RESET}")
    print(f"    {DIM}Linux:   sudo apt install python3.12 python3.12-venv git{RESET}")
    print(f"    {DIM}Windows: https://www.python.org/downloads/  +  https://git-scm.com/{RESET}")
    print()

    if py_ok:
        ok(f"Python {py_ver} detected.")
    else:
        fail(
            f"Python {py_ver} is too old (need {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+)."
        )
    if git_ok:
        ok(f"Git detected: {git_ver}")
    else:
        fail("Git not found in PATH.")

    print()
    answer = ask_yes_no(
        "Confirm both Python and Git are installed and you want to continue?",
        default=py_ok and git_ok,
    )
    if not answer:
        fail("Install the missing prerequisites and re-run ./run.")
        sys.exit(1)
    if not (py_ok and git_ok):
        warn("Continuing despite missing prerequisites — expect failures.")


# ---------------------------------------------------------------------------
# Prompt 1 — confirm Repos sheet
# ---------------------------------------------------------------------------
def prompt_confirm_repos(repos) -> bool:
    step(1, 6, "Confirm the Repos sheet")
    print(f"  Configured in {DIM}{CONFIG_XLSX.name}{RESET} → sheet {BOLD}Repos{RESET}:")
    print()
    for i, r in enumerate(repos, 1):
        print(f"    {i}. {BLUE}{r.url}{RESET}")
    print()
    answer = ask_yes_no("Are these the correct repos?", default=True)
    if not answer:
        fail("Edit the Repos sheet in Monthly Report Data.xlsx, save, and re-run.")
        sys.exit(1)
    ok("Repos confirmed.")
    return True


# ---------------------------------------------------------------------------
# Prompt 2 — confirm Team sheet
# ---------------------------------------------------------------------------
def prompt_confirm_team(team) -> bool:
    step(2, 6, "Confirm the Team sheet")
    print(f"  Configured in {DIM}{CONFIG_XLSX.name}{RESET} → sheet {BOLD}Team{RESET}:")
    print()
    print(f"    {DIM}{'Name':<28}{RESET} {DIM}{'V-id (UPN)':<40}{RESET}")
    print(f"    {DIM}{'-'*28}{RESET} {DIM}{'-'*40}{RESET}")
    for name, upn in team:
        print(f"    {name:<28} {BLUE}{upn}{RESET}")
    print()
    answer = ask_yes_no(
        f"Are all {len(team)} team members correct and up to date?",
        default=True,
    )
    if not answer:
        fail("Edit the Team sheet in Monthly Report Data.xlsx, save, and re-run.")
        sys.exit(1)
    ok("Team confirmed.")
    return True


# ---------------------------------------------------------------------------
# Prompt 3 — collect credentials (PAT or Azure CLI bearer token)
# ---------------------------------------------------------------------------
def prompt_auth() -> tuple[str, str]:
    """Returns (kind, value) where kind is 'pat' or 'bearer'."""
    step(3, 6, "Azure DevOps credentials")

    # 1. Env var ADO_PAT wins if present.
    existing = os.environ.get("ADO_PAT", "").strip()
    if existing:
        ok(f"Found ADO_PAT in environment ({len(existing)} chars).")
        if ask_yes_no("Use the PAT from your environment?", default=True):
            return ("pat", existing)
        warn("Ignoring environment value.")

    # 2. Try `az` CLI for a bearer token (no secret in dotfiles).
    from generate_report import _try_az_token  # local import to avoid cycle
    import shutil
    az_present = shutil.which("az") is not None
    if az_present:
        print(f"  {DIM}Trying Azure CLI for a bearer token...{RESET}")
        token = _try_az_token()
        if token:
            ok(f"Got a bearer token from `az` ({len(token)} chars).")
            print(f"  {DIM}No PAT needed. This token is valid for ~1\u201312 hours.{RESET}")
            if ask_yes_no("Use this token?", default=True):
                return ("bearer", token)
        else:
            warn("`az` CLI is installed but not authenticated, or token fetch failed.")
            print(f"  {DIM}Try: {BOLD}az login --tenant 72f988bf-86f1-41af-91ab-2d7cd011db47{RESET}")

    # 3. Fall back to PAT entry.
    print()
    print(f"  {BOLD}How to create a PAT:{RESET}")
    print(f"    1. Open: {BLUE}https://dev.azure.com/MicrosoftIT/_usersSettings/tokens{RESET}")
    print(f"    2. Click {BOLD}+ New Token{RESET}.")
    print(f"    3. Name:          {CYAN}pr-sla-report{RESET}")
    print(f"    4. Organization:  {CYAN}All accessible organizations{RESET}")
    print(f"    5. Expiration:    {CYAN}90 days{RESET} (max)")
    print(f"    6. Scopes:        click {BOLD}Custom defined{RESET}, then check ONLY:")
    print(f"         • {GREEN}Code{RESET}        → {BOLD}Read{RESET}")
    print(f"         • {GREEN}Work Items{RESET}  → {BOLD}Read{RESET}")
    print(f"         • {GREEN}Identity{RESET}    → {BOLD}Read{RESET}")
    print(f"    7. Click {BOLD}Create{RESET}, then copy the token "
          f"(ADO shows it once).")
    if not az_present:
        print()
        print(f"  {DIM}Tip: install Azure CLI to avoid manual PATs entirely:{RESET}")
        print(f"  {DIM}    brew install azure-cli && az login{RESET}")
    print()

    pat = ask_secret("Paste your PAT here")
    ok(f"PAT received ({len(pat)} chars). It is NOT saved to disk.")
    return ("pat", pat)


# ---------------------------------------------------------------------------
# Prompt 4 — year/month
# ---------------------------------------------------------------------------
_MONTH_PATTERNS = [
    re.compile(r"^(\d{4})[/\-_.](\d{1,2})$"),  # 2026/05 or 2026-05
    re.compile(r"^(\d{1,2})[/\-_.](\d{4})$"),  # 05/2026
]


def _parse_month_input(raw: str) -> tuple[int, int] | None:
    for pat in _MONTH_PATTERNS:
        m = pat.match(raw)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            year, month = (a, b) if a > 12 else (b, a)
            if 1 <= month <= 12 and 2000 <= year <= 2100:
                return year, month
    return None


def prompt_month() -> tuple[dt.date, dt.date, str, str]:
    step(4, 6, "Reporting period")
    today = dt.date.today()
    # Default = previous calendar month.
    first_of_this = today.replace(day=1)
    last_of_prev = first_of_this - dt.timedelta(days=1)
    default = last_of_prev.strftime("%Y/%m")

    print(f"  Format: {CYAN}YYYY/MM{RESET}  (e.g. 2026/05 for May 2026)")
    print(f"  Press Enter to accept default ({BOLD}{default}{RESET} = last month).")
    print()

    while True:
        raw = ask_text("Year/Month to report on", default=default)
        parsed = _parse_month_input(raw)
        if not parsed:
            warn(f"Could not parse {raw!r}. Use YYYY/MM (e.g. 2026/05).")
            continue
        y, m = parsed
        if dt.date(y, m, 1) > today:
            warn(f"{y}/{m:02d} is in the future. Try again.")
            continue
        since = dt.date(y, m, 1)
        until = dt.date(y, m, calendar.monthrange(y, m)[1])
        label = since.strftime("%B %Y")
        slug = f"{y}-{m:02d}"
        ok(f"Window: {since} → {until}  ({label})")
        return since, until, label, slug


# ---------------------------------------------------------------------------
# Prompt 5 — sorting
# ---------------------------------------------------------------------------
def prompt_sort() -> tuple[str, bool]:
    step(5, 6, "Sort the results?")
    answer = ask_yes_no("Should the final results be sorted?", default=True)
    if not answer:
        ok("Default ordering: by Approved Date (ascending).")
        return "approved", False

    print()
    print(f"  Choose a sort key:")
    print(f"    {BOLD}1{RESET}) approved      — by Approved Date (default)")
    print(f"    {BOLD}2{RESET}) dev-complete  — by Dev Complete / PR Submitted Date")
    print(f"    {BOLD}3{RESET}) days          — by Days To Process (slowest = useful with desc)")
    print(f"    {BOLD}4{RESET}) author        — alphabetical by team member name")
    print()
    choice_map = {
        "1": "approved",
        "2": "dev-complete",
        "3": "days",
        "4": "author",
    }
    choice = ask_text(
        "Pick 1-4",
        default="1",
        validator=lambda x: x in choice_map,
        error_msg="Pick 1, 2, 3, or 4.",
    )
    sort_by = choice_map[choice]

    desc = ask_yes_no("Sort descending? (No = ascending)", default=False)
    ok(f"Sort: {BOLD}{sort_by}{RESET}{' (desc)' if desc else ' (asc)'}")
    return sort_by, desc


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Generate a PR SLA report interactively.")
    ap.add_argument(
        "--over-sla",
        action="store_true",
        help="Only include PRs taking more than 2 business days from last commit to approval.",
    )
    args = ap.parse_args(argv)
    banner("PR SLA — Interactive Monthly Report")
    if args.over_sla:
        ok("Filter: only PRs taking more than 2 business days from last commit to approval.")

    # Timezone notice — printed FIRST so the user always sees it.
    from generate_report import report_tz_name
    print()
    print(f"  {YELLOW}{BOLD}NOTE{RESET}  All dates in the generated report are "
          f"shown in {BOLD}{report_tz_name()} (PT){RESET}.")
    print(f"        ADO returns timestamps in UTC; we convert to PT before "
          f"truncating to a date so values match what")
    print(f"        Microsoft-internal reports usually use. Override with "
          f"--tz <IANA name> on the command line.")
    print()

    if not CONFIG_XLSX.exists():
        fail(f"Config file not found: {CONFIG_XLSX}")
        return 1

    # 0. Prerequisites — must come before everything else
    prompt_prerequisites()

    try:
        repos, team = load_config(CONFIG_XLSX)
    except Exception as e:
        fail(f"Could not load config: {e}")
        return 1

    if not repos:
        fail("No repos configured. Add at least one row in the Repos sheet.")
        return 1
    if not team:
        fail("No team members configured. Add at least one row in the Team sheet.")
        return 1

    # 1. Confirm repos
    prompt_confirm_repos(repos)

    # 2. Confirm team
    prompt_confirm_team(team)

    # 3. Credentials
    auth_kind, auth_value = prompt_auth()

    # 4. Month
    since, until, month_label, slug = prompt_month()

    # 5. Sort
    sort_by, sort_desc = prompt_sort()

    # ---- Execute -----------------------------------------------------------
    banner("Generating report")
    print(f"  {DIM}This can take 30s–2min depending on team size + PR count.{RESET}")
    print(f"  {DIM}Each PR triggers ~3 ADO API calls (commits, threads, work items).{RESET}")
    print()

    try:
        if args.over_sla:
            configure_holidays(since, until)
        if auth_kind == "bearer":
            client = AdoClient(bearer_token=auth_value)
        else:
            client = AdoClient(pat=auth_value)
        data = extract_rows(client, repos, team, since, until, over_sla=args.over_sla)
    except Exception as e:
        print()
        fail(f"Extraction failed: {e}")
        return 2

    if args.over_sla:
        month_label += "; >2 business days"
        slug += "_over_sla"

    total = sum(len(v) for v in data.values())
    print()
    print(hr())
    print(f"  {BOLD}Extraction summary{RESET}")
    for repo, rows in data.items():
        marker = GREEN + "●" + RESET if rows else DIM + "○" + RESET
        print(f"    {marker} {repo.short_name:<40} {len(rows):>3} PR(s)")
    print(f"  {BOLD}Total: {total} PR(s){RESET}")

    if total == 0:
        if args.over_sla:
            warn("No PRs matched the >2 business-day filter in the reporting window.")
        else:
            warn("No PRs matched. Check the team UPNs and the reporting window.")
        warn("A deck will still be generated with empty tables.")

    out_path = OUTPUT_DIR / f"PR_SLA_{slug}.pptx"
    try:
        build_deck(
            data,
            month_label,
            out_path,
            sort_by=sort_by,
            sort_desc=sort_desc,
        )
    except Exception as e:
        print()
        fail(f"PowerPoint generation failed: {e}")
        return 3

    # ---- Done --------------------------------------------------------------
    banner("✓ Report generated")
    abs_path = out_path.resolve()
    size_kb = abs_path.stat().st_size / 1024
    print(f"  {GREEN}{BOLD}SUCCESS{RESET}")
    print(f"  File:  {BOLD}{abs_path}{RESET}")
    print(f"  Size:  {size_kb:.1f} KB")
    print(f"  Slides: {len(repos)}  (one per repo)")
    print()
    print(f"  {DIM}Open with:{RESET}")
    if sys.platform == "darwin":
        print(f"    open '{abs_path}'")
    elif sys.platform.startswith("win"):
        print(f"    start \"\" \"{abs_path}\"")
    else:
        print(f"    xdg-open '{abs_path}'")
    print()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print(f"\n  {YELLOW}Cancelled by user.{RESET}")
        sys.exit(130)
