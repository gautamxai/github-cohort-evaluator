#!/usr/bin/env python3
"""
Step 1: Discover all repositories and teams in a GitHub organization.
Reads GH_TOKEN from the environment (never hardcoded).
Author: Gautam Gupta (https://www.linkedin.com/in/higautam/)
"""

import argparse
import json
import os
import re
import ssl
import sys
import urllib.request


def api_get(url: str, token: str) -> list | dict:
    ctx = ssl.create_default_context()
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "github-cohort-evaluator",
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8"))


def normalize_team_id(name: str) -> str | None:
    """Strict regex normalization so TEAM001 never collides with TEAM0010."""
    m = re.search(r"TEAM[_-]?0*([0-9]+)", name.upper())
    return f"TEAM{int(m.group(1)):03d}" if m else None


def main() -> None:
    parser = argparse.ArgumentParser(description="Discover GitHub org repos and teams.")
    parser.add_argument("--org", required=True, help="GitHub organization name")
    parser.add_argument("--out", default="org_audit.json", help="Output JSON path")
    args = parser.parse_args()

    token = os.environ.get("GH_TOKEN")
    if not token:
        print("Error: GH_TOKEN environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    repos = api_get(f"https://api.github.com/orgs/{args.org}/repos?per_page=100", token)
    teams = api_get(f"https://api.github.com/orgs/{args.org}/teams?per_page=100", token)

    summary = {
        "organization": args.org,
        "repo_count": len(repos),
        "team_count": len(teams),
        "repos": [
            {
                "name": r["name"],
                "normalized_team": normalize_team_id(r["name"]),
                "private": r["private"],
                "size_kb": r["size"],
                "default_branch": r["default_branch"],
                "pushed_at": r["pushed_at"],
                "html_url": r["html_url"],
            }
            for r in repos
        ],
        "teams": [{"name": t["name"], "slug": t["slug"]} for t in teams],
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)
    print(f"Saved {len(repos)} repos and {len(teams)} teams to {args.out}")


if __name__ == "__main__":
    main()
