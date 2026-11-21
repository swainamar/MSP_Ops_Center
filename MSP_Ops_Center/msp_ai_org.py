#!/usr/bin/env python3
"""MSP Ops Center assistant: answers about teams, members, customers and services.

Reads only the person's own view of the team setup (msp_org.view), so the same rules apply as on the Teams page:
members and leads see their own teams in full, administrators see every team, and nobody else sees a team's
member list. Used by msp_ai.py; returns None when a question is not about teams or customers.
"""
from __future__ import annotations

import re

ROLE_ORDER = {"Lead": 0, "Engineer": 1, "Viewer": 2}
ROLE_WORD = {"Lead": "Team leads", "Engineer": "Engineers", "Viewer": "Operators (read-only)"}


def _find(names, q: str):
    """The longest name that appears in the question as whole words (Unix SA before SA)."""
    ql = " " + re.sub(r"[^a-z0-9]+", " ", q.lower()) + " "
    hits = [n for n in names if (" " + re.sub(r"[^a-z0-9]+", " ", n.lower()).strip() + " ") in ql]
    return max(hits, key=len) if hits else None


def answer(q: str, org: dict | None, actor: dict, make_answer):
    """Return an Answer for team or customer questions, else None. make_answer() creates an empty Answer."""
    if not org:
        return None
    ql = q.lower()
    visible = {t["name"]: t for t in org.get("teams", [])}
    every = org.get("team_names", {})
    customers = {c["name"]: c for c in org.get("customers", [])}
    team = _find(set(every.values()) | set(visible), q)
    cust = _find(customers, q)
    ans = make_answer()

    def team_src(t):
        return ans.source("team", f"{t['name']} team", f"{len(t['members'])} people", "teams")

    # "my teams", "which team am I in"
    if re.search(r"\b(my teams?|which teams? am i|what teams? am i|am i in)\b", ql):
        mine = [t for t in org.get("teams", []) if any(m["username"].lower() == actor["user"].lower() for m in t["members"])]
        if not mine:
            return ans.p("You are not in a team yet. A team lead can add you on the Teams page (Administration > Teams).")
        ans.p("You are in:")
        ans.items([f"{t['name']}" + (" (you lead it)" if t.get("you_lead") else "") + f" {team_src(t)}" for t in mine])
        return ans

    asks_team = bool(re.search(r"\b(team|teams|members?|who|lead|leads|people|staff|engineers?|operators?)\b", ql))
    if team and asks_team:
        t = visible.get(team)
        if t is None:
            ans.found = False
            leads_hint = "its team lead or an administrator"
            return ans.p(f"You are not in the {team} team, so its member list is not shown to you. "
                         f"Ask {leads_hint}, or see who is on shift for that team on the Shift roster page.")
        src = team_src(t)
        if re.search(r"\blead(s|er|ers)?\b|\bwho runs\b|\bmanager\b", ql) and not re.search(r"\bmembers?\b", ql):
            ans.p(f"{t['name']} is led by {', '.join(l['name'] for l in t['leads']) or 'nobody yet'} {src}.")
            return ans
        ans.p(f"The {t['name']} team has {len(t['members'])} people {src}:")
        groups = {}
        for m in t["members"]:
            groups.setdefault(m["role"] if m["role"] in ROLE_ORDER else "Engineer", []).append(m)
        for role in sorted(groups, key=lambda r: ROLE_ORDER[r]):
            ans.p(f"{ROLE_WORD[role]}:")
            ans.items([f"{m['name']} ({m['username']})" + (" (lead)" if any(l["username"] == m["username"] for l in t["leads"]) else "")
                       for m in sorted(groups[role], key=lambda x: x["name"])])
        served = [c["name"] for c in org.get("customers", []) if t["id"] in c.get("teams", [])]
        if served:
            primary = [c["name"] for c in org.get("customers", []) if c.get("primary_team") == t["id"]]
            ans.p(f"Customers served: {', '.join(served)}." + (f" Primary team for {', '.join(primary)}." if primary else ""))
        if t.get("can_manage"):
            ans.p("To add or remove people, open Administration > Teams and use Add member on this team.")
        return ans

    asks_cust = bool(re.search(r"\b(serv\w*|team|teams|primary|who|projects?|customers?|works? on)\b", ql))
    if cust and asks_cust:
        c = customers[cust]
        src = ans.source("customer", c["name"], f"{len(c['projects'])} services", "customers")
        names = org.get("team_names", {})
        ans.p(f"{c['name']}{' (' + c['code'] + ')' if c.get('code') else ''} {src}:")
        ans.items([f"Teams serving it: {', '.join(names.get(x, x) for x in c['teams']) or 'none yet'}",
                   f"Primary team (owns the core CI data): {names.get(c.get('primary_team'), 'not set')}",
                   f"Services: {', '.join(p['name'] for p in c['projects']) or 'none yet'}"])
        people = {}
        for t in org.get("teams", []):
            for m in t["members"]:
                for p in m.get("projects", []):
                    people.setdefault(p, set()).add(m["name"])
        lines = [f"{p['name']}: {', '.join(sorted(people[p['id']]))}" for p in c["projects"] if people.get(p["id"])]
        if lines:
            ans.p("People assigned (from your teams):")
            ans.items(lines)
        return ans
    return None
