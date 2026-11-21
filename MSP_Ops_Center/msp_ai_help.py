#!/usr/bin/env python3
"""MSP Ops Center assistant: built-in how-to guides.

Each guide has a title, the words people use for it, who can do it, where it is, and the steps.
Used by msp_ai.py for "how do I ...", "where can I ...", "how to configure ..." questions. To add or correct a
guide, edit GUIDES below; the tests in test_msp_ai_org.py check that every guide is reachable.
"""
from __future__ import annotations

import re

# title, keywords, who can do it, page (link), steps
GUIDES = [
    ("Add a person to your team", "add member team join person engineer user account create new user onboard",
     "Team leads (for their teams) and administrators", "teams",
     ["Open Administration > Teams.", "On your team, select Add member.",
      "Choose An existing user, or A new user to create an account (Engineer or Viewer).",
      "For a new local account, give the one-time password to the person; they choose their own at first sign-in."]),
    ("Create a team and name its leads", "create team new team lead leads assign lead",
     "Administrators", "teams",
     ["Give the person the Lead role first: Admin > Users.", "Open Administration > Teams and select New team.",
      "Enter the name and description, tick the team leads, and save."]),
    ("Add a customer, its services and its teams", "customer add customer service services project primary team serve",
     "Administrators", "customers",
     ["Open Service relationships > Customers and services and select New customer.",
      "Select Edit on the customer to tick the teams that serve it and choose the primary team.",
      "Select Add service for each service you deliver to the customer."]),
    ("Assign people to services", "assign service services project projects member work on",
     "Team leads and administrators", "teams",
     ["Open Administration > Teams.", "On the person's row, select Services.",
      "Tick the services of customers your team serves, and save."]),
    ("Request leave", "leave holiday time off vacation request annual sick",
     "Engineers (for themselves) and team leads", "roster",
     ["Open Operations > Shift roster and pick your team's roster.", "Select Request leave, choose the dates and type.",
      "Choose the approver if asked. The request waits for them unless the roster records leave straight away."]),
    ("Swap or cover a shift", "swap cover exchange shift trade change shift",
     "Engineers (their own shifts) and team leads", "roster",
     ["Open Operations > Shift roster and pick your roster.", "Select Swap or cover, choose the day and the other person.",
      "The other person accepts on My work, then the approver decides. The roster checks rest hours first."]),
    ("Approve leave or swaps", "approve reject decide request pending approval",
     "The assigned approver (or any lead of the team if none is set)", "mywork",
     ["Open My work: requests waiting for you are at the top of Your shifts.",
      "Select Approve or Reject. You can also do this on the Shift roster page."]),
    ("Change a roster's shifts, rules or holidays", "roster settings configure shift times rotation rules holidays on-call oncall",
     "A lead of the roster's team", "roster",
     ["Open Operations > Shift roster and pick the roster.", "Select Settings (shifts, rotation, rules, approvers, on-call) or Holidays.",
      "Click a cell (by person) or a shift box (by shift) to change one day."]),
    ("Import a roster sheet", "import roster sheet excel upload template roster file",
     "Team leads", "roster",
     ["Open Operations > Shift roster and select Template to download a blank month, then fill it in.",
      "Select Import sheet, choose the file, give the roster a name and your team, and import.",
      "Read the notes it shows: possible duplicate names, days not covered, minimum staff."]),
    ("Open a shift in the shift log", "open shift log shift log new shift handover overview carry",
     "Engineers and team leads", "shiftlog",
     ["Open Operations > Shift log. It shows your team's log (switch team at the top if you are in two).",
      "Pick the date and shift and select Open shift. Tick Carry rows forward to copy open items."]),
    ("Write a shift handover", "handover hand over write handover outgoing incoming",
     "Engineers and team leads", "Shift_Transitions",
     ["Open Operations > Shift handovers and select Add handover.", "Only your team sees it."]),
    ("Add a configuration item and share it", "add ci configuration item cmdb share sharing shared core fully team only owner",
     "Engineers and team leads of the owning team", "Asset_Master",
     ["Open Knowledge > CMDB and select Add CI. The owner team defaults to your team that serves the customer.",
      "Set Sharing: Shared core (name, hostname, IP and remote access), Fully shared, or Team only.",
      "Set Shared with teams (empty means every team serving the customer) and Shared fields.",
      "Each team can add its own notes in its '(team only)' column on CIs shared with it."]),
    ("Write a runbook entry and share it", "runbook entry fix known issue write share runbook",
     "Engineers and team leads", "runbook",
     ["Open Knowledge > Runbook and add an entry for the CI.",
      "It belongs to your team; list other teams in Shared with teams to let them read it."]),
    ("Open an SSH or RDP session", "ssh rdp terminal connect remote session console",
     "Engineers and team leads", "terminals",
     ["Open Operations > Terminals, or the SSH button on a CI in the CMDB.",
      "You can connect to CIs your team owns or that are shared with your team."]),
    ("Move the shift log from before teams into your team", "move old shift log legacy before teams claim",
     "A lead of the team", "shiftlog",
     ["Open Operations > Shift log as a team lead.", "Select Move them into <your team> on the notice at the top."]),
    ("Set up the local AI model", "ai model local ollama configure assistant llm gpu",
     "Administrators", "admin",
     ["Install Ollama on your server or laptop and download a model (for example a 7 to 8 billion parameter one).",
      "Open Admin > Settings: set Assistant to model, the model server address (this machine or your private network) and the model name.",
      "Select Test the model server, then Save. Public addresses are refused."]),
    ("Change your password or theme", "password change password theme dark light",
     "Everyone", "mywork",
     ["Use the buttons at the bottom of the sidebar: the sun switches light and dark.",
      "Change your password from your account menu; administrators can reset it in Admin > Users."]),
]
CUE = re.compile(r"\b(how|where|configure|set ?up|steps?|can i|do i|add|create|change|request|approve|import|share|open|write|move|assign)\b")
STOP = set("how do i to the a an of for in on and or can where is are my what which with please".split())


def words(text: str) -> set:
    return {w for w in re.findall(r"[a-z0-9]+", text.lower()) if w not in STOP and len(w) > 1}


def best_guide(q: str):
    """The guide that best fits the question, or None."""
    if not CUE.search(q.lower()):
        return None
    qw = words(q)
    scored = []
    for g in GUIDES:
        kw = words(g[0] + " " + g[1])
        hit = len(qw & kw)
        if hit:
            scored.append((hit, -len(kw), g))
    if not scored:
        return None
    scored.sort(reverse=True)
    return scored[0][2] if scored[0][0] >= 2 or len(qw) <= 2 else None


def answer(q: str, make_answer):
    g = best_guide(q)
    if g is None:
        return None
    title, _kw, who, page, steps = g
    ans = make_answer()
    src = ans.source("help", title, who, page)
    ans.p(f"{title} {src}")
    ans.p(f"Who can do it: {who}.")
    ans.items([f"{i}. {s}" for i, s in enumerate(steps, 1)])
    return ans
