#!/usr/bin/env python3
"""MSP Ops Center demo world: everything fictional, generated the same way every time.

10 customers, 5 teams of 10+ people (leads, engineers, read-only operators, some people in two teams),
one roster per team, about 80 CIs carrying an ownership-and-sharing model, and runbook, incidents,
handovers, bulletins, execution plans (one shared across three teams) and a week of shift log.

Used by msp_demo.py. Nothing here talks to a server; it only builds Python data.
"""
from __future__ import annotations

import datetime as dt
import random

RNG_SEED = 2026
TODAY = dt.date(2026, 11, 13)  # the demo story is told around this date

TEAMS = {  # name: (description, roster template, region)
    "Unix SA": ("Linux and Solaris servers, storage and backup", "24x7"),
    "Windows": ("Windows servers, Active Directory, Citrix and VMware", "24x7"),
    "DBA": ("Oracle and SQL Server databases", "24x7"),
    "Boomi": ("Boomi integration platform", "business"),
    "NOC": ("24x7 monitoring, first response and network", "24x7"),
}
# Each team: leads first, then engineers, then operators (read-only Viewers). First team listed = home team (their roster).
PEOPLE = [
    # Unix SA
    ("arun.kumar", "Arun Kumar", "Lead", ["Unix SA"]), ("meera.pillai", "Meera Pillai", "Lead", ["Unix SA"]),
    ("suresh.babu", "Suresh Babu", "Engineer", ["Unix SA"]), ("nikhil.jain", "Nikhil Jain", "Engineer", ["Unix SA"]),
    ("farah.ali", "Farah Ali", "Engineer", ["Unix SA", "NOC"]), ("rohit.das", "Rohit Das", "Engineer", ["Unix SA"]),
    ("lena.fox", "Lena Fox", "Engineer", ["Unix SA"]), ("vivek.nanda", "Vivek Nanda", "Engineer", ["Unix SA"]),
    ("chen.li", "Chen Li", "Engineer", ["Unix SA", "DBA"]), ("kavya.menon", "Kavya Menon", "Engineer", ["Unix SA"]),
    ("tom.berg", "Tom Berg", "Viewer", ["Unix SA"]),
    # Windows
    ("karthik.reddy", "Karthik Reddy", "Lead", ["Windows"]), ("anita.desai", "Anita Desai", "Lead", ["Windows"]),
    ("rahul.verma", "Rahul Verma", "Engineer", ["Windows", "NOC"]), ("mark.evans", "Mark Evans", "Engineer", ["Windows"]),
    ("deepa.krishnan", "Deepa Krishnan", "Engineer", ["Windows"]), ("sanjay.patel", "Sanjay Patel", "Engineer", ["Windows"]),
    ("olivia.chen", "Olivia Chen", "Engineer", ["Windows"]), ("imran.sheikh", "Imran Sheikh", "Engineer", ["Windows", "Boomi"]),
    ("pooja.nair", "Pooja Nair", "Engineer", ["Windows"]), ("varun.mehta", "Varun Mehta", "Engineer", ["Windows"]),
    ("ana.ito", "Ana Ito", "Viewer", ["Windows"]),
    # DBA
    ("priya.sharma", "Priya Sharma", "Lead", ["DBA"]), ("divya.menon", "Divya Menon", "Lead", ["DBA"]),
    ("sneha.iyer", "Sneha Iyer", "Engineer", ["DBA"]), ("ashok.rao", "Ashok Rao", "Engineer", ["DBA"]),
    ("neha.gupta", "Neha Gupta", "Engineer", ["DBA"]), ("joseph.mathew", "Joseph Mathew", "Engineer", ["DBA"]),
    ("swati.kulkarni", "Swati Kulkarni", "Engineer", ["DBA"]), ("daniel.kim", "Daniel Kim", "Engineer", ["DBA"]),
    ("ritu.saxena", "Ritu Saxena", "Engineer", ["DBA"]), ("manoj.pillai", "Manoj Pillai", "Engineer", ["DBA"]),
    ("grace.lee", "Grace Lee", "Viewer", ["DBA"]),
    # Boomi
    ("vikram.rao", "Vikram Rao", "Lead", ["Boomi"]), ("omar.khan", "Omar Khan", "Engineer", ["Boomi"]),
    ("lakshmi.iyer", "Lakshmi Iyer", "Engineer", ["Boomi"]), ("arjun.bose", "Arjun Bose", "Engineer", ["Boomi"]),
    ("sara.thomas", "Sara Thomas", "Engineer", ["Boomi"]), ("kiran.shetty", "Kiran Shetty", "Engineer", ["Boomi"]),
    ("maya.singh", "Maya Singh", "Engineer", ["Boomi"]), ("ravi.nair", "Ravi Nair", "Engineer", ["Boomi"]),
    ("tanvi.joshi", "Tanvi Joshi", "Engineer", ["Boomi"]), ("ben.cole", "Ben Cole", "Viewer", ["Boomi"]),
    # NOC
    ("anjali.nair", "Anjali Nair", "Lead", ["NOC", "Windows"]), ("harish.kumar", "Harish Kumar", "Lead", ["NOC"]),
    ("dinesh.yadav", "Dinesh Yadav", "Engineer", ["NOC"]), ("aisha.khan", "Aisha Khan", "Engineer", ["NOC"]),
    ("george.paul", "George Paul", "Engineer", ["NOC"]), ("shruti.rao", "Shruti Rao", "Engineer", ["NOC"]),
    ("naveen.reddy", "Naveen Reddy", "Engineer", ["NOC"]), ("emma.wright", "Emma Wright", "Engineer", ["NOC"]),
    ("prakash.iyer", "Prakash Iyer", "Viewer", ["NOC"]), ("leah.morris", "Leah Morris", "Viewer", ["NOC"]),
    # no team
    ("sam.wilson", "Sam Wilson", "Viewer", []),
]
ROLE_NOTE = {"Lead": "Team lead", "Engineer": "Engineer", "Viewer": "Operator (read-only)"}

# name, code, industry, projects, teams serving it, primary team
CUSTOMERS = [
    ("Acme Corp", "ACME", "Manufacturing", ["ERP platform", "Integration hub"], ["Unix SA", "DBA", "Boomi", "NOC"], "Unix SA"),
    ("Globex", "GBX", "Retail", ["Retail data platform", "Cloud and web"], ["Windows", "DBA", "NOC"], "Windows"),
    ("Initech", "INI", "Banking", ["Core banking"], ["Unix SA", "DBA"], "Unix SA"),
    ("Umbrella Health", "UMB", "Healthcare", ["Clinical workspace", "Identity"], ["Windows", "NOC"], "Windows"),
    ("Bluepeak Logistics", "BPL", "Logistics", ["Partner EDI"], ["Boomi"], "Boomi"),
    ("Harbourline Retail", "HRL", "Retail", ["Store systems", "Online shop", "Integration"], ["Unix SA", "Windows", "DBA", "Boomi", "NOC"], "Unix SA"),
    ("Cobalt Labs", "CBL", "Research", ["Research data"], ["DBA", "Unix SA"], "DBA"),
    ("Greenfield Foods", "GFF", "Food", ["Plant systems", "Supplier integration"], ["Windows", "Boomi"], "Windows"),
    ("Skyway Air", "SKY", "Aviation", ["Flight ops", "Network"], ["Unix SA", "NOC"], "Unix SA"),
    ("Lumen Media", "LUM", "Media", ["Streaming platform", "Billing"], ["Windows", "DBA", "Boomi"], "Windows"),
]

SHARED_CORE = "CI name, Hostname, IP address, Remote access (SSH/RDP)"
CORE_COLUMNS = ["Customer", "Serial No", "CI name", "CI type", "Hostname / FQDN", "IP address", "Environment", "Platform type",
                "Hypervisor / cloud account", "OS", "OS version", "Application", "Role", "CPU", "Memory", "Root disk", "Data disk",
                "Data centre / rack", "Criticality", "Patch group", "Maintenance window", "Backup policy", "Monitoring",
                "Warranty end", "OS end of life", "Upstream dependencies", "Downstream dependencies", "Special care", "Owner",
                "Contacts", "CI status"]
# Sharing model columns (custom CMDB columns). Enforced once isolation is switched on.
SHARE_COLUMNS = ["Owner team", "Sharing", "Shared with teams", "Shared fields",
                 "Unix SA details (team only)", "Windows details (team only)", "DBA details (team only)",
                 "Boomi details (team only)", "NOC details (team only)"]


def members(team):
    return [p for p in PEOPLE if team in p[3]]


def home_members(team):
    return [p for p in PEOPLE if p[3] and p[3][0] == team]


def leads(team):
    """A lead leads their home team (first listed); in any other team they are a member."""
    return [p for p in members(team) if p[2] == "Lead" and p[3][0] == team]


def engineers(team):
    return [p for p in members(team) if p[2] == "Engineer"]


def build(rng: random.Random | None = None) -> dict:
    """The whole demo world as plain data."""
    rng = rng or random.Random(RNG_SEED)  # nosec B311 - repeatable demo data, not security
    cis, runbook, incidents, plans, handovers, shiftlog = [], [], [], [], [], []
    incident_team, plan_teams = {}, {}
    customers = {c[0]: c for c in CUSTOMERS}

    def eng(team):
        return rng.choice(engineers(team))[1]

    def lead(team):
        return leads(team)[0][1]

    # ---------------------------------------------------------------- CMDB with ownership and sharing
    n_ip = {}
    for cname, code, industry, projects, teams, primary in CUSTOMERS:
        low = code.lower()
        net = 10 + CUSTOMERS.index(customers[cname]) * 3
        n_ip[cname] = 10

        def ip():
            n_ip[cname] += 1
            return f"10.{net}.{n_ip[cname] // 250}.{n_ip[cname] % 250}"

        def ci(name, typ, owner, os_, osv, appn, role, plat="Virtualization", crit="High", care="", up="", down="",
               sharing="Shared core", shared_with=None, team_only=None, status="Active", env="Production", cpu="8 vCPU", mem="32 GB"):
            others = [t for t in teams if t != owner]
            sw = shared_with if shared_with is not None else others
            row = {"Customer": cname, "Serial No": f"{code}-{rng.randint(10000, 99999)}", "CI name": name, "CI type": typ,
                   "Hostname / FQDN": f"{name}.{low}.example", "IP address": ip(), "Environment": env, "Platform type": plat,
                   "Hypervisor / cloud account": "VMware vSphere 8 / cluster PROD" if plat == "Virtualization" else ("Azure / prod" if plat == "Cloud" else ""),
                   "OS": os_, "OS version": osv, "Application": appn, "Role": role, "CPU": cpu, "Memory": mem,
                   "Root disk": "100 GB", "Data disk": "500 GB" if typ in ("Database", "Server") else "",
                   "Data centre / rack": f"DC{1 + CUSTOMERS.index(customers[cname]) % 3} / R{rng.randint(1, 20):02d}", "Criticality": crit,
                   "Patch group": f"{owner}-Monthly", "Maintenance window": rng.choice(["Sat 22:00-02:00", "Sun 22:00-02:00", "Wed 20:00-22:00"]),
                   "Backup policy": "Nightly" if typ != "Network device" else "Config backup daily", "Monitoring": "Enabled",
                   "Warranty end": (TODAY + dt.timedelta(days=rng.choice([40, 75, 200, 420, 700]))).isoformat() if plat == "Physical" else "",
                   "OS end of life": (TODAY + dt.timedelta(days=rng.choice([60, 300, 900, 1500]))).isoformat(),
                   "Upstream dependencies": up, "Downstream dependencies": down, "Special care": care,
                   "Owner": owner, "Contacts": f"{lead(owner)} (team lead)", "CI status": status,
                   "Owner team": owner, "Sharing": sharing if sw or sharing == "Team only" else "Team only",
                   "Shared with teams": ", ".join(sw) if sharing != "Team only" else "",
                   "Shared fields": ("All" if sharing == "Fully shared" else SHARED_CORE) if sharing != "Team only" and sw else ""}
            for t, txt in (team_only or {}).items():
                row[f"{t} details (team only)"] = txt
            cis.append(row)
            return name

        host = dbhost = None
        if "Unix SA" in teams:
            host = ci(f"{low}-lnx-app01", "Server", "Unix SA", "Oracle Linux", "8.9", f"{projects[0]} application tier", "App server",
                      care="Drain from the load balancer before reboot.", down="Users",
                      team_only={"Unix SA": "Kernel 5.15 UEK; tuned profile throughput-performance; /u01 on LVM"})
            ci(f"{low}-lnx-batch01", "Server", "Unix SA", "Ubuntu", "22.04", "Batch scheduler", "Batch server", crit="Medium",
               sharing="Team only", team_only={"Unix SA": "Control-M agent; cron jobs owned by svc_batch"})
            ci(f"{low}-lnx-app02", "Server", "Unix SA", "Oracle Linux", "8.9", f"{projects[0]} application tier", "App server (node 2)",
               team_only={"Unix SA": "Clustered with app01 behind F5; patch one node at a time"})
            dbhost = ci(f"{low}-lnx-db01", "Server", "Unix SA", "Oracle Linux", "8.9", "Database host", "DB host", plat="Physical",
                        crit="Critical", cpu="2 x 24 cores", mem="512 GB", care="Coordinate every reboot with the DBA team.",
                        team_only={"Unix SA": "HugePages 200 GB; multipath to SAN; firmware 2.4"})
        if "Windows" in teams:
            host = host or ci(f"{low}-win-web01", "Application", "Windows", "Windows Server", "2022", "IIS 10", "Web front end",
                              care="Recycle app pools one at a time.", team_only={"Windows": "App pools: Orders, Portal; cert renewal via internal PKI"})
            ci(f"{low}-win-file01", "Server", "Windows", "Windows Server", "2019", "File services (DFS)", "File server", crit="Medium",
               team_only={"Windows": f"DFS namespace {low}-files/shares; shadow copies 07:00 and 12:00"})
            ci(f"{low}-win-dc01", "Server", "Windows", "Windows Server", "2022", "Active Directory, DNS", "Domain controller", crit="Critical",
               care="Patch one domain controller per night.", sharing="Fully shared",
               team_only={"Windows": "FSMO roles: PDC, RID; replication check with repadmin"})
            if "Unix SA" not in teams:
                dbhost = ci(f"{low}-win-db01", "Server", "Windows", "Windows Server", "2022", "Database host", "DB host", crit="Critical",
                            cpu="16 vCPU", mem="128 GB", care="Fail over databases before patching.", team_only={"Windows": "Cluster CL-DB01; quorum disk Q:"})
        if "DBA" in teams:
            dbh = dbhost or ci(f"{low}-db-host01", "Server", primary, "Oracle Linux", "8.9", "Database host", "DB host", crit="Critical")
            engine = rng.choice(["Oracle 19c", "SQL Server 2022", "PostgreSQL 15"])
            ci(f"{low}-db-{'ora' if 'Oracle' in engine else 'sql'}01", "Database", "DBA", "", "", engine, f"{projects[0]} database",
               crit="Critical", up=dbh, care="Never restart during month-end batch.", sharing="Shared core",
               shared_with=[t for t in teams if t in ("Unix SA", "Windows", "NOC")],
               team_only={"DBA": f"{engine}; SID {code}PRD; RMAN daily incremental, archivelog every 30 min; DG standby {low}-db-dr01"})
            ci(f"{low}-db-test01", "Database", "DBA", "", "", engine, "Test database", env="Test", crit="Low", up=dbh,
               sharing="Team only", team_only={"DBA": "Refreshed weekly from production (masked)."})
        if "Boomi" in teams:
            ci(f"{low}-boomi01", "Middleware", "Boomi", "RHEL", "9.3", "Boomi Molecule", "Integration node", crit="High",
               care="Drain the node before reboot; never restart both nodes together.", shared_with=[t for t in teams if t in ("Unix SA", "Windows", "NOC")],
               team_only={"Boomi": "Atom 2.4; listeners: EDI 850/855/856; heap 4 GB"})
            ci(f"{low}-boomi02", "Middleware", "Boomi", "RHEL", "9.3", "Boomi Molecule", "Integration node 2",
               team_only={"Boomi": "Joined to molecule MOL-1; process schedules run on node 1 only"})
        if "NOC" in teams:
            ci(f"{low}-fw01", "Network device", "NOC", "FortiOS", "7.4.3", "Perimeter firewall", "Internet edge", plat="Physical",
               crit="Critical", care="HA pair: upgrade the passive unit first.", sharing="Fully shared",
               team_only={"NOC": "HA active/passive; monitored via SNMP v3; alert profile EDGE"})
            ci(f"{low}-lb01", "Network device", "NOC", "F5 TMOS", "17.1", "Load balancer", "Application delivery", plat="Physical",
               crit="High", care="Disable pool members before maintenance.", team_only={"NOC": "Pools: app, web; health monitor HTTP 200"})
            ci(f"{low}-sw01", "Network device", "NOC", "Cisco IOS-XE", "17.9", "Core switch", "Core", plat="Physical", crit="High",
               team_only={"NOC": "VLANs 10,20,30; uplinks Po1/Po2"})
    by_customer = {}
    for c in cis:
        by_customer.setdefault(c["Customer"], []).append(c)

    # ---------------------------------------------------------------- runbook, incidents
    cats = ["Known issue", "Change", "Maintenance", "Health check", "Workaround", "Incident", "Information"]
    stories = {
        "Server": [("Filesystem /u01 fills after log rotation fails", "logrotate config was overwritten by a package update.",
                    "Restore /etc/logrotate.d/app and run logrotate -f.", "df -h /u01\nlogrotate -f /etc/logrotate.d/app")],
        "Database": [("Archiver stuck when the FRA is full", "Archive log backups failed after a credential change.",
                      "Re-register the backup credential, back up archive logs, then purge.", "rman target / cmdfile=arch_backup.rman"),
                     ("Long-running query blocks month-end batch", "A missing index on the orders table.",
                      "Create the index online during the window.", "CREATE INDEX ix_orders_cust ON orders(customer_id) ONLINE;")],
        "Middleware": [("EDI listener stops after Java heap exhaustion", "A large 850 batch exhausted the heap.",
                        "Raise the heap to 4 GB and restart the listener only.", "grep -i outofmemory /opt/boomi/logs/*.log")],
        "Application": [("App pool stops after client certificate expiry", "Rapid-fail protection tripped.",
                         "Renew the certificate, then start the app pool.", "Start-WebAppPool Orders")],
        "Network device": [("IPsec tunnel to supplier flapping", "Dead-peer detection timers did not match.",
                            "Align DPD timers with the supplier.", "diagnose vpn tunnel list")],
    }
    tick = 104300
    for cname, rows in by_customer.items():
        cust = customers[cname]
        for row in rows:
            for title, note, sol, cmd in stories.get(row["CI type"], [])[:1 + (row["CI type"] == "Database")]:
                if rng.random() < 0.55:
                    owner = row["Owner team"]
                    ki = "Yes" if rng.random() < 0.35 else "No"
                    runbook.append({"Customer": cname, "Environment": row["Environment"], "CI": row["CI name"],
                                    "Date": (TODAY - dt.timedelta(days=rng.randint(1, 40))).isoformat(),
                                    "Category": "Known issue" if ki == "Yes" else rng.choice(cats[1:]), "Title": title, "Note": note,
                                    "Solution": sol, "Commands": cmd, "Ticket ref": f"INC-{tick}", "Engineer": eng(owner),
                                    "Tags": owner.lower().replace(" ", "-"), "Known issue": ki, "Owner team": owner,
                                    # known issues are shared with the other teams serving this customer; the rest stay with the owner
                                    "Shared with teams": ", ".join(t for t in cust[4] if t != owner) if ki == "Yes" else ""})
                    if rng.random() < 0.6:
                        pri = rng.choice(["P1", "P2", "P2", "P3", "P3", "P4"])
                        st = rng.choice(["New", "In Progress", "In Progress", "On Hold", "Resolved", "Closed"])
                        opened = dt.datetime.combine(TODAY - dt.timedelta(days=rng.randint(0, 9)), dt.time(rng.randint(0, 23), rng.choice([5, 20, 40])))
                        incident_team[f"INC-{tick}"] = owner
                        incidents.append({"Ticket ID": f"INC-{tick}", "Customer": cname, "CI": row["CI name"], "Priority": pri, "Status": st,
                                          "Summary": f"{title} on {row['CI name']}", "Action": sol, "Reason": note, "Technical tracking": cmd,
                                          "Engineer": eng(owner), "Opened at": opened.strftime("%Y-%m-%d %H:%M"),
                                          "Resolved at": (opened + dt.timedelta(hours=rng.randint(1, 6))).strftime("%Y-%m-%d %H:%M")
                                          if st in ("Resolved", "Closed") else ""})
                    tick += 1

    # ---------------------------------------------------------------- execution plans (one shared across three teams)
    def plan(date, cname, name, tk, start, end, status, phases):
        for ph, chk, steps in phases:
            for no, step, owner_team, mins, cmd, rb in steps:
                plan_teams.setdefault(tk, set()).add(owner_team)
                plans.append({"Date": date, "Customer": cname, "Task_Name": name, "Ticket_ID": tk, "Change_Ref": tk,
                              "Window_Start": start, "Window_End": end, "Approver": f"{cname} CAB", "Task_Status": status,
                              "Phase": ph, "Checkpoint": chk, "Step_No": no, "Step": step,
                              "Owner": lead(owner_team) if no.count(".") == 1 and rng.random() < 0.3 else eng(owner_team),
                              "Planned_Start": None, "Approx_Min": mins, "Actual_Start": None, "Actual_End": None, "Status": "Pending",
                              "Notes": f"{owner_team} step" if owner_team else None, "Command": cmd or None, "Rollback": rb or None, "Expected": None})
    d = lambda n: (TODAY + dt.timedelta(days=n)).isoformat()  # noqa: E731
    plan(d(4), "Harbourline Retail", "Store ERP database upgrade (shared: Unix SA, DBA, Boomi)", "CHG-4101", f"{d(4)} 22:00", f"{d(5)} 03:00", "Scheduled", [
        ("Pre-execution", "Yes", [("1.1", "Stop Boomi listeners for store orders", "Boomi", 10, "systemctl stop atom", "systemctl start atom"),
                                  ("1.2", "Take a full RMAN backup", "DBA", 40, "rman target / cmdfile=full.rman", ""),
                                  ("1.3", "Snapshot the database host", "Unix SA", 15, "lvcreate -s -n snap_u01 -L 50G vg01/u01", "lvconvert --merge vg01/snap_u01")]),
        ("Upgrade", "Yes", [("2.1", "Run the database upgrade", "DBA", 90, "dbupgrade -n 8", "Flashback to restore point PRE_UPG"),
                            ("2.2", "Apply OS kernel parameters", "Unix SA", 15, "sysctl -p /etc/sysctl.d/99-oracle.conf", "")]),
        ("Post execution", "No", [("3.1", "Validate and start Boomi listeners", "Boomi", 15, "systemctl start atom", ""),
                                  ("3.2", "Hand back to the business", "DBA", 10, "", "")])])
    for i, (cname, team, name, tk, days) in enumerate([
            ("Acme Corp", "Unix SA", "Quarterly OS patching, ERP tier", "CHG-4102", 3), ("Globex", "DBA", "SQL Always On failover test", "CHG-4103", 6),
            ("Umbrella Health", "Windows", "ESXi upgrade, cluster HC-01", "CHG-4104", 9), ("Greenfield Foods", "Boomi", "Boomi runtime upgrade", "CHG-4105", 12),
            ("Skyway Air", "NOC", "Firewall firmware upgrade", "CHG-4106", 15), ("Lumen Media", "Windows", "IIS certificate rollover", "CHG-4107", 18),
            ("Cobalt Labs", "DBA", "PostgreSQL minor upgrade", "CHG-4108", 21)]):
        plan(d(days), cname, name, tk, f"{d(days)} 21:00", f"{d(days + 1)} 01:00", "Scheduled" if i < 5 else "Draft", [
            ("Pre-execution", "Yes", [("1.1", "Confirm backups and monitoring", team, 10, "", ""), ("1.2", "Notify the customer", team, 5, "", "")]),
            ("Execution", "No", [("2.1", f"{name}: main change", team, 60, "", "Restore from the pre-change snapshot")]),
            ("Post execution", "No", [("3.1", "Validate services", team, 15, "", ""), ("3.2", "Update the runbook", team, 5, "", "")])])

    # ---------------------------------------------------------------- handovers, shift log, bulletins
    regions = {"Morning": "EMEA Morning", "Afternoon": "US Day", "Night": "APAC Night"}
    checks = {"Unix SA": [("Backup jobs", "All nightly backups completed"), ("Filesystem usage", "No filesystem above 85%")],
              "Windows": [("Patch compliance", "Servers patched to the October baseline"), ("AD replication", "No replication errors")],
              "DBA": [("RMAN backups", "All database backups completed"), ("Data Guard lag", "Standby lag under 5 minutes")],
              "Boomi": [("Integration queues", "No documents stuck in queues"), ("Listener health", "All listeners running")],
              "NOC": [("Monitoring alerts", "No unacknowledged critical alerts"), ("WAN links", "All links up")]}
    for team in TEAMS:
        mine = [i for i in incidents if incident_team.get(i["Ticket ID"]) == team]
        open_p = [i for i in mine if i["Priority"] in ("P1", "P2") and i["Status"] in ("New", "In Progress")]
        for back in range(4, -1, -1):
            day = (TODAY - dt.timedelta(days=back)).isoformat()
            for sh in (("Morning", "Night") if TEAMS[team][1] == "24x7" else ("Morning",)):
                handovers.append({"Team": team, "Date": day, "Shift": sh, "Region": regions[sh], "Outgoing engineer": eng(team), "Incoming engineer": eng(team),
                                  "P1/P2 in flight": ", ".join(f"{i['Ticket ID']} ({i['Customer']})" for i in open_p[:2]) or "None",
                                  "Planned changes": ", ".join(sorted(k for k, v in plan_teams.items() if team in v))[:120],
                                  "High-risk CIs": ", ".join(sorted({i["CI"] for i in open_p}))[:120],
                                  "Notes": f"{team}: {checks[team][0][1].lower()}.", "Acknowledged": "Yes" if back else "No"})
                for name, ok in checks[team]:
                    shiftlog.append({"Team": team, "Date": day, "Shift": sh, "Section": "Shift Handover Overview", "Check": name, "Client": "All",
                                     "Status": "Green", "Details": ok, "Owner": eng(team)})
        for i in mine:
            if i["Status"] in ("New", "In Progress"):
                shiftlog.append({"Team": team, "Date": TODAY.isoformat(), "Shift": "Morning", "Section": "In Work - All Clients", "Ticket": i["Ticket ID"],
                                 "Client": i["Customer"], "Priority": i["Priority"], "Task": i["Summary"][:80], "Owner": i["Engineer"],
                                 "Status": i["Status"], "Next_Step": i["Action"][:80]})
            elif i["Status"] == "On Hold":
                shiftlog.append({"Team": team, "Date": TODAY.isoformat(), "Shift": "Morning", "Section": "On-Hold OR Customer Hold", "Ticket": i["Ticket ID"],
                                 "Client": i["Customer"], "Priority": i["Priority"], "Blocked_On": "Customer confirmation",
                                 "Next_Follow_Up": (TODAY + dt.timedelta(days=2)).isoformat(), "Notes": i["Summary"][:80]})
        for p in sorted({(r["Date"], r["Customer"], r["Task_Name"], r["Change_Ref"], r["Window_Start"], r["Window_End"]) for r in plans}):
            if team in plan_teams.get(p[3], ()):
                shiftlog.append({"Team": team, "Date": TODAY.isoformat(), "Shift": "Morning", "Section": "Scheduled", "Client": p[1], "Event": p[2],
                                 "Assets": "", "Window_Start": p[4], "Window_End": p[5], "Change_Ref": p[3], "Status": "Approved"})
    bulletins = [
        {"Title": "OpenSSH security advisory: patch Linux bastions", "Severity": "Critical", "Customers": "All", "Body": "Patch all Linux bastion hosts this week.",
         "Link": "", "Start": "2026-11-11 08:00", "Expiry": "2026-12-15 23:59", "Pinned": "Yes", "Posted by": "Arun Kumar"},
        {"Title": "Change freeze: Initech month-end", "Severity": "Advisory", "Customers": "Initech", "Body": "No production changes in the last 3 business days of the month.",
         "Link": "", "Start": "2026-11-10 00:00", "Expiry": "2026-12-31 23:59", "Pinned": "Yes", "Posted by": "Priya Sharma"},
        {"Title": "Harbourline DB upgrade: three teams involved", "Severity": "Advisory", "Customers": "Harbourline Retail",
         "Body": "CHG-4101 involves Unix SA, DBA and Boomi. Each team runs its own steps in the shared plan.", "Link": "",
         "Start": "2026-11-12 00:00", "Expiry": "2026-11-19 06:00", "Pinned": "No", "Posted by": "Divya Menon"},
        {"Title": "New escalation number for Globex", "Severity": "Info", "Customers": "Globex", "Body": "The major incident bridge is ext 4123.",
         "Link": "", "Start": "2026-11-05 09:00", "Expiry": "2026-12-31 23:59", "Pinned": "No", "Posted by": "Karthik Reddy"},
        {"Title": "Team rosters for December published", "Severity": "Info", "Customers": "All", "Body": "Each team's December roster is in Shift roster.",
         "Link": "", "Start": "2026-11-12 09:00", "Expiry": "2026-12-31 23:59", "Pinned": "No", "Posted by": "Anjali Nair"},
        {"Title": "Bluepeak partner onboarding", "Severity": "Info", "Customers": "Bluepeak Logistics", "Body": "Three new EDI partners go live on 1 Dec.",
         "Link": "", "Start": "2026-11-12 09:00", "Expiry": "2026-12-05 23:59", "Pinned": "No", "Posted by": "Vikram Rao"},
    ]
    return {"cis": cis, "runbook": runbook, "incidents": incidents, "plans": plans, "handovers": handovers,
            "shiftlog": shiftlog, "bulletins": bulletins}
