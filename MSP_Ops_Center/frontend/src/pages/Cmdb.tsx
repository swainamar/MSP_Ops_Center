import { useEffect, useMemo, useState } from "react";
import { go, classic } from "../lib/router";
import { isOpen, nk, truthy } from "../lib/util";
import type { Row, Snapshot } from "../lib/types";
import { Icon } from "../components/Icon";

const GROUPS: [string, string[]][] = [
  ["Identity", ["Customer", "CI_Type", "Status", "Environment", "Criticality", "Hostname_FQDN", "IP_Address", "Serial_No"]],
  ["Platform", ["Platform_Type", "Virtualization_Cloud", "OS", "OS_Version", "Application", "Role", "CPU", "Memory", "Root_Disk", "Data_Disk", "Data_Centre_Rack"]],
  ["Operations", ["Patch_Group", "Maintenance_Window", "Backup_Policy", "Monitoring", "Upstream_Dependencies", "Downstream_Dependencies"]],
  ["Support", ["Owner", "Contacts", "Warranty_End", "OS_EOL"]],
];
const STATUS_BADGE: Record<string, string> = { Active: "b-ok", Maintenance: "b-P2", "In build": "b-Info", Planned: "b-muted", Decommissioned: "b-muted" };

export function Cmdb({ snap, open }: { snap: Snapshot; open: string }) {
  const t = snap.tables.Asset_Master, sc = snap.schema.Asset_Master;
  const label = useMemo(() => Object.fromEntries(t.columns.map(c => [c.key, c.label])), [t]);
  const [q, setQ] = useState(""), [cust, setCust] = useState(""), [typ, setTyp] = useState(""), [st, setSt] = useState("");
  const [sort, setSort] = useState<[string, 1 | -1]>(["Device", 1]);
  const shown = ["Device", "CI_Type", "Status", "Customer", "Environment", "IP_Address", "OS", "Application", "Criticality", "Owner"];
  const rows = useMemo(() => {
    const needle = q.trim().toLowerCase();
    return t.rows.filter(r => (!cust || r.Customer === cust) && (!typ || r.CI_Type === typ) && (!st || r.Status === st)
        && (!needle || Object.values(r).some(v => String(v).toLowerCase().includes(needle))))
      .sort((a, b) => sort[1] * String(a[sort[0]] || "").localeCompare(String(b[sort[0]] || ""), undefined, { numeric: true }));
  }, [t, q, cust, typ, st, sort]);
  const tally = (key: string) => t.rows.reduce<Record<string, number>>((m, r) => { if (r[key]) m[r[key]] = (m[r[key]] || 0) + 1; return m; }, {});
  const types = tally("CI_Type"), statuses = tally("Status");
  const customers = [...new Set(t.rows.map(r => r.Customer).filter(Boolean))].sort();
  const sel = open ? t.rows.find(r => nk(r.Device) === nk(open)) : undefined;
  const chip = (cur: string, set: (v: string) => void, v: string, n: number) =>
    <button key={v} className="chip" aria-pressed={cur === v} onClick={() => set(cur === v ? "" : v)}>{v}<span className="n">{n}</span></button>;
  return (
    <>
      <h1>{sc.title}</h1>
      <p className="sub">Every configuration item (CI) with its type, status, platform, dependencies and special-care warnings. Select a CI to see its details, runbook and incidents.</p>
      <div className="toolbar">
        <select className="select" value={cust} onChange={e => setCust(e.target.value)} aria-label="Customer">
          <option value="">All customers</option>{customers.map(c => <option key={c}>{c}</option>)}</select>
        <input className="input grow" type="search" placeholder="Search configuration items" value={q} onChange={e => setQ(e.target.value)} aria-label="Search configuration items" />
        <a className="btn" href={classic("Asset_Master")}>Add or edit in classic view</a>
      </div>
      {Object.keys(types).length > 0 && <div className="chips" role="group" aria-label="CI type">{Object.entries(types).sort().map(([v, n]) => chip(typ, setTyp, v, n))}</div>}
      <div className="chips" role="group" aria-label="CI status">{Object.entries(statuses).sort().map(([v, n]) => chip(st, setSt, v, n))}</div>
      <div className="tablewrap">
        <table>
          <thead><tr>{shown.map(k => <th key={k} aria-sort={sort[0] === k ? (sort[1] > 0 ? "ascending" : "descending") : "none"}>
            <button onClick={() => setSort([k, sort[0] === k ? (-sort[1] as 1 | -1) : 1])}>{label[k] || k}{sort[0] === k ? (sort[1] > 0 ? " ↑" : " ↓") : ""}</button></th>)}</tr></thead>
          <tbody>{rows.map(r => (
            <tr key={r.Record_ID} tabIndex={0} onClick={() => go("cmdb", r.Device)} onKeyDown={e => e.key === "Enter" && go("cmdb", r.Device)}>
              {shown.map(k => k === "Device" ? <td key={k} className="name">{r.Device}</td>
                : k === "Status" ? <td key={k}>{r.Status && <span className={"badge " + (STATUS_BADGE[r.Status] || "b-muted")}>{r.Status}</span>}</td>
                : <td key={k}>{r[k]}</td>)}</tr>))}
            {!rows.length && <tr><td colSpan={shown.length}><p className="empty">No configuration items match.</p></td></tr>}</tbody>
        </table>
      </div>
      <p className="foot num">{rows.length} of {t.rows.length} configuration items</p>
      {sel && <CiDrawer ci={sel} snap={snap} label={label} />}
    </>
  );
}

function CiDrawer({ ci, snap, label }: { ci: Row; snap: Snapshot; label: Record<string, string> }) {
  const close = () => go("cmdb");
  useEffect(() => { const k = (e: KeyboardEvent) => e.key === "Escape" && close(); window.addEventListener("keydown", k); return () => window.removeEventListener("keydown", k); }, []);
  const rb = (snap.tables.Runbook?.rows || []).filter(r => nk(r.Asset) === nk(ci.Device)).sort((a, b) => String(b.Date).localeCompare(String(a.Date)));
  const incs = (snap.tables.Incidents?.rows || []).filter(r => nk(r.Asset) === nk(ci.Device));
  return (
    <>
      <div className="scrim" onClick={close} />
      <aside className="drawer" role="dialog" aria-modal="true" aria-labelledby="ci-name">
        <header><div style={{ flex: 1 }}><small style={{ color: "var(--dim)", textTransform: "uppercase", letterSpacing: ".06em", fontSize: 11 }}>{ci.CI_Type || "Configuration item"}</small>
            <h2 id="ci-name">{ci.Device}</h2><small style={{ color: "var(--muted)" }}>{ci.Customer} · {ci.Environment}</small></div>
          {ci.Status && <span className={"badge " + (STATUS_BADGE[ci.Status] || "b-muted")}>{ci.Status}</span>}
          <button className="iconbtn" onClick={close} aria-label="Close"><Icon name="close" /></button></header>
        <div className="body">
          {(ci as Record<string, any>)._view_only && <div className="notice" role="note">Shared with your team by <b>{String((ci as Record<string, any>)._owner || "another team")}</b>: view only. Only the fields they share are shown; the owning team makes changes.</div>}
          {ci.Special_Care && <div className="care"><b>Special care:</b> {ci.Special_Care}</div>}
          {(() => { // ownership and sharing model (custom columns, found by label)
            const keyOf = (l: string) => Object.entries(label).find(([, v]) => v === l)?.[0];
            const share = ["Owner team", "Sharing", "Shared with teams", "Shared fields"].map(l => [l, keyOf(l)] as const).filter(([, k]) => k && String(ci[k!] || "").trim());
            const teamOnly = Object.entries(label).filter(([k, l]) => l.endsWith("(team only)") && String(ci[k] || "").trim());
            return <>
              {share.length > 0 && <section className="panel"><h2>Ownership and sharing</h2><dl className="facts">{share.map(([l, k]) =>
                <div key={l}><dt>{l}</dt><dd>{l === "Sharing" ? <span className={"badge " + (ci[k!] === "Team only" ? "b-P2" : ci[k!] === "Fully shared" ? "b-ok" : "b-Info")}>{ci[k!]}</span> : ci[k!]}</dd></div>)}</dl></section>}
              {teamOnly.length > 0 && <section className="panel"><h2>Team-only details</h2><dl className="facts">{teamOnly.map(([k, l]) =>
                <div key={k} style={{ gridColumn: "1 / -1" }}><dt>{l.replace(" (team only)", "")}</dt><dd>{ci[k]}</dd></div>)}</dl>
                <p className="foot">Visible only to that team once isolation is switched on.</p></section>}
            </>;
          })()}
          {GROUPS.map(([g, keys]) => { const f = keys.filter(k => String(ci[k] || "").trim()); return f.length ? (
            <section key={g} className="panel"><h2>{g}</h2><dl className="facts">{f.map(k => <div key={k}><dt>{label[k] || k}</dt><dd>{ci[k]}</dd></div>)}</dl></section>) : null; })}
          <section className="panel"><h2>Runbook<span className="count">{rb.length || ""}</span></h2>
            {rb.length ? <ul className="list">{rb.map(r => <li key={r.Record_ID}>{truthy(r.Known_Issue) && <span className="badge b-P2">Known issue</span>}
              <div className="t"><div>{r.Title}</div><small>{r.Date} · {r.Category} · {r.Engineer}</small>{r.Solution && <small style={{ color: "var(--muted)" }}>Solution: {r.Solution}</small>}</div></li>)}</ul>
              : <p className="empty">No runbook entries for this CI yet.</p>}</section>
          <section className="panel"><h2>Incidents<span className={"count" + (incs.some(r => isOpen(r.Status)) ? " hot" : "")}>{incs.length || ""}</span></h2>
            {incs.length ? <ul className="list">{incs.map(r => <li key={r.Record_ID}><span className={"badge b-" + r.Priority}>{r.Priority}</span>
              <div className="t"><div>{r.Ticket_ID} {r.Summary}</div><small>{r.Status} · {r.Engineer}</small></div></li>)}</ul>
              : <p className="empty">No incidents recorded for this CI.</p>}</section>
          <a className="btn" href={`/#runbook`}>Open the full runbook in the classic view</a>
        </div>
      </aside>
    </>
  );
}
