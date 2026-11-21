import { useEffect, useMemo, useState } from "react";
import { api } from "../lib/api";
import { classic } from "../lib/router";
import { addDays, fmtTime, isOpen, nk, nowLocal, today, truthy } from "../lib/util";
import type { MyShift, Snapshot, Step, Task, Phase, Row, User } from "../lib/types";

const SEV: Record<string, number> = { Critical: 3, Advisory: 2, Info: 1 };

// The same rules as the classic My work page, so both views list the same items.
export function myWork(snap: Snapshot, user: User) {
  const mine = (v: unknown) => { const n = nk(v); return !!n && (n === nk(user.name) || n === nk(user.username)); };
  const td = today(), from = addDays(td, -1), to = addDays(td, 7);
  const steps: { t: Task; ph: Phase; st: Step }[] = [];
  snap.plans.tasks.forEach(t => {
    if (t.date < from || t.date > to || ["Completed", "Cancelled"].includes(t.status)) return;
    t.phases.forEach(ph => ph.steps.forEach(st => { if (mine(st.owner) && !["Done", "Skipped"].includes(st.status || "")) steps.push({ t, ph, st }); }));
  });
  steps.sort((a, b) => String(a.st.planned).localeCompare(String(b.st.planned)));
  const gates = user.rank >= 3 ? snap.plans.tasks.filter(t => t.date >= from && t.date <= to && ["Scheduled", "In Progress"].includes(t.status))
    .flatMap(t => t.phases.filter(p => p.checkpoint && !p.signed_off_by).map(p => ({ t, p }))) : [];
  const incs = (snap.tables.Incidents?.rows || []).filter(r => isOpen(r.Status) && mine(r.Engineer));
  const now = nowLocal(), reads = new Set(snap.my_reads);
  const unread = (snap.tables.Bulletins?.rows || []).filter(b => (!b.Start || b.Start <= now) && (!b.Expiry || b.Expiry >= now) && !reads.has(b.Record_ID))
    .sort((a, b) => (+truthy(b.Pinned) - +truthy(a.Pinned)) || ((SEV[b.Severity] || 0) - (SEV[a.Severity] || 0)));
  return { steps, gates, incs, unread };
}

function Panel({ title, count, hot, children }: { title: string; count?: number; hot?: boolean; children: React.ReactNode }) {
  return <section className="panel"><h2>{title}{count ? <span className={"count" + (hot ? " hot" : "")}>{count}</span> : null}</h2>{children}</section>;
}

export function MyWork({ snap, user }: { snap: Snapshot; user: User }) {
  const w = useMemo(() => myWork(snap, user), [snap, user]);
  const [shifts, setShifts] = useState<MyShift[] | null>(null);
  useEffect(() => { api<{ shifts: MyShift[] }>("GET", "/api/roster/mine?days=7", undefined, { background: true })
    .then(d => setShifts(d.shifts.filter(s => s.code !== "OFF"))).catch(() => setShifts([])); }, [snap.revision]);
  const recent = snap.audit.filter(a => nk(a.User) === nk(user.username)).slice(0, 7);
  const hour = new Date().getHours();
  const hello = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";
  return (
    <>
      <h1>{hello}, {(user.name || user.username).split(" ")[0]}</h1>
      <p className="sub">{new Date().toLocaleDateString([], { weekday: "long", day: "numeric", month: "long" })}. You are signed in as {user.role}.</p>
      <div className="grid3">
        <Panel title="Your steps in the next 7 days" count={w.steps.length}>
          {w.steps.length ? <ul className="list">{w.steps.slice(0, 8).map(({ t, ph, st }) =>
            <li key={st.id}><div className="t"><div>{st.no ? st.no + " " : ""}{st.title}</div><small>{t.name} · {ph.name} · {t.date}</small></div>
              <a className="btn sm" href={classic(`plans`)}>Open</a></li>)}</ul>
            : <p className="empty">No steps are assigned to you. Steps count as yours when the Owner matches your name.</p>}
        </Panel>
        <Panel title="Waiting for your go/no-go" count={w.gates.length} hot={w.gates.length > 0}>
          {w.gates.length ? <ul className="list">{w.gates.map(({ t, p }) =>
            <li key={p.id}><div className="t"><div>{t.name}: {p.name}</div><small>{t.date}, {t.customer}</small></div></li>)}</ul>
            : <p className="empty">{user.rank >= 3 ? "No checkpoints are waiting." : "Checkpoints are signed off by Leads."}</p>}
        </Panel>
        <Panel title="Your shifts, next 7 days" count={shifts?.length}>
          {shifts === null ? <p className="empty">Loading…</p> : shifts.length ? <ul className="list">{shifts.map(s =>
            <li key={s.date + s.roster}><span className="badge b-muted">{s.code}</span><div className="t"><div>{s.weekday} {s.date}: {s.shift} {s.start && `${s.start} to ${s.end}`}</div>
              <small>{s.roster}{s.oncall.length ? `, on call (${s.oncall.join(", ")})` : ""}{s.holiday ? `, ${s.holiday}` : ""}</small></div></li>)}</ul>
            : <p className="empty">No roster shifts in the next 7 days. A Lead adds you to a roster with your user name.</p>}
        </Panel>
        <Panel title="Your open incidents" count={w.incs.length} hot={w.incs.some(r => ["P1", "P2"].includes(r.Priority))}>
          {w.incs.length ? <ul className="list">{w.incs.map((r: Row) =>
            <li key={r.Record_ID}><span className={"badge b-" + r.Priority}>{r.Priority}</span><div className="t"><div>{r.Ticket_ID} {r.Summary}</div>
              <small>{r.Customer} / {r.Asset} / {r.Status}</small></div></li>)}</ul>
            : <p className="empty">Nothing assigned to you.</p>}
        </Panel>
        <Panel title="Unread bulletins" count={w.unread.length}>
          {w.unread.length ? <ul className="list">{w.unread.map(b =>
            <li key={b.Record_ID}><span className={"badge b-" + b.Severity}>{b.Severity}</span><div className="t"><div>{b.Title}</div><small>{b.Customers}</small></div></li>)}</ul>
            : <p className="empty">You are up to date.</p>}
        </Panel>
        <Panel title="Your recent changes">
          {recent.length ? <ul className="list">{recent.map((a, i) =>
            <li key={i}><div className="t"><div>{a.Action}{a.Target ? `: ${a.Target}` : ""}</div><small>{fmtTime(a.Timestamp)}</small></div></li>)}</ul>
            : <p className="empty">No changes yet.</p>}
        </Panel>
      </div>
    </>
  );
}
