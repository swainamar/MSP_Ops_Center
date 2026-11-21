import { useEffect, useState } from "react";
import { api } from "../lib/api";
import { classic } from "../lib/router";
import type { RosterNow } from "../lib/types";

// The shift running now, who is on it and the time left; refreshes every minute.
export function OnShift() {
  const [now, setNow] = useState<RosterNow | null>(null);
  useEffect(() => {
    let alive = true;
    const load = () => api<RosterNow>("GET", "/api/roster/now", undefined, { background: true })
      .then(d => alive && setNow(d)).catch(() => alive && setNow(null));
    load();
    const t = window.setInterval(() => { if (!document.hidden) load(); }, 60_000);
    return () => { alive = false; window.clearInterval(t); };
  }, []);
  const s = now?.now?.shifts?.[0];
  if (!s) return null;
  const h = Math.floor(s.minutes_left / 60), m = s.minutes_left % 60;
  const oncall = Object.entries(now?.now?.oncall || {}).map(([k, v]) => `${k} ${v}`).join(", ");
  return (
    <a className="onshift" href={classic("roster")} style={{ ["--sc" as string]: s.color || undefined }}
       title={`${now?.roster?.name}: ${s.name} ${s.start} to ${s.end}${oncall ? ". On call: " + oncall : ""}`}>
      <span className="pulse" aria-hidden="true" /><b>{s.name}</b>
      <span className="who">{s.people.length ? s.people.join(", ") : "nobody assigned"}</span>
      <span className="left num">{h ? `${h} h ` : ""}{m} m left</span>
    </a>
  );
}
