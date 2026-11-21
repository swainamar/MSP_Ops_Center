import { useCallback, useEffect, useState } from "react";
import { api, ApiError, setCsrf } from "./lib/api";
import { useRoute, classic } from "./lib/router";
import { isOpen } from "./lib/util";
import type { AuthStatus, Snapshot, User } from "./lib/types";
import { Shell } from "./components/Shell";
import { Login } from "./components/Login";
import { AskPanel } from "./components/AskPanel";
import { MyWork, myWork } from "./pages/MyWork";
import { Cmdb } from "./pages/Cmdb";

const THEME_KEY = "msp.theme";

export default function App() {
  const [status, setStatus] = useState<AuthStatus | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [snap, setSnap] = useState<Snapshot | null>(null);
  const [error, setError] = useState("");
  const [ask, setAsk] = useState(false);
  const route = useRoute();

  useEffect(() => { document.documentElement.dataset.theme = localStorage.getItem(THEME_KEY) === "light" ? "light" : "dark"; }, []);
  const toggleTheme = () => { const t = document.documentElement.dataset.theme === "light" ? "dark" : "light";
    document.documentElement.dataset.theme = t; localStorage.setItem(THEME_KEY, t); };

  const checkAuth = useCallback(() => api<AuthStatus>("GET", "/api/auth/status").then(s => {
    setStatus(s);
    if (s.authenticated && s.user) { setCsrf(s.csrf || ""); setUser(s.user); } else { setUser(null); setSnap(null); }
  }).catch((e: ApiError) => setError(e.message)), []);
  useEffect(() => { checkAuth(); const out = () => checkAuth(); window.addEventListener("msp:signed-out", out);
    return () => window.removeEventListener("msp:signed-out", out); }, [checkAuth]);

  // Data: load once signed in, then refresh every 30 seconds while the tab is visible.
  useEffect(() => {
    if (!user) return;
    let alive = true;
    const load = (bg: boolean) => api<Snapshot>("GET", "/api/state", undefined, { background: bg })
      .then(d => { if (alive) { setSnap(d); setError(""); } }).catch((e: ApiError) => alive && e.status !== 401 && setError(e.message));
    load(false);
    const t = window.setInterval(() => { if (!document.hidden) load(true); }, 30_000);
    return () => { alive = false; window.clearInterval(t); };
  }, [user]);

  useEffect(() => {
    const k = (e: KeyboardEvent) => { if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === "j" && user) { e.preventDefault(); setAsk(a => !a); }
                                      else if (e.key === "Escape") setAsk(false); };
    window.addEventListener("keydown", k); return () => window.removeEventListener("keydown", k);
  }, [user]);

  const signOut = async () => { try { await api("POST", "/api/auth/logout", {}); } finally { checkAuth(); } };

  if (!status) return <div className="view">{error ? <p className="err" role="alert">{error}</p> : <p className="empty">Loading…</p>}</div>;
  if (!user) return <Login status={status} onSignedIn={(u, c) => { setCsrf(c); setUser(u); }} />;
  if (user.must_change) return <div className="view"><h1>Change your password</h1>
    <p className="sub">Your administrator asked you to choose a new password. Do it in the classic view, then come back.</p>
    <a className="btn primary" href="/">Open the classic view</a></div>;

  const counts: Record<string, [number, boolean]> = {};
  if (snap) {
    const w = myWork(snap, user);
    counts.mywork = [w.steps.length + w.gates.length + w.incs.length, false];
    counts.cmdb = [snap.tables.Asset_Master?.rows.length || 0, false];
    const open = (snap.tables.Incidents?.rows || []).filter(r => isOpen(r.Status));
    counts.Incidents = [open.length, open.some(r => ["P1", "P2"].includes(r.Priority))];
  }
  const page = route.page === "cmdb" || route.page === "mywork" ? route.page : "mywork";
  return (
    <Shell user={user} route={{ ...route, page }} counts={counts} onAsk={() => setAsk(true)} onTheme={toggleTheme} onSignOut={signOut}>
      {error && <div className="notice" role="alert">{error}</div>}
      {!snap ? <p className="empty">Loading…</p>
        : page === "cmdb" ? <Cmdb snap={snap} open={route.param} />
        : <MyWork snap={snap} user={user} />}
      {route.page !== page && <div className="notice">That screen has not moved to the new interface yet. <a href={classic(route.page)}>Open it in the classic view</a>.</div>}
      {ask && <AskPanel onClose={() => setAsk(false)} />}
    </Shell>
  );
}
