import { FormEvent, useState } from "react";
import { api, ApiError } from "../lib/api";
import type { AuthStatus, User } from "../lib/types";

export function Login({ status, onSignedIn }: { status: AuthStatus; onSignedIn: (u: User, csrf: string) => void }) {
  const [username, setU] = useState(""), [password, setP] = useState(""), [busy, setBusy] = useState(false), [err, setErr] = useState(status.note || "");
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setBusy(true); setErr("");
    try { const d = await api<{ user: User; csrf: string }>("POST", "/api/auth/login", { username, password }); onSignedIn(d.user, d.csrf); }
    catch (x) { setErr((x as ApiError).message); }
    finally { setBusy(false); }
  };
  if (status.setup_required) {
    return <div className="login"><form><h1>First-time setup</h1>
      <p className="sub">Create the first administrator in the classic view, then come back here.</p>
      <a className="btn primary" href="/">Open the classic view</a></form></div>;
  }
  return (
    <div className="login">
      <div className="art">
        <div className="brand"><div className="logo" aria-hidden="true">M</div><div><b>MSP Ops Center</b><span>ITIL service operations</span></div></div>
        <div><h2>Plans, runbooks, shifts and terminals for the whole team.</h2>
          <p>Every change and every terminal session is recorded against the person who made it.</p></div>
        <small style={{ color: "#8291AB" }}>New interface preview</small>
      </div>
      <form onSubmit={submit} aria-labelledby="signin">
        <h1 id="signin">Sign in</h1>
        <label className="sr-only" htmlFor="u">User name or email</label>
        <input id="u" className="input" autoComplete="username" placeholder="User name or email" value={username} onChange={e => setU(e.target.value)} required autoFocus />
        <label className="sr-only" htmlFor="p">Password</label>
        <input id="p" className="input" type="password" autoComplete="current-password" placeholder="Password" value={password} onChange={e => setP(e.target.value)} required />
        {err && <div className="err" role="alert">{err}</div>}
        <button className="btn primary" disabled={busy} style={{ justifyContent: "center", height: 42 }}>{busy ? "Signing in…" : "Sign in"}</button>
        {status.methods.entra && <a className="btn" href="/auth/entra/login" style={{ justifyContent: "center" }}>Sign in with Microsoft 365</a>}
        <small style={{ color: "var(--dim)" }}>Forgotten your password? Ask an administrator to reset it.</small>
      </form>
    </div>
  );
}
