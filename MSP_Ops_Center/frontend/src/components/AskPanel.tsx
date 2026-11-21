import { FormEvent, Fragment, KeyboardEvent, ReactNode, useEffect, useRef, useState } from "react";
import { api, ApiError } from "../lib/api";
import { Icon } from "./Icon";
import type { AskAnswer } from "../lib/types";

const SUGG = ["Who is on shift now?", "Who is on call this week?", "Open P1 incidents", "What changes are planned this week?", "Known issues"];
type Turn = { q: string; a?: AskAnswer; err?: string };

// Turn an answer link into a URL: CMDB items open here, everything else in the classic view.
function hrefFor(link: string) {
  const i = link.indexOf("="), view = i < 0 ? link : link.slice(0, i), p = i < 0 ? "" : link.slice(i + 1);
  if (view === "Asset_Master") return `#/cmdb/${encodeURIComponent(p)}`;
  return `/#${link}`;
}

function cites(text: string): ReactNode[] {
  return text.split(/(\[\d{1,2}\])/g).map((part, i) => {
    const m = part.match(/^\[(\d{1,2})\]$/);
    return m ? <span key={i} className="cite" aria-label={`source ${m[1]}`}>{m[1]}</span> : <Fragment key={i}>{part}</Fragment>;
  });
}

export function AskPanel({ onClose }: { onClose: () => void }) {
  const [log, setLog] = useState<Turn[]>([]), [q, setQ] = useState(""), [busy, setBusy] = useState(false);
  const box = useRef<HTMLTextAreaElement>(null), end = useRef<HTMLDivElement>(null);
  useEffect(() => { box.current?.focus(); }, []);
  useEffect(() => { end.current?.scrollIntoView({ block: "end" }); }, [log]);
  const send = async (text: string) => {
    text = text.trim(); if (!text || busy) return;
    setQ(""); setBusy(true); setLog(l => [...l, { q: text }]);
    try { const d = await api<{ answer: AskAnswer }>("POST", "/api/ai/ask", { question: text });
          setLog(l => l.map((t, i) => i === l.length - 1 ? { ...t, a: d.answer } : t)); }
    catch (e) { setLog(l => l.map((t, i) => i === l.length - 1 ? { ...t, err: (e as ApiError).message } : t)); }
    finally { setBusy(false); }
  };
  const onKey = (e: KeyboardEvent) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(q); } };
  return (
    <>
      <div className="ask-scrim" onClick={onClose} />
      <aside className="ask" role="dialog" aria-modal="true" aria-labelledby="ask-t">
        <header><div className="logo" aria-hidden="true"><Icon name="spark" /></div>
          <div><h2 id="ask-t">Ask MSP Ops Center</h2><small>Answers from this tool's records, on this server only</small></div>
          <span className="sp" /><button className="iconbtn" onClick={onClose} aria-label="Close"><Icon name="close" /></button></header>
        <div className="log" aria-live="polite">
          {!log.length && <div><h3 style={{ font: "600 22px var(--display)", margin: "4px 0 6px" }}>What do you need to know?</h3>
            <p className="sub">Ask in plain words: who is on shift, how something was fixed before, what is planned, or anything about a CI.</p>
            <div className="sugg">{SUGG.map(s => <button key={s} className="chip" onClick={() => send(s)}>{s}</button>)}</div></div>}
          {log.map((t, i) => (
            <Fragment key={i}>
              <div className="q">{t.q}</div>
              {t.err ? <div className="a"><p>{t.err}</p></div> : !t.a ? <div className="a"><p>Searching the records…</p></div> : (
                <div className="a">
                  {t.a.blocks.map((b, j) => b.type === "list" ? <ul key={j}>{b.items!.map((it, k) => <li key={k}>{cites(it)}</li>)}</ul>
                    : b.type === "code" ? <pre key={j}>{b.text}</pre> : <p key={j}>{cites(b.text || "")}</p>)}
                  {t.a.sources.length > 0 && <div className="srcs">{t.a.sources.map(s => (
                    <a key={s.n} className="src" href={hrefFor(s.link)} onClick={onClose}><span className="n">{s.n}</span>
                      <span><span className="k">{s.label}</span>{s.title}{s.where && <span className="w">{s.where}</span>}</span></a>))}</div>}
                  {t.a.notes.map((n, k) => <div key={k} className="meta" style={{ color: "var(--care)" }}>{n}</div>)}
                  <div className="meta">{t.a.mode === "model" ? "Written by the local model" : `Searched ${t.a.records} records`} · {(t.a.took_ms / 1000).toFixed(2)} s</div>
                </div>)}
            </Fragment>))}
          <div ref={end} />
        </div>
        <form onSubmit={(e: FormEvent) => { e.preventDefault(); send(q); }}>
          <div className="box"><textarea ref={box} rows={1} maxLength={500} value={q} onChange={e => setQ(e.target.value)} onKeyDown={onKey}
                                         placeholder="Ask about CIs, fixes, shifts, changes, incidents…" aria-label="Your question" />
            <button className="btn primary" disabled={busy || !q.trim()} aria-label="Ask"><Icon name="arrow" /></button></div>
          <p className="meta">Answers only use records you can see, and show their sources. Check anything important before acting on it.</p>
        </form>
      </aside>
    </>
  );
}
