import { ReactNode } from "react";
import { Icon } from "./Icon";
import { OnShift } from "./OnShift";
import { classic, Route } from "../lib/router";
import { initials } from "../lib/util";
import type { User } from "../lib/types";

// Screens marked `moved` render in this front end; the rest open in the classic view until they are moved.
export const NAV: [string, [string, string, string, boolean][]][] = [
  ["Home", [["mywork", "My work", "home", true], ["overview", "Overview", "grid", false]]],
  ["Operations", [["plans", "Execution plans", "plans", false], ["shiftlog", "Shift log", "cal", false],
                  ["Shift_Transitions", "Shift handovers", "swap", false], ["roster", "Shift roster", "roster", false],
                  ["terminals", "Terminals", "term", false]]],
  ["Knowledge", [["runbook", "Runbook", "book", false], ["cmdb", "CMDB", "server", true], ["Incidents", "Incidents", "alert", false]]],
  ["Communication", [["Bulletins", "Bulletins", "megaphone", false]]],
  ["Administration", [["activity", "Activity", "pulse", false], ["admin", "Admin", "shield", false]]],
];
const TITLES: Record<string, [string, string]> = { mywork: ["Home", "My work"], cmdb: ["Knowledge", "CMDB"] };

type Props = { user: User; route: Route; counts: Record<string, [number, boolean]>; onAsk: () => void; onTheme: () => void;
               onSignOut: () => void; children: ReactNode };

export function Shell({ user, route, counts, onAsk, onTheme, onSignOut, children }: Props) {
  const [grp, title] = TITLES[route.page] || ["", ""];
  return (
    <div className="app">
      <aside className="side" aria-label="Main navigation">
        <div className="brand"><div className="logo" aria-hidden="true">M</div>
          <div><b>MSP Ops Center</b><span>ITIL service operations</span></div></div>
        <nav className="nav">
          {NAV.map(([g, items]) => (
            <div key={g}><div className="grp">{g}</div>
              {items.filter(([id]) => id !== "admin" || user.rank >= 4).map(([id, label, icon, moved]) => {
                const c = counts[id];
                return (
                  <a key={id} className={"blade" + (route.page === id ? " on" : "")} href={moved ? `#/${id}` : classic(id)}
                     aria-current={route.page === id ? "page" : undefined} title={moved ? undefined : "Opens in the classic view"}>
                    <span className="ic"><Icon name={icon} /></span>{label}
                    {c && c[0] > 0 ? <span className={"cnt" + (c[1] ? " hot" : "")}>{c[0]}</span> : !moved ? <span className="ext">classic</span> : null}
                  </a>
                );
              })}
            </div>
          ))}
        </nav>
        <div className="side-foot">
          <div className="avatar" aria-hidden="true">{initials(user.name || user.username)}</div>
          <div className="who"><b>{user.name || user.username}</b><span className="rolepill">{user.role}</span></div>
          <button className="iconbtn" onClick={onTheme} aria-label="Switch light or dark theme"><Icon name="sun" /></button>
          <button className="iconbtn" onClick={onSignOut} aria-label="Sign out"><Icon name="out" /></button>
        </div>
      </aside>
      <div className="main">
        <header className="topbar">
          <div className="crumbs">{grp} / <b>{title}</b></div>
          <OnShift />
          <span className="sp" />
          <a className="btn sm" href="/" title="The full classic interface, with every screen">Classic view</a>
          <button className="askbtn" onClick={onAsk} aria-label="Ask MSP Ops Center (Ctrl J)"><Icon name="spark" />Ask<kbd>Ctrl J</kbd></button>
        </header>
        <main className="view" id="main">{children}</main>
      </div>
    </div>
  );
}
