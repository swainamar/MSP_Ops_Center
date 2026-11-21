import { useEffect, useState } from "react";

// Hash routes inside /app/: #/mywork, #/cmdb, #/cmdb/<ci name>. Screens not yet moved link to the classic view.
export type Route = { page: string; param: string };

function parse(): Route {
  const h = decodeURIComponent(window.location.hash.replace(/^#\/?/, ""));
  const [page, ...rest] = h.split("/");
  return { page: page || "mywork", param: rest.join("/") };
}

export function useRoute(): Route {
  const [route, setRoute] = useState(parse);
  useEffect(() => {
    const on = () => { setRoute(parse()); window.scrollTo(0, 0); };
    window.addEventListener("hashchange", on);
    return () => window.removeEventListener("hashchange", on);
  }, []);
  return route;
}

export const go = (page: string, param = "") => { window.location.hash = "/" + page + (param ? "/" + encodeURIComponent(param) : ""); };

// The classic page uses #<view> or #<view>=<param>.
export const classic = (view: string) => `/#${view}`;
