// Typed client for the MSP Ops Center API. Same origin, session cookie, CSRF header on writes.
export class ApiError extends Error {
  constructor(public status: number, message: string) { super(message); }
}

let csrf = "";
export const setCsrf = (token: string) => { csrf = token || ""; };

type Opts = { background?: boolean; signal?: AbortSignal };

export async function api<T = any>(method: string, path: string, body?: unknown, opts: Opts = {}): Promise<T> {
  const headers: Record<string, string> = { Accept: "application/json" };
  if (body !== undefined) headers["Content-Type"] = "application/json";
  if (method !== "GET") headers["X-CSRF-Token"] = csrf;
  if (opts.background) headers["X-Background"] = "1";
  let res: Response;
  try {
    res = await fetch(path, { method, headers, credentials: "same-origin", signal: opts.signal,
                              body: body === undefined ? undefined : JSON.stringify(body) });
  } catch (e) {
    if ((e as Error).name === "AbortError") throw e;
    throw new ApiError(0, "The MSP Ops Center server cannot be reached. Check that it is running.");
  }
  let data: any = {};
  try { data = await res.json(); } catch { /* empty or non-JSON reply */ }
  if (!res.ok || data.ok === false) {
    if (res.status === 401) window.dispatchEvent(new Event("msp:signed-out"));
    throw new ApiError(res.status, data.error || `Request failed (${res.status}).`);
  }
  return data as T;
}
