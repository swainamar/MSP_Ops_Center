const P: Record<string, string> = {
  home: "M3 11 12 4l9 7M5 10v10h14V10", grid: "M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z",
  plans: "M9 6h11M9 12h11M9 18h11M4 6h.01M4 12h.01M4 18h.01", cal: "M4 6h16v14H4zM4 10h16M8 3v4M16 3v4",
  swap: "M4 8h13l-3-3M20 16H7l3 3", roster: "M3 5h18v16H3zM3 10h18M8 2v4M16 2v4M7 14h3M7 18h3M14 14h3M14 18h3",
  term: "M4 5h16v14H4zM8 10l3 2-3 2M13 15h3", book: "M5 4h11a3 3 0 0 1 3 3v13H8a3 3 0 0 1-3-3zM5 17a3 3 0 0 1 3-3h11",
  server: "M4 4h16v6H4zM4 14h16v6H4zM8 7h.01M8 17h.01", alert: "M12 3 2 20h20zM12 10v4M12 17h.01",
  megaphone: "M3 11v2a1 1 0 0 0 1 1h2l5 4V6L6 10H4a1 1 0 0 0-1 1zM15 9a3 3 0 0 1 0 6", pulse: "M3 12h4l3-8 4 16 3-8h4",
  shield: "M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z", sun: "M12 4V2M12 22v-2M4 12H2M22 12h-2M5 5 3.6 3.6M20.4 20.4 19 19M5 19l-1.4 1.4M20.4 3.6 19 5M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10z",
  out: "M15 4h4v16h-4M10 17l5-5-5-5M15 12H3", search: "M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-4.3-4.3",
  close: "M6 6l12 12M18 6 6 18", arrow: "M5 12h14M13 6l6 6-6 6", ext: "M14 4h6v6M20 4 10 14M18 14v6H4V6h6",
  spark: "M12 3l1.8 4.6L18.5 9l-4.7 1.6L12 15l-1.8-4.4L5.5 9l4.7-1.4zM19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9z",
};
export function Icon({ name, title }: { name: keyof typeof P | string; title?: string }) {
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round"
         aria-hidden={title ? undefined : true} role={title ? "img" : undefined}>
      {title && <title>{title}</title>}
      <path d={P[name] || P.grid} />
    </svg>
  );
}
