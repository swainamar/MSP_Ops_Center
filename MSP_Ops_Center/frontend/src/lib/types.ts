export type Role = "Viewer" | "Engineer" | "Lead" | "Admin";
export interface User { username: string; name: string; role: Role; rank: number; must_change?: boolean }
export interface AuthStatus { authenticated: boolean; setup_required: boolean; user?: User; csrf?: string; note?: string;
                              methods: { local: boolean; entra: boolean; proxy: boolean } }
export interface Column { key: string; label: string; core: boolean; system: boolean; hidden: boolean }
export type Row = Record<string, string> & { Record_ID: string };
export interface Table { columns: Column[]; rows: Row[] }
export interface Step { id: string; no?: string; title: string; owner?: string; status?: string; planned?: string; command?: string; rollback?: string }
export interface Phase { id: string; name: string; checkpoint?: boolean; signed_off_by?: string; steps: Step[] }
export interface Task { id: string; date: string; customer: string; name: string; ticket_id: string; change_ref?: string;
                        window_start?: string; window_end?: string; status: string; phases: Phase[];
                        computed?: { steps: number; done: number; progress: number } }
export interface Schema { title: string; noun: string; choices: Record<string, string[]>; dates: string[]; label_key: string }
export interface Snapshot { tables: Record<string, Table>; plans: { tasks: Task[] }; revision: number; my_reads: string[];
                            schema: Record<string, Schema>; audit: AuditEntry[] }
export interface AuditEntry { Timestamp: string; User: string; Area: string; Action: string; Target: string; Details: string }
export interface RosterShiftNow { name: string; start: string; end: string; color: string; minutes_left: number; people: string[] }
export interface RosterNow { roster: { id: string; name: string } | null; now?: { shifts: RosterShiftNow[]; oncall: Record<string, string> };
                             today?: Record<string, string[]> }
export interface MyShift { date: string; weekday: string; code: string; shift: string; start: string; end: string;
                           roster: string; oncall: string[]; holiday: string; shift_log: string }
export interface AskSource { n: number; kind: string; label: string; title: string; where: string; link: string }
export interface AskBlock { type: "p" | "list" | "code"; text?: string; items?: string[] }
export interface AskAnswer { blocks: AskBlock[]; sources: AskSource[]; mode: string; found: boolean; notes: string[];
                             took_ms: number; records: number; question: string }
