// Calendar: calendars, events (via the week grid and event "peek" popovers), create / edit / delete.
import type { AxNode } from "../ax";
import { KEY } from "../ax";
import { ax, bodyEditor, need, OutlookError, secondaryWindows, sleep, switchModule, until, withFocus } from "./session";

// ---------------------------------------------------------------------------------------------
// Calendars

export interface CalendarInfo {
  path: string; // "account/group/name", e.g. "you@example.com/My Calendars/Calendar"
  name: string;
  account: string;
  shown: boolean;
}

interface CalRow extends CalendarInfo {
  checkbox: string; // handle
  isGroup: boolean;
  level: number;
}

async function navPane(): Promise<AxNode> {
  await switchModule("Calendar");
  return need("main", { role: "AXOutline", desc: "Navigation pane" }, "the calendar list", { depth: 6 });
}

async function readCalRows(): Promise<CalRow[]> {
  const nav = await navPane();
  const rows = await ax.find(nav.h, { role: "AXRow" }, { depth: 1, limit: 500, subtree: 3, extra: ["AXDisclosureLevel"] });
  const stack: string[] = [];
  const out: CalRow[] = [];
  for (const r of rows) {
    const level = Number(r.x?.AXDisclosureLevel ?? 0);
    const flat: AxNode[] = [];
    const walk = (n: AxNode) => {
      flat.push(n);
      n.children?.forEach(walk);
    };
    walk(r);
    const cb = flat.find((n) => n.role === "AXCheckBox");
    if (!cb?.desc) continue;
    const isGroup = flat.some((n) => n.role === "AXDisclosureTriangle");
    const [rawName, state] = cb.desc.split(/;\s*/);
    const name = rawName!.trim();
    stack.length = level;
    stack[level] = name;
    out.push({
      path: stack.slice(0, level + 1).join("/"),
      name,
      account: stack[0] ?? name,
      shown: state === "Shown" || state === "Partially Shown",
      checkbox: cb.h,
      isGroup,
      level,
    });
  }
  return out;
}

export async function listCalendars(): Promise<CalendarInfo[]> {
  return (await readCalRows()).filter((r) => !r.isGroup).map(({ checkbox, isGroup, level, ...c }) => c);
}

function norm(s: string) {
  return s.trim().replace(/^\/+|\/+$/g, "").toLowerCase();
}

/** Accepts a full path, "account/name", or a unique calendar name. */
async function findCalendar(spec: string): Promise<CalRow> {
  const rows = (await readCalRows()).filter((r) => !r.isGroup);
  const want = norm(spec);
  const hit =
    rows.find((r) => norm(r.path) === want) ??
    rows.find((r) => norm(`${r.account}/${r.name}`) === want);
  if (hit) return hit;
  const byName = rows.filter((r) => r.name.toLowerCase() === want);
  if (byName.length === 1) return byName[0]!;
  if (byName.length > 1) throw new OutlookError(`Calendar "${spec}" is ambiguous: ${byName.map((r) => `"${r.path}"`).join(", ")}`);
  throw new OutlookError(`Calendar "${spec}" not found. Use outlook_list_calendars.`);
}

/** Show exactly the given calendars (leaf rows); returns a restore function. */
async function isolateCalendars(paths: Set<string>): Promise<() => Promise<void>> {
  const rows = (await readCalRows()).filter((r) => !r.isGroup);
  const original = new Map(rows.map((r) => [r.path, r.shown]));
  const toggle = async (path: string, shown: boolean) => {
    const row = (await readCalRows()).find((r) => r.path === path && !r.isGroup);
    if (row && row.shown !== shown) {
      await ax.press(row.checkbox);
      await until(async () => (await readCalRows()).find((r) => r.path === path)?.shown === shown, 3000, 200);
    }
  };
  const restore = async () => {
    for (const [path, shown] of original) await toggle(path, shown).catch(() => {});
  };
  try {
    for (const r of rows) if (paths.has(r.path) && !r.shown) await toggle(r.path, true);
    for (const r of rows) if (!paths.has(r.path) && r.shown) await toggle(r.path, false);
    const now = (await readCalRows()).filter((r) => !r.isGroup);
    const missing = now.filter((r) => paths.has(r.path) && !r.shown).map((r) => r.path);
    if (missing.length) throw new OutlookError(`Could not show calendar(s): ${missing.join(", ")}`);
  } catch (e) {
    await restore();
    throw e;
  }
  await sleep(600);
  return restore;
}

// ---------------------------------------------------------------------------------------------
// Dates

const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
const monthIndex = (m: string) => MONTHS.findIndex((x) => x.startsWith(m.toLowerCase().slice(0, 3)));

export function ymd(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function parseYmd(s: string): Date {
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new OutlookError(`Expected a date like 2026-11-20, got "${s}".`);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** Parse "Thursday, 5 November 2026" or "Thursday, November 5, 2026" (and "5 Nov 2026"). */
export function parseLongDate(s: string): Date | null {
  let m = s.match(/(\d{1,2}) ([A-Za-zéû]+)\.? (\d{4})/);
  if (m && monthIndex(m[2]!) >= 0) return new Date(Number(m[3]), monthIndex(m[2]!), Number(m[1]));
  m = s.match(/([A-Za-z]+)\.? (\d{1,2}),? (\d{4})/);
  if (m && monthIndex(m[1]!) >= 0) return new Date(Number(m[3]), monthIndex(m[1]!), Number(m[2]));
  return null;
}

/** Parse the grid header ("November 2 - November 8, 2026", "December 28, 2026 - January 3, 2027"). */
export function parseRange(s: string): { start: Date; end: Date } | null {
  const m = s.replace(/[–—]/g, "-").match(/^([A-Za-z]+) (\d{1,2})(?:, (\d{4}))? - (?:([A-Za-z]+) )?(\d{1,2}), (\d{4})$/);
  if (m) {
    const endYear = Number(m[6]);
    const sm = monthIndex(m[1]!);
    const em = m[4] ? monthIndex(m[4]) : sm;
    const startYear = m[3] ? Number(m[3]) : sm > em ? endYear - 1 : endYear;
    return { start: new Date(startYear, sm, Number(m[2])), end: new Date(endYear, em, Number(m[5])) };
  }
  const one = parseLongDate(s);
  return one ? { start: one, end: one } : null;
}

function parseTimes(s: string): { start: string; end: string } | null {
  const m = s.match(/(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*(?:to|-|–)\s*(\d{1,2}:\d{2}(?:\s?[AP]M)?)/i);
  return m ? { start: to24(m[1]!), end: to24(m[2]!) } : null;
}

function to24(t: string) {
  const m = t.trim().match(/^(\d{1,2}):(\d{2})\s?([AP]M)?$/i);
  if (!m) return t;
  let h = Number(m[1]);
  if (m[3]) h = (h % 12) + (m[3].toUpperCase() === "PM" ? 12 : 0);
  return `${String(h).padStart(2, "0")}:${m[2]}`;
}

// ---------------------------------------------------------------------------------------------
// Grid navigation

async function header(): Promise<{ text: string; range: { start: Date; end: Date } }> {
  const h = await need("main", { id: "cal_grid_header_date_range" }, "the calendar date header", { depth: 2 });
  const range = parseRange(h.value ?? "");
  if (!range) throw new OutlookError(`Unrecognised calendar header "${h.value}".`);
  return { text: h.value ?? "", range };
}

async function ensureWeekView() {
  const sw = await need("main", { id: "cal_grid_header_view_switcher" }, "the calendar view switcher", { depth: 2 });
  if (sw.title === "Week") return;
  await withFocus(async () => {
    await ax.press(sw.h);
    const item = await ax.waitFor("app", { role: "AXMenuItem", id: "weekSelected" }, { timeoutMs: 3000, depth: 8 });
    if (item.ok && item.node) await ax.press(item.node.h);
    await sleep(600);
  });
  const ok = await until(async () => (await ax.findOne("main", { id: "cal_grid_header_view_switcher" }, { depth: 2 }))?.title === "Week", 4000);
  if (!ok) throw new OutlookError("Please switch Outlook's calendar to Week view (the server could not do it).");
}

/** Navigate the grid to the week containing `date`. */
async function goToWeek(date: Date) {
  await switchModule("Calendar");
  await ensureWeekView();
  for (let guard = 0; guard < 120; guard++) {
    const { text, range } = await header();
    const end = new Date(range.end.getFullYear(), range.end.getMonth(), range.end.getDate(), 23, 59);
    if (date >= range.start && date <= end) return range;
    const forward = date > end;
    // Large jumps: start from today first.
    if (guard === 0 && Math.abs(date.getTime() - range.start.getTime()) > 1000 * 86400 * 70) {
      const today = await need("main", { id: "cal_grid_header_today" }, "the Today button", { depth: 2 });
      await ax.press(today.h);
      await until(async () => (await header()).text !== text, 3000);
      continue;
    }
    const btn = await need("main", { id: forward ? "cal_grid_header_next" : "cal_grid_header_previous" }, "the week navigation buttons", { depth: 2 });
    await ax.press(btn.h);
    const changed = await until(async () => (await header()).text !== text, 4000, 100);
    if (!changed) throw new OutlookError("The calendar did not move to the next week.");
  }
  throw new OutlookError("Could not navigate the calendar to the requested date.");
}

// ---------------------------------------------------------------------------------------------
// Events

export interface EventInfo {
  ref: string;
  subject: string;
  date: string; // YYYY-MM-DD (first day)
  endDate: string | null; // for multi-day events
  start: string | null; // HH:MM, null for all-day
  end: string | null;
  allDay: boolean;
  location: string | null;
  recurrence: string | null;
  organizer: string | null;
  response: string | null; // "You accepted", "You have not responded", …
  showAs: string | null; // Busy, Tentative, Free, …
  bodyPreview: string | null;
  attendees: { accepted?: string; declined?: string; names: string[] };
  canEdit: boolean;
  canDelete: boolean;
  online: boolean;
}

interface RefData {
  v: 1;
  d: string;
  s: string;
  t: string; // start "HH:MM" or "" for all-day
  k: number;
}

const encodeRef = (r: RefData) => "evt_" + Buffer.from(JSON.stringify(r)).toString("base64url");
function decodeRef(ref: string): RefData {
  try {
    const r = JSON.parse(Buffer.from(ref.replace(/^evt_/, ""), "base64url").toString());
    if (r.v !== 1) throw new Error();
    return r;
  } catch {
    throw new OutlookError("Invalid event ref; use refs from outlook_list_events.");
  }
}

interface GridButton {
  outer: AxNode;
  inner: AxNode | null;
  label: string;
  allDay: boolean;
  day: string | null; // from the day column, timed events only
}

async function gridButtons(): Promise<GridButton[]> {
  const out: GridButton[] = [];
  const allday = await ax.findOne("main", { role: "AXGroup", descPrefix: "All day and multi-day events" }, { depth: 6 });
  if (allday) {
    for (const b of await ax.find(allday.h, { role: "AXButton", idPrefix: "107{" }, { depth: 6, limit: 200, subtree: 1 })) {
      out.push({ outer: b, inner: b.children?.find((c) => c.role === "AXButton") ?? null, label: b.desc ?? "", allDay: true, day: null });
    }
  }
  const grid = await ax.findOne("main", { id: "GridHourlyView" }, { depth: 6 });
  if (grid) {
    const days = await ax.find(grid.h, { role: "AXGroup", descRegex: "^\\w{3}, \\d{2} \\w{3} \\d{4}$" }, { depth: 3, limit: 14, subtree: 2 });
    for (const d of days) {
      const date = parseLongDate(d.desc ?? "");
      for (const b of d.children ?? []) {
        if (b.role !== "AXButton") continue;
        out.push({ outer: b, inner: b.children?.find((c) => c.role === "AXButton") ?? null, label: b.desc ?? "", allDay: false, day: date ? ymd(date) : null });
      }
    }
  }
  return out;
}

/** Open the event's peek popover in the grid, read it, close it. Works in the background. */
async function peek(btn: GridButton): Promise<{ fields: Record<string, string>; flat: AxNode[]; popover: AxNode } | null> {
  const target = btn.inner ?? btn.outer;
  try {
    await ax.press(target.h, "Show event").catch(() => ax.press(target.h));
  } catch {
    return null; // the grid re-rendered; caller re-queries
  }
  const pop = await ax.waitFor(btn.outer.h, { role: "AXPopover" }, { timeoutMs: 3000, depth: 1 });
  if (!pop.ok || !pop.node) return null;
  const tree = await ax.tree(pop.node.h, 4, { maxChildren: 80 });
  const flat: AxNode[] = [];
  const walk = (n: AxNode) => {
    flat.push(n);
    n.children?.forEach(walk);
  };
  walk(tree);
  const fields: Record<string, string> = {};
  for (const n of flat) {
    if (n.id && (n.value || n.title || n.desc)) fields[n.id] = n.value ?? n.title ?? n.desc ?? "";
  }
  return { fields, flat, popover: pop.node };
}

async function closePeek(popover: AxNode) {
  await ax.press(popover.h, "AXCancel").catch(() => {});
}

function toEvent(btn: GridButton, p: { fields: Record<string, string>; flat: AxNode[] }, k: number): EventInfo {
  const f = p.fields;
  const dateText = f.dateField ?? "";
  const dates = [...dateText.matchAll(/(?:\w+day, )?(\d{1,2} [A-Za-z]+ \d{4}|[A-Za-z]+ \d{1,2}, \d{4})/g)].map((m) => parseLongDate(m[1]!)).filter(Boolean) as Date[];
  const times = parseTimes(f.timeField ?? "") ?? parseTimes(dateText);
  const first = dates[0] ? ymd(dates[0]) : btn.day ?? "";
  const allDay = btn.allDay && !times;
  const subject = f.subjectField ?? btn.label.split(", ")[0] ?? "";
  const response = p.flat.find((n) => n.role === "AXStaticText" && /^You (have|accepted|declined|tentatively|are the organizer)/i.test(n.value ?? ""))?.value ?? null;
  const names = p.flat.filter((n) => n.role === "AXButton" && n.children?.some((c) => c.role === "AXImage")).map((n) => n.desc ?? "").filter(Boolean);
  return {
    ref: encodeRef({ v: 1, d: first, s: subject, t: allDay ? "" : times?.start ?? "", k }),
    subject,
    date: first,
    endDate: dates[1] ? ymd(dates[1]) : null,
    start: allDay ? null : times?.start ?? null,
    end: allDay ? null : times?.end ?? null,
    allDay,
    location: f.locationField ?? null,
    recurrence: f.recurrenceField ?? null,
    organizer: f.organizerField ?? null,
    response,
    showAs: f.freeBusyBar ?? null,
    bodyPreview: f.bodyField ?? null,
    attendees: { accepted: f.acceptedAttendeesField, declined: f.declinedAttendeesField, names },
    canEdit: "editEvent" in f,
    canDelete: "deleteAppointmentInstance" in f,
    online: "joinOnlineMeeting" in f,
  };
}

/** Read every event in the visible week (only for calendars currently shown). */
async function readWeek(): Promise<EventInfo[]> {
  const count = (await gridButtons()).length;
  const events: EventInfo[] = [];
  const seen = new Map<string, number>();
  const allDaySeen = new Set<string>();
  for (let i = 0; i < count; i++) {
    // The grid re-renders after a peek closes, so re-query the buttons for every event.
    let b = (await gridButtons())[i];
    if (!b) break;
    // Multi-day all-day events may render one tile per row; they share an AX identifier.
    if (b.allDay && b.outer.id) {
      if (allDaySeen.has(b.outer.id)) continue;
      allDaySeen.add(b.outer.id);
    }
    let p = await peek(b);
    if (!p) {
      b = (await gridButtons())[i];
      p = b ? await peek(b) : null;
    }
    if (!p || !b) continue;
    try {
      const probe = toEvent(b, p, 0);
      const key = `${probe.date}|${probe.subject}|${probe.start ?? ""}`;
      // A multi-day event shows a tile on each day; keep only the tile for its first day.
      if (!probe.allDay && probe.endDate && probe.endDate !== probe.date && b.day && b.day !== probe.date) continue;
      const k = seen.get(key) ?? 0;
      seen.set(key, k + 1);
      events.push(k ? toEvent(b, p, k) : probe);
    } finally {
      await closePeek(p.popover);
    }
  }
  return events;
}

export async function listEvents(opts: { start: string; end?: string; calendars?: string[] }) {
  const start = parseYmd(opts.start);
  const end = opts.end ? parseYmd(opts.end) : start;
  if (end < start) throw new OutlookError("end must not be before start.");
  if ((end.getTime() - start.getTime()) / 86400000 > 62) throw new OutlookError("Please request at most ~2 months at a time.");
  await switchModule("Calendar");
  let restore: (() => Promise<void>) | null = null;
  let shown: string[];
  if (opts.calendars?.length) {
    const rows = await Promise.all(opts.calendars.map(findCalendar));
    restore = await isolateCalendars(new Set(rows.map((r) => r.path)));
    shown = rows.map((r) => r.path);
  } else {
    shown = (await listCalendars()).filter((c) => c.shown).map((c) => c.path);
  }
  try {
    const all: EventInfo[] = [];
    const cursor = new Date(start);
    const done = new Set<string>();
    while (cursor <= end) {
      const range = await goToWeek(cursor);
      const key = ymd(range.start);
      if (!done.has(key)) {
        done.add(key);
        all.push(...(await readWeek()));
      }
      cursor.setTime(new Date(range.end.getFullYear(), range.end.getMonth(), range.end.getDate() + 1).getTime());
    }
    const s = ymd(start), e = ymd(end);
    const unique = new Map<string, EventInfo>();
    for (const ev of all) {
      const key = `${ev.date}|${ev.subject}|${ev.start ?? ""}|${ev.endDate ?? ""}|${ev.ref}`;
      if (!unique.has(key)) unique.set(key, ev);
    }
    const events = [...unique.values()]
      .filter((ev) => (ev.endDate ?? ev.date) >= s && ev.date <= e)
      .sort((a, b) => (a.date + (a.start ?? "")).localeCompare(b.date + (b.start ?? "")));
    return { calendars: shown, start: s, end: e, events };
  } finally {
    if (restore) await restore();
  }
}

/** Find an event's grid button by ref (navigates to its week). */
async function candidates(data: RefData): Promise<GridButton[]> {
  return (await gridButtons()).filter((b) => b.label.startsWith(data.s) && (b.allDay || !b.day || b.day === data.d));
}

const isRefEvent = (ev: EventInfo, data: RefData) =>
  ev.subject === data.s && (ev.start ?? "") === data.t && (ev.date === data.d || (!!ev.endDate && ev.date <= data.d && data.d <= ev.endDate));

interface LocatedEvent {
  data: RefData;
  /** Re-query the grid and open the event's peek (grid handles go stale after every peek). */
  open(): Promise<{ btn: GridButton; p: NonNullable<Awaited<ReturnType<typeof peek>>> }>;
}

/** Find an event by ref (navigates to its week). */
async function locateEvent(ref: string): Promise<LocatedEvent> {
  const data = decodeRef(ref);
  await goToWeek(parseYmd(data.d));
  await sleep(400);
  const n = (await candidates(data)).length;
  let k = 0;
  for (let i = 0; i < n; i++) {
    const b = (await candidates(data))[i];
    const p = b ? await peek(b) : null;
    if (!b || !p) continue;
    const ev = toEvent(b, p, 0);
    await closePeek(p.popover);
    if (!isRefEvent(ev, data)) continue;
    if (k++ !== data.k) continue;
    const index = i;
    return {
      data,
      async open() {
        for (let attempt = 0; attempt < 3; attempt++) {
          await sleep(300);
          const btn = (await candidates(data))[index];
          const pk = btn ? await peek(btn) : null;
          if (btn && pk) {
            if (isRefEvent(toEvent(btn, pk, 0), data)) return { btn, p: pk };
            await closePeek(pk.popover);
          }
        }
        throw new OutlookError("Could not reopen the event preview (the calendar kept changing).");
      },
    };
  }
  throw new OutlookError(`Event "${data.s}" on ${data.d}${data.t ? " at " + data.t : ""} not found (is its calendar shown?). List events again for a fresh ref.`);
}

export async function getEvent(ref: string, opts: { full?: boolean } = {}) {
  const located = await locateEvent(ref);
  const { btn, p } = await located.open();
  const ev = toEvent(btn, p, located.data.k);
  if (!opts.full) {
    await closePeek(p.popover);
    return { ...ev, ref };
  }
  // Full details: open the event window (brings Outlook forward), read it, close it.
  return withFocus(async () => {
    const before = await secondaryWindows();
    // The peek has an "Open event" button.
    const openBtn = p.flat.find((n) => n.id === "openEvent");
    if (openBtn) await ax.press(openBtn.h);
    else await ax.press((btn.inner ?? btn.outer).h, "Open event");
    const win = await until(async () => (await secondaryWindows()).find((w) => !before.some((b) => b.title === w.title)) ?? null, 6000);
    if (!win) throw new OutlookError("The event window did not open.");
    try {
      const tree = await ax.tree(win.h, 8, { maxChildren: 120 });
      const flat: AxNode[] = [];
      const walk = (n: AxNode) => {
        flat.push(n);
        n.children?.forEach(walk);
      };
      walk(tree);
      const titleParts = (win.title ?? "").split(" • ");
      const attendees = flat.filter((n) => n.role === "AXButton" && /, (Required|Optional)$/.test(n.title ?? "")).map((n) => n.title!);
      const bodyArea = await ax.findOne(win.h, { role: "AXWebArea" }, { depth: 12 });
      const windowBody = bodyArea ? (await ax.text(bodyArea.h)).text.trim() : "";
      // The window may keep the description behind a "Meeting Details" tab; fall back to the preview's text.
      const body = windowBody || ev.bodyPreview;
      return {
        ...ev,
        ref,
        calendar: titleParts.length >= 3 ? `${titleParts[2]}/${titleParts[1]}` : null,
        when: flat.find((n) => n.id === "recurrenceField")?.value ?? null,
        attendeeList: attendees,
        body,
      };
    } finally {
      const close = await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 });
      if (close) await ax.press(close.h);
    }
  });
}

// ---------------------------------------------------------------------------------------------
// Event form (create / edit)

export interface EventInput {
  subject?: string;
  start?: string; // ISO local "2026-11-20T10:00" (or with offset) / date "2026-11-20" for all-day
  end?: string;
  allDay?: boolean;
  location?: string;
  body?: string;
  calendar?: string;
}

/** Interpret "2026-11-20T10:00" as local time; accepts explicit offsets too. */
function parseLocal(s: string): Date {
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0));
  const d = new Date(s);
  if (isNaN(d.getTime())) throw new OutlookError(`Invalid date/time "${s}"; use e.g. 2026-11-20T10:00.`);
  return d;
}

async function setDateField(win: AxNode, title: string, d: Date) {
  const f = await ax.findOne(win.h, { role: "AXDateTimeArea", title }, { depth: 6 });
  if (!f) return false;
  await ax.call("set", { h: f.h, attr: "AXValue", value: d.toISOString().replace(/\.\d{3}Z$/, "Z"), type: "date" });
  // Read back: date fields keep their time component and vice versa, so compare the relevant part.
  const ok = await until(async () => {
    const v = (await ax.get(f.h)).value;
    if (!v) return false;
    const got = new Date(v);
    return /date/i.test(title) ? ymd(got) === ymd(d) : got.getHours() === d.getHours() && got.getMinutes() === d.getMinutes();
  }, 2000, 150);
  if (!ok) throw new OutlookError(`Outlook did not accept ${title} = ${d.toString()}.`);
  return true;
}

async function chooseCalendar(win: AxNode, cal: CalRow) {
  const picker = await need(win.h, { role: "AXPopUpButton", descRegex: ", From$" }, "the calendar picker", { depth: 4 });
  if ((picker.desc ?? "").startsWith(`${cal.name} (${cal.account})`)) return;
  await ax.press(picker.h);
  const table = await until(async () => {
    const menuWin = (await ax.children("app")).find((w) => w.title === "menuWindow");
    return menuWin ? ax.findOne(menuWin.h, { role: "AXTable" }, { depth: 4 }) : null;
  }, 3000);
  if (!table) throw new OutlookError("The calendar picker did not open.");
  const cells = await ax.find(table.h, { role: "AXCell" }, { depth: 2, limit: 300 });
  let account = "";
  const accounts = new Set((await readCalRows()).filter((r) => r.level === 0).map((r) => r.name.toLowerCase()));
  let target: AxNode | null = null;
  for (const c of cells) {
    const label = (c.desc ?? "").replace(/^Ticked, /, "");
    if (accounts.has(label.toLowerCase())) {
      account = label;
      continue;
    }
    if (account.toLowerCase() === cal.account.toLowerCase() && label === cal.name) {
      target = c;
      break;
    }
  }
  if (!target) {
    await ax.key(KEY.ESCAPE);
    throw new OutlookError(`Calendar "${cal.path}" is not offered in the event's calendar picker (it may be read-only).`);
  }
  // Selecting the row first, then pressing the cell, is what makes the SwiftUI picker commit.
  const row = await ax.call<AxNode>("parent", { h: target.h });
  await ax.set(row.h, "AXSelected", true).catch(() => {});
  await sleep(150);
  await ax.press(target.h);
  const applied = async () => ((await ax.get(picker.h)).desc ?? "").startsWith(`${cal.name} (${cal.account})`);
  if (!(await until(applied, 2000, 200))) {
    if ((await ax.children("app")).some((w) => w.title === "menuWindow")) await ax.key(KEY.RETURN);
    if (!(await until(applied, 2000, 200))) {
      if ((await ax.children("app")).some((w) => w.title === "menuWindow")) await ax.key(KEY.ESCAPE);
      throw new OutlookError(`Could not select calendar "${cal.path}" in the event form.`);
    }
  }
}

/** Fill the event form; the window must be key (Outlook focused). */
async function fillForm(win: AxNode, input: EventInput, isNew: boolean) {
  // Make sure the form is the key window so typing lands in it.
  await ax.press(win.h, "AXRaise").catch(() => {});
  await ax.set(win.h, "AXMain", true).catch(() => {});
  await until(async () => (await ax.get("focusedWindow").catch(() => null))?.title === win.title, 2000, 200);
  if (input.calendar) {
    if (!isNew) throw new OutlookError("Moving an existing event to another calendar is not supported.");
    await chooseCalendar(win, await findCalendar(input.calendar));
  }
  if (input.subject !== undefined) {
    // Setting AXValue only changes the display; type it so Outlook's model sees the edit.
    const field = () => need(win.h, { role: "AXTextField", desc: "Subject" }, "the Subject field", { depth: 4 });
    const read = async () => (await field()).value ?? "";
    let ok = false;
    for (let attempt = 0; attempt < 2 && !ok; attempt++) {
      const subj = await field();
      const focused = await until(async () => {
        await ax.set(subj.h, "AXFocused", true);
        return (await ax.get(subj.h, { attrs: ["AXFocused"] })).attrs?.AXFocused === true;
      }, 3000, 250);
      if (!focused) continue;
      await ax.menu(["Edit", "Select All"]).catch(() => {});
      await ax.type(input.subject); // typing (not AXValue) so the edit reaches Outlook's model
      ok = !!(await until(async () => (await read()) === input.subject, 3000, 200));
    }
    if (!ok) throw new OutlookError(`Could not enter the event subject (field reads "${await read()}").`);
  }
  if (input.allDay !== undefined) {
    const sw = await ax.findOne(win.h, { role: "AXCheckBox", title: "All day" }, { depth: 4 });
    if (sw && (sw.value === "1") !== input.allDay) {
      await ax.press(sw.h);
      await sleep(400);
    }
  }
  const allDay = input.allDay ?? (await ax.findOne(win.h, { role: "AXCheckBox", title: "All day" }, { depth: 4 }))?.value === "1";
  if (input.allDay !== undefined && !(await ax.findOne(win.h, { role: "AXCheckBox", title: "All day" }, { depth: 4 }))) {
    throw new OutlookError("The event form has no All day switch.");
  }
  if (input.start) {
    const s = parseLocal(input.start);
    if (!(await ax.findOne(win.h, { role: "AXDateTimeArea", title: "Start date" }, { depth: 6 }))) throw new OutlookError("The event form has no Start date field.");
    // Setting the date first keeps Outlook's duration logic predictable.
    await setDateField(win, "Start date", s);
    if (!allDay) await setDateField(win, "Start time", s);
    await sleep(250);
  }
  if (input.end) {
    const e = parseLocal(input.end);
    const s = input.start ? parseLocal(input.start) : null;
    const hasEndDate = await setDateField(win, "End date", e);
    if (!hasEndDate && s && ymd(s) !== ymd(e) && !allDay) {
      throw new OutlookError("Outlook's form has no End date field for timed events; multi-day timed events are not supported.");
    }
    if (!allDay) await setDateField(win, "End time", e);
    await sleep(250);
  }
  if (input.location !== undefined) {
    const loc = await need(win.h, { role: "AXTextField", descRegex: "^(Add a location|Locations?)$" }, "the location field", { depth: 6 });
    await ax.set(loc.h, "AXFocused", true);
    await sleep(150);
    await ax.menu(["Edit", "Select All"]).catch(() => {});
    await ax.type(input.location);
    await sleep(400);
    await ax.key(KEY.ESCAPE); // dismiss room suggestions
    await sleep(200);
  }
  if (input.body !== undefined) {
    const target = await bodyEditor(win.h);
    // Paste only changes the editor's DOM; typed keystrokes reach Outlook's model reliably.
    const focused = await until(async () => {
      await ax.set(target.h, "AXFocused", true);
      return (await ax.get(target.h, { attrs: ["AXFocused"] })).attrs?.AXFocused === true;
    }, 4000, 300);
    if (!focused) throw new OutlookError("Could not focus the event description editor.");
    await ax.menu(["Edit", "Select All"]).catch(() => {});
    await ax.key(KEY.DELETE);
    const lines = input.body.replace(/\r\n/g, "\n").split("\n");
    for (let i = 0; i < lines.length; i++) {
      if (lines[i]) await ax.type(lines[i]!);
      if (i < lines.length - 1) await ax.key(KEY.RETURN);
    }
    const probe = input.body.trim().split(/\n/)[0]!.slice(0, 30);
    if (probe && !(await until(async () => (await ax.text(target.h)).text.includes(probe), 4000, 250))) {
      throw new OutlookError("The description text did not appear in the event form.");
    }
    // The description editor commits on blur: move focus to the subject field before saving.
    const subj = await ax.findOne(win.h, { role: "AXTextField", desc: "Subject" }, { depth: 4 });
    if (subj) await ax.set(subj.h, "AXFocused", true);
    await sleep(1000);
  }
  if (input.location !== undefined) {
    await checkLocation(win, input.location);
  }
}

async function checkLocation(win: AxNode, location: string) {
  const loc = await ax.findOne(win.h, { role: "AXTextField", descRegex: "^(Add a location|Locations?)$" }, { depth: 6 });
  // A confirmed location becomes a token, which reads as U+FFFC.
  const v = (loc?.value ?? "").trim();
  if (loc && location.trim() && v !== location.trim() && !v.includes("\uFFFC")) {
    throw new OutlookError(`The location field reads "${loc.value ?? ""}" instead of "${location}".`);
  }
}

async function saveForm(win: AxNode) {
  const save = await ax.findOne(win.h, { role: "AXButton", descRegex: "^Save(,|$)" }, { depth: 3 });
  if (!save) {
    const send = await ax.findOne(win.h, { role: "AXButton", descRegex: "^Send(,|$)" }, { depth: 3 });
    if (send) throw new OutlookError("This event has attendees, so Outlook would send invitations/updates. The server only saves events without notifying anyone; the form was left open.");
    throw new OutlookError("Could not find the event's Save button.");
  }
  await ax.press(save.h);
  // Recurring events ask whether to change one occurrence or the series.
  const prompt = await ax.waitFor(win.h, { role: "AXSheet" }, { timeoutMs: 1200, depth: 1 }).catch(() => null);
  if (prompt?.ok && prompt.node) {
    const buttons = await ax.find(prompt.node.h, { role: "AXButton" }, { depth: 2 });
    const occ = buttons.find((b) => /occurrence|this event/i.test(b.title ?? ""));
    if (occ) await ax.press(occ.h);
  }
  const closed = await until(async () => !(await secondaryWindows()).some((w) => w.title === win.title || w.h === win.h), 8000);
  if (!closed) throw new OutlookError("The event window did not close after saving; check Outlook for a validation message.");
}

/** After saving, wait until the grid shows a tile with this subject on that day (Outlook redraws lazily). */
async function waitForTile(day: Date, subject: string): Promise<boolean> {
  await goToWeek(day);
  const d = ymd(day);
  return !!(await until(async () => (await gridButtons()).some((b) => b.label.startsWith(subject) && (b.allDay || b.day === d)), 8000, 400));
}

/** Close an event form without saving, answering a "discard changes?" prompt. */
async function discardForm(win: AxNode) {
  const discard = await ax.findOne(win.h, { role: "AXButton", descRegex: "^(Discard|Close)(,|$)" }, { depth: 3 }).catch(() => null);
  const close = discard ?? (await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 }).catch(() => null));
  if (!close) return;
  await ax.press(close.h).catch(() => {});
  const sheet = await ax.waitFor(win.h, { role: "AXSheet" }, { timeoutMs: 1500, depth: 1 }).catch(() => null);
  if (sheet?.ok && sheet.node) {
    const btn = await ax.findOne(sheet.node.h, { role: "AXButton", titleRegex: "^(Discard|Don.t Save|Delete)" }, { depth: 2 });
    if (btn) await ax.press(btn.h).catch(() => {});
  }
}

export async function createEvent(input: EventInput & { subject: string; start: string; end?: string }) {
  if (!input.end && !input.allDay) throw new OutlookError("end is required for timed events.");
  await switchModule("Calendar");
  const s = parseLocal(input.start);
  return withFocus(async () => {
    const before = await secondaryWindows();
    const newBtn = await need("main", { role: "AXButton", descRegex: "^New Event(,|$)" }, "the New Event button", { depth: 2 });
    await ax.press(newBtn.h);
    const win = await until(async () => (await secondaryWindows()).find((w) => (w.title ?? "").startsWith("New Event") && !before.some((b) => b.title === w.title)) ?? null, 8000);
    if (!win) throw new OutlookError("The New Event window did not open.");
    try {
      await fillForm(win, { ...input, allDay: input.allDay ?? false }, true);
      const title = (await ax.get(win.h)).title ?? "";
      await saveForm({ ...win, title });
      const cal = title.split(" • ");
      const verified = await waitForTile(s, input.subject);
      return {
        created: true,
        verified,
        ref: encodeRef({ v: 1, d: ymd(s), s: input.subject, t: input.allDay ? "" : `${String(s.getHours()).padStart(2, "0")}:${String(s.getMinutes()).padStart(2, "0")}`, k: 0 }),
        calendar: cal.length >= 3 ? `${cal[2]}/${cal[1]}` : null,
      };
    } catch (e) {
      // Leave nothing half-created behind.
      if (!(e instanceof OutlookError && /attendees/.test(e.message))) await discardForm(win);
      throw e;
    }
  });
}

export async function updateEvent(ref: string, changes: EventInput) {
  const located = await locateEvent(ref);
  const data = located.data;
  const { p } = await located.open();
  const edit = p.flat.find((n) => n.id === "editEvent");
  if (!edit) {
    await closePeek(p.popover);
    throw new OutlookError("This event cannot be edited (you are probably not its organizer).");
  }
  return withFocus(async () => {
    const before = await secondaryWindows();
    await ax.press(edit.h);
    // Recurring events: Outlook may ask "this event or the series?" before opening the editor.
    const occ = await ax.waitFor("app", { role: "AXButton", titleRegex: "(?i)(this event|this occurrence|^Occurrence)" }, { timeoutMs: 1500, depth: 6 }).catch(() => null);
    if (occ?.ok && occ.node) await ax.press(occ.node.h);
    const win = await until(async () => (await secondaryWindows()).find((w) => !before.some((b) => b.title === w.title)) ?? null, 8000);
    if (!win) throw new OutlookError("The event editor did not open.");
    try {
      await fillForm(win, changes, false);
      const title = (await ax.get(win.h)).title ?? "";
      await saveForm({ ...win, title });
    } catch (e) {
      if (!(e instanceof OutlookError && /attendees/.test(e.message))) await discardForm(win);
      throw e;
    }
    const s = changes.start ? parseLocal(changes.start) : null;
    const allDay = changes.allDay ?? data.t === "";
    const verified = await waitForTile(s ?? parseYmd(data.d), changes.subject ?? data.s);
    return {
      updated: true,
      verified,
      ref: encodeRef({
        v: 1,
        d: s ? ymd(s) : data.d,
        s: changes.subject ?? data.s,
        t: allDay ? "" : s ? `${String(s.getHours()).padStart(2, "0")}:${String(s.getMinutes()).padStart(2, "0")}` : data.t,
        k: 0,
      }),
    };
  });
}

export async function deleteEvent(ref: string) {
  const located = await locateEvent(ref);
  const data = located.data;
  const { p } = await located.open();
  const del = p.flat.find((n) => n.id === "deleteAppointmentInstance");
  if (!del) {
    await closePeek(p.popover);
    throw new OutlookError("This event cannot be deleted from its preview (meetings organized by others must be declined instead).");
  }
  const key = (w: AxNode) => `${w.title}|${w.subrole}`;
  const winsBefore = new Set((await ax.call<AxNode[]>("windows")).map(key));
  const sheetsBefore = (await ax.find("main", { role: "AXSheet" }, { depth: 1, limit: 10 })).length;
  const countBefore = (await gridButtons()).filter((b) => b.label.startsWith(data.s) && (b.allDay || b.day === data.d)).length;
  await ax.press(del.h);
  // If Outlook asks anything (e.g. whether to notify attendees), refuse and back out.
  const asked = await until(async () => {
    const extra = (await ax.call<AxNode[]>("windows")).filter((w) => w.id !== "OutlookWindow" && !winsBefore.has(key(w)));
    const sheets = await ax.find("main", { role: "AXSheet" }, { depth: 1, limit: 10 });
    if (extra.length || sheets.length > sheetsBefore) return { extra, sheet: sheets[sheets.length - 1] ?? null };
    const still = (await gridButtons()).filter((b) => b.label.startsWith(data.s) && (b.allDay || b.day === data.d)).length;
    return still < countBefore ? { extra: [], sheet: null, gone: true } : null;
  }, 2500, 200);
  if (asked && !("gone" in asked)) {
    for (const w of asked.extra) {
      const cancel = await ax.findOne(w.h, { role: "AXButton", titleRegex: "^Cancel$" }, { depth: 3 });
      const close = cancel ?? (await ax.findOne(w.h, { subrole: "AXCloseButton" }, { depth: 1 }));
      if (close) await ax.press(close.h);
    }
    if (asked.sheet) {
      const cancel = await ax.findOne(asked.sheet.h, { role: "AXButton", titleRegex: "^Cancel$" }, { depth: 2 });
      if (cancel) await ax.press(cancel.h);
    }
    await sleep(500);
    const still = (await gridButtons()).filter((b) => b.label.startsWith(data.s) && (b.allDay || b.day === data.d)).length;
    if (still < countBefore) return { deleted: true, notified: "unknown", subject: data.s, date: data.d };
    throw new OutlookError("Outlook asked for confirmation (probably to notify attendees). Nothing was deleted.");
  }
  const gone = await until(async () => {
    const still = (await gridButtons()).filter((b) => b.label.startsWith(data.s) && (b.allDay || b.day === data.d)).length;
    return still < countBefore ? true : null;
  }, 4000, 300);
  if (!gone) throw new OutlookError("The event still appears in the calendar; it may not have been deleted.");
  return { deleted: true, subject: data.s, date: data.d };
}
