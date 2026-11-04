// Mail: folders, message lists, search, reading, .eml export, moving and flags — all through Outlook's UI.
import { existsSync, mkdirSync, renameSync, statSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { ax, bodyEditor, need, OutlookError, sleep, switchModule, until, withFocus } from "./session";
import { KEY, type AxNode } from "../ax";

// ---------------------------------------------------------------------------------------------
// Folders

export interface Folder {
  path: string; // "account/Parent/Child"; the unified section is "All Accounts/…"
  name: string;
  account: string;
  depth: number;
  unread: number | null;
  total: number | null;
  hasChildren: boolean;
  expanded: boolean;
}

interface FolderRow extends Folder {
  h: string;
  section: string;
  selected: boolean;
}

async function folderOutline(): Promise<AxNode> {
  await switchModule("Mail");
  return need("main", { role: "AXOutline" }, "the folder sidebar", { depth: 4 });
}

/** Parse the sidebar into rows. Section headers (accounts, "Favorites", "All Accounts") are depth 0. */
async function readFolderRows(): Promise<FolderRow[]> {
  const outline = await folderOutline();
  const rows = await ax.find(outline.h, { role: "AXRow" }, { depth: 1, limit: 2000, subtree: 3, extra: ["AXDisclosureLevel", "AXDisclosing"] });
  const out: FolderRow[] = [];
  const stack: string[] = [];
  let section = "";
  for (const row of rows) {
    const level = Number(row.x?.AXDisclosureLevel ?? 0);
    const cell = row.children?.find((c) => c.role === "AXCell");
    const desc = cell?.desc ?? "";
    if (!desc) continue; // separators
    const inner = cell?.children?.find((c) => c.role === "AXUnknown");
    const tri = cell?.children?.find((c) => c.role === "AXDisclosureTriangle");
    const name = desc.split(";")[0]!.trim();
    if (level === 0) {
      section = name;
      stack.length = 0;
      continue;
    }
    stack.length = level - 1;
    stack.push(name);
    const unreadM = desc.match(/;\s*(\d+) unread/);
    const countM = desc.match(/;\s*(\d+) messages?$/); // Drafts/Junk show a plain message count
    const totalM = (inner?.help ?? "").match(/ - (\d+) Items?$/);
    out.push({
      h: row.h,
      selected: !!row.selected,
      section,
      account: section,
      path: [section, ...stack].join("/"),
      name,
      depth: level,
      unread: unreadM ? Number(unreadM[1]) : countM ? 0 : 0,
      total: totalM ? Number(totalM[1]) : countM ? Number(countM[1]) : null,
      hasChildren: !!tri,
      expanded: tri ? tri.value === "1" : false,
    });
  }
  return out;
}

/** Expand every collapsed folder (and account) so nested folders become visible; returns rows to re-collapse. */
async function expandAll(): Promise<string[]> {
  const expandedNow: string[] = [];
  for (let pass = 0; pass < 6; pass++) {
    const outline = await folderOutline();
    const rows = await ax.find(outline.h, { role: "AXRow" }, { depth: 1, limit: 2000, subtree: 3, extra: ["AXDisclosing", "AXDisclosureLevel"] });
    const collapsed = rows.filter((r) => {
      const cell = r.children?.find((c) => c.role === "AXCell");
      const tri = cell?.children?.find((c) => c.role === "AXDisclosureTriangle");
      const name = cell?.desc?.split(";")[0] ?? "";
      // Leave "Favorites" alone; it only duplicates folders.
      return tri && tri.value === "0" && name !== "Favorites" && name !== "Groups";
    });
    if (!collapsed.length) break;
    for (const r of collapsed) {
      const cell = r.children!.find((c) => c.role === "AXCell")!;
      expandedNow.push(`${r.x?.AXDisclosureLevel}|${cell.desc!.split(";")[0]}`);
      await ax.set(r.h, "AXDisclosing", true);
    }
    await sleep(500);
  }
  return expandedNow;
}

async function collapse(keys: string[]) {
  if (!keys.length) return;
  const outline = await folderOutline();
  const rows = await ax.find(outline.h, { role: "AXRow" }, { depth: 1, limit: 2000, subtree: 2, extra: ["AXDisclosureLevel"] });
  // Collapse deepest first so parents do not hide children before they are processed.
  const want = new Set(keys);
  const targets = rows
    .filter((r) => want.has(`${r.x?.AXDisclosureLevel}|${(r.children?.[0]?.desc ?? "").split(";")[0]}`))
    .sort((a, b) => Number(b.x?.AXDisclosureLevel) - Number(a.x?.AXDisclosureLevel));
  for (const r of targets) await ax.set(r.h, "AXDisclosing", false);
}

let folderCache: { at: number; rows: FolderRow[] } | null = null;

export async function listFolders(opts: { expand?: boolean; includeFavorites?: boolean } = {}): Promise<Folder[]> {
  let reopened: string[] = [];
  if (opts.expand) reopened = await expandAll();
  try {
    const rows = await readFolderRows();
    folderCache = { at: Date.now(), rows };
    return rows
      .filter((r) => opts.includeFavorites || r.section !== "Favorites")
      .map(({ h, section, selected, ...f }) => f);
  } finally {
    if (reopened.length) await collapse(reopened).catch(() => {});
  }
}

function normalizePath(p: string) {
  return p.trim().replace(/^\/+|\/+$/g, "").toLowerCase();
}

/** Resolve a user-supplied folder (full path, or a unique folder name) to its sidebar row. */
async function findFolderRow(folder: string): Promise<FolderRow> {
  const want = normalizePath(folder);
  const attempt = async (rows: FolderRow[]) => {
    const exact = rows.filter((r) => r.section !== "Favorites" && normalizePath(r.path) === want);
    if (exact.length) return exact[0]!;
    const byName = rows.filter((r) => r.section !== "Favorites" && r.section !== "All Accounts" && (r.name.toLowerCase() === want || normalizePath(r.path).endsWith("/" + want)));
    if (byName.length === 1) return byName[0]!;
    if (byName.length > 1) {
      throw new OutlookError(`Folder "${folder}" is ambiguous; use a full path, e.g. ${byName.slice(0, 4).map((r) => `"${r.path}"`).join(", ")}`);
    }
    return null;
  };
  let row = await attempt(await readFolderRows());
  if (row) return row;
  // Maybe nested under a collapsed parent.
  await expandAll();
  row = await attempt(await readFolderRows());
  if (row) return row;
  throw new OutlookError(`Folder "${folder}" not found. Use outlook_list_folders to see valid paths.`);
}

/**
 * Refuse to navigate away while the user is writing a message inline in the main window. When Outlook is in
 * the background nobody is typing there, and Outlook autosaves drafts when the selection changes.
 */
async function assertNoInlineDraft() {
  const front = await ax.call<{ bundleId?: string }>("frontmost");
  if (front.bundleId !== "com.microsoft.Outlook") return;
  const draft = await ax.findOne("main", { id: "subjectTextField" }, { depth: 10 });
  if (draft) {
    throw new OutlookError(
      `A draft ("${draft.value || "no subject"}") is open in Outlook's main window. Finish, close or pop it out first; the server will not navigate away from a message you are writing.`,
    );
  }
}

/** Open a folder in the main window; returns its row. Leaves any search mode. */
export async function openFolder(folder: string): Promise<Folder> {
  await assertNoInlineDraft();
  await exitSearch();
  const row = await findFolderRow(folder);
  await ax.set(row.h, "AXSelected", true);
  // Window title is "<folder> • <account>" (the unified view may omit the account).
  const ok = await until(async () => {
    const t = (await ax.get("main")).title ?? "";
    if (row.section === "All Accounts") return t === row.name || t.startsWith(row.name + " •");
    return t === `${row.name} • ${row.account}`;
  }, 6000);
  if (!ok) throw new OutlookError(`Outlook did not open folder "${row.path}".`);
  await waitListSettled();
  const { h, section, selected, ...f } = row;
  return f;
}

// ---------------------------------------------------------------------------------------------
// Message list

export interface MessageSummary {
  ref: string;
  index: number;
  sender: string;
  subject: string;
  date: string;
  preview: string;
  unread: boolean;
  flagged: boolean;
  replied: boolean;
  forwarded: boolean;
  hasAttachments: boolean;
  meeting: boolean; // meeting invitation / response
  completed: boolean; // flag marked as completed
  conversationCount: number | null;
  unreadInConversation: number | null;
  folder: string | null; // for search results
  group: string | null; // "Today", "Yesterday", "Last Week"…
}

interface RowInfo extends Omit<MessageSummary, "ref" | "index"> {
  h: string;
}

const DATE_RE = /(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{1,2}:\d{2}(?: ?[AP]M)?|Yesterday|Today|(?:Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day(?: \d{1,2}[/.-]\d{1,2})?|(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) \d{1,2}[/.-]\d{1,2})/;

/**
 * Parse the accessibility description Outlook gives each message row, e.g.
 *   " Unread,    Sender: A, Subject: B,      Yesterday,  Flagged,      Message preview: …, "
 *   "   Conversation, 5 messages, Sender: A, Subject: B,   Replied Latest message:  23/09/2026, …"
 * Fields after the subject are separated by a comma plus 2+ spaces; markers are only looked for
 * outside the subject and preview so their text cannot cause false positives.
 */
export function parseRowDesc(desc: string): Omit<RowInfo, "h" | "group"> | null {
  const si = desc.indexOf("Sender: ");
  if (si < 0) return null;
  const prefix = desc.slice(0, si);
  const afterSender = desc.slice(si + "Sender: ".length);
  let sender: string, subject: string, rest: string;
  const subjIdx = afterSender.indexOf(", Subject: ");
  const noSubjIdx = afterSender.search(/, No subject(,|$)/);
  if (subjIdx >= 0 && (noSubjIdx < 0 || subjIdx < noSubjIdx)) {
    sender = afterSender.slice(0, subjIdx);
    const body = afterSender.slice(subjIdx + ", Subject: ".length);
    const end = body.search(/,\s{2,}/);
    subject = end >= 0 ? body.slice(0, end) : body.replace(/,\s*$/, "");
    rest = end >= 0 ? body.slice(end) : "";
  } else if (noSubjIdx >= 0) {
    sender = afterSender.slice(0, noSubjIdx);
    subject = "";
    rest = afterSender.slice(noSubjIdx + ", No subject".length);
  } else {
    return null;
  }
  const pi = rest.indexOf("Message preview: ");
  const mid = (pi >= 0 ? rest.slice(0, pi) : rest).replace(/\s+/g, " ");
  const preview = pi >= 0 ? rest.slice(pi + "Message preview: ".length).replace(/,\s*$/, "").replace(/\s+/g, " ").trim() : "";
  const convM = prefix.match(/Conversation, (\d+) messages?/);
  const unreadConvM = prefix.match(/(\d+) unread messages?/);
  return {
    sender: sender.trim(),
    subject: subject.trim(),
    date: mid.match(DATE_RE)?.[1] ?? "",
    preview,
    unread: /\bUnread\b/.test(prefix) || !!unreadConvM,
    flagged: /\bFlagged\b/.test(mid),
    replied: /\bReplied\b/.test(mid),
    forwarded: /\bForwarded\b/.test(mid),
    hasAttachments: /\b(Has files|Has attachments?)\b/.test(mid),
    meeting: /\bMeeting message\b/.test(mid),
    completed: /\bMarked as completed\b/.test(mid),
    conversationCount: convM ? Number(convM[1]) : null,
    unreadInConversation: unreadConvM ? Number(unreadConvM[1]) : null,
    folder: mid.match(/Folder: ([^,]+)/)?.[1]?.trim() ?? null,
  };
}

async function messageTable(): Promise<AxNode> {
  return need("main", { role: "AXTable", desc: "Message List" }, "the message list", { depth: 8, timeoutMs: 4000 });
}

async function waitListSettled() {
  // The list re-renders after a folder switch or search; wait until two consecutive reads agree.
  let last = "";
  await until(async () => {
    const t = await ax.findOne("main", { role: "AXTable", desc: "Message List" }, { depth: 8 });
    if (!t) return false;
    const cells = await ax.find(t.h, { role: "AXCell" }, { depth: 2, limit: 6 });
    const sig = cells.map((c) => c.desc ?? "").join("|") + `#${cells.length}`;
    const stable = sig === last;
    last = sig;
    return stable;
  }, 4000, 250);
}

function isGroupHeader(desc: string) {
  return /^[^,]+, (Expanded|Collapsed)$/.test(desc.trim());
}

async function readRows(table: AxNode): Promise<RowInfo[]> {
  // AXRows lists every loaded row; AXChildren only those currently built around the viewport.
  const rowsEl = await ax.call<AxNode[]>("elements", { h: table.h, attr: "AXRows", subtree: 1 });
  const rows: RowInfo[] = [];
  let group: string | null = null;
  for (const r of rowsEl) {
    const desc = r.children?.find((c) => c.role === "AXCell")?.desc ?? "";
    if (isGroupHeader(desc)) {
      group = desc.split(",")[0]!.trim();
      continue;
    }
    const p = parseRowDesc(desc);
    if (p) rows.push({ ...p, group, h: r.h });
  }
  return rows;
}

async function cellCount(): Promise<number> {
  const t = await ax.findOne("main", { role: "AXTable", desc: "Message List" }, { depth: 8 });
  return t ? (await ax.call<AxNode[]>("elements", { h: t.h, attr: "AXRows" })).length : 0;
}

async function scrollBar(table: AxNode): Promise<AxNode | null> {
  const area = await ax.call<AxNode>("parent", { h: table.h });
  return ax.findOne(area.h, { role: "AXScrollBar" }, { depth: 1 });
}

/** Load rows until at least `want` are available (scrolling to the end triggers lazy loading). */
async function loadRows(want: number): Promise<RowInfo[]> {
  let table = await messageTable();
  let rows = await readRows(table);
  const deadline = Date.now() + 90_000;
  for (let pass = 0; rows.length < want && pass < 40 && Date.now() < deadline; pass++) {
    const bar = await scrollBar(table);
    if (!bar) break;
    const before = await cellCount();
    await ax.set(bar.h, "AXValue", 1, "number");
    const grew = await until(async () => ((await cellCount()) > before ? true : null), 3500, 300);
    table = await messageTable();
    const next = await readRows(table);
    if (!grew || next.length <= rows.length) {
      rows = next.length > rows.length ? next : rows;
      break; // nothing more to load
    }
    rows = next;
  }
  const bar = await scrollBar(table);
  if (bar) await ax.set(bar.h, "AXValue", 0, "number").catch(() => {});
  return rows;
}

// Message references ------------------------------------------------------------------------

interface RefData {
  v: 2;
  f: string; // folder path
  q?: string; // search query (results view)
  s: string; // subject
  n: string; // sender
  d: string; // date label
  p: string; // start of the preview, to tell apart same-subject rows
  k: number; // occurrence among identical rows
}

function encodeRef(r: RefData) {
  return "msg_" + Buffer.from(JSON.stringify(r)).toString("base64url");
}

function decodeRef(ref: string): RefData {
  try {
    const r = JSON.parse(Buffer.from(ref.replace(/^msg_/, ""), "base64url").toString()) as RefData;
    if (r.v !== 2 || typeof r.s !== "string") throw new Error();
    return r;
  } catch {
    throw new OutlookError(`Invalid message ref "${ref.slice(0, 40)}…". Use refs returned by outlook_list_messages or outlook_search_messages.`);
  }
}

const previewKey = (preview: string) => preview.slice(0, 60);
const sameRow = (r: RowInfo, d: { s: string; n: string; d: string; p: string }) =>
  r.subject === d.s && r.sender === d.n && r.date === d.d && previewKey(r.preview) === d.p;

function toSummaries(rows: RowInfo[], f: string, q?: string): MessageSummary[] {
  const seen = new Map<string, number>();
  return rows.map((r, i) => {
    const p = previewKey(r.preview);
    const key = `${r.subject}\u0000${r.sender}\u0000${r.date}\u0000${p}`;
    const k = seen.get(key) ?? 0;
    seen.set(key, k + 1);
    const { h, ...rest } = r;
    return { ref: encodeRef({ v: 2, f, q, s: r.subject, n: r.sender, d: r.date, p, k }), index: i, ...rest };
  });
}

export async function listMessages(opts: { folder: string; limit?: number; offset?: number; unreadOnly?: boolean }) {
  const limit = Math.min(opts.limit ?? 25, 500);
  const offset = opts.offset ?? 0;
  const folder = await openFolder(opts.folder);
  const rows = await loadRows(offset + limit);
  let list = toSummaries(rows, folder.path);
  if (opts.unreadOnly) list = list.filter((m) => m.unread);
  return { folder: folder.path, totalInFolder: folder.total, unreadInFolder: folder.unread, loaded: rows.length, messages: list.slice(offset, offset + limit) };
}

async function searchField(): Promise<AxNode> {
  return need("main", { role: "AXTextField", subrole: "AXSearchField", id: "Search Bar" }, "the search box", { depth: 3 });
}

/** Leave search mode. Selecting a folder ends it, but re-selecting the selected folder does not, so hop. */
async function exitSearch() {
  const title = (await ax.get("main")).title ?? "";
  if (!title.startsWith("Searching")) return;
  const field = await ax.findOne("main", { role: "AXTextField", subrole: "AXSearchField", id: "Search Bar" }, { depth: 3 });
  if (field) await ax.set(field.h, "AXValue", "");
  const rows = await readFolderRows();
  const hops = rows.filter((r) => !r.selected).sort((a, b) => Number(a.section === "Favorites") - Number(b.section === "Favorites"));
  for (const hop of hops.slice(0, 3)) {
    await ax.set(hop.h, "AXSelected", true);
    if (await until(async () => !((await ax.get("main")).title ?? "").startsWith("Searching"), 4000)) return;
  }
  throw new OutlookError("Could not leave Outlook's search results view.");
}

async function runSearch(query: string) {
  const field = await searchField();
  await ax.set(field.h, "AXValue", query);
  await ax.press(field.h, "AXConfirm");
  const ok = await until(async () => ((await ax.get("main")).title ?? "").startsWith("Searching"), 6000);
  if (!ok) throw new OutlookError("Outlook did not start the search.");
  // Results arrive after a few seconds: wait for rows or for Outlook's "No Search Results" placeholder.
  await until(async () => {
    const t = await ax.findOne("main", { role: "AXTable", desc: "Message List" }, { depth: 8 });
    if (t) {
      const cells = await ax.find(t.h, { role: "AXCell", descContains: "Sender:" }, { depth: 2, limit: 1 });
      if (cells.length) return true;
    }
    return !!(await ax.findOne("main", { descContains: "No Search Results" }, { depth: 12 }));
  }, 20_000, 400);
  await waitListSettled();
}

async function defaultScope(): Promise<string> {
  const rows = await readFolderRows();
  if (rows.some((r) => r.section === "All Accounts")) return "All Accounts/Inbox";
  const inbox = rows.find((r) => r.section !== "Favorites" && r.depth === 1 && r.name === "Inbox");
  if (!inbox) throw new OutlookError("No Inbox found to search from; pass a folder.");
  return inbox.path;
}

/**
 * Search with Outlook's own search box. The scope is the mailbox of `folder` (Outlook searches the whole
 * account of the selected folder; an "All Accounts/…" folder searches everything).
 */
export async function searchMessages(opts: { query: string; folder?: string; limit?: number }) {
  const scope = opts.folder ?? (await defaultScope());
  const folder = await openFolder(scope);
  await runSearch(opts.query);
  const rows = await loadRows(opts.limit ?? 25);
  const title = (await ax.get("main")).title ?? "";
  return {
    query: opts.query,
    scope: title.replace(/^Searching\s*/, "").replace(/^"|"$/g, "") || folder.path,
    messages: toSummaries(rows, folder.path, opts.query).slice(0, opts.limit ?? 25),
  };
}

interface Located {
  row: RowInfo;
  data: RefData;
  matches: number;
}

/** Navigate to the list a ref came from and return the matching row. */
async function locate(ref: string): Promise<Located> {
  const data = decodeRef(ref);
  await assertNoInlineDraft();
  // Reuse the results already on screen when they are for the same search.
  const title = (await ax.get("main")).title ?? "";
  const field = data.q ? await ax.findOne("main", { role: "AXTextField", subrole: "AXSearchField", id: "Search Bar" }, { depth: 3 }) : null;
  const sameSearch = !!data.q && title.startsWith("Searching") && field?.value === data.q;
  if (!sameSearch) {
    await openFolder(data.f);
    if (data.q) await runSearch(data.q);
  }
  let want = 25;
  for (let i = 0; i < 6; i++) {
    const rows = await loadRows(want);
    const matches = rows.filter((r) => sameRow(r, data));
    const row = matches[data.k];
    if (row) return { row, data, matches: matches.length };
    if (rows.length < want) break; // everything loaded
    want *= 3;
  }
  throw new OutlookError(`Message "${data.s}" from ${data.n} (${data.d}) is no longer in "${data.f}". List the folder again to get a fresh ref.`);
}

/**
 * Conversation rows are not safe targets: selecting one makes Outlook open (and select) its newest item,
 * which may be a draft or a message in another folder. Expand it instead and ask for a specific message.
 */
async function guardConversation(row: RowInfo) {
  if (!row.conversationCount || row.conversationCount < 2) return;
  await ax.set(row.h, "AXSelected", true).catch(() => {});
  await sleep(800);
  throw new OutlookError(
    `This ref is a conversation of ${row.conversationCount} messages. Outlook has expanded it: list the folder again and use the ref of the specific message (conversation members from other folders show a "folder" field).`,
  );
}

async function headerInfo(): Promise<{ subject: string | null; from: string | null; draftSubject: string | null }> {
  const header = await ax.findOne("main", { id: "HeaderMainContainer" }, { depth: 7, subtree: 3 });
  let subject: string | null = null;
  let from: string | null = null;
  if (header) {
    subject = header.children?.find((c) => c.role === "AXStaticText")?.value ?? null;
    const walk = (n: AxNode): void => {
      if (n.id === "messageHeaderFromContent") from = n.value ?? null;
      n.children?.forEach(walk);
    };
    walk(header);
  }
  // Drafts open as an editable form instead of a header.
  const draft = header ? null : await ax.findOne("main", { id: "subjectTextField" }, { depth: 9 });
  return { subject, from, draftSubject: draft?.value ?? null };
}

/** Select a row and wait until the reading pane shows that exact message. */
async function selectRow(row: RowInfo) {
  await ax.set(row.h, "AXSelected", true);
  const norm = (x: string) => x.replace(/^((re|fw|fwd|aw|tr|wg)\s*:\s*)+/i, "").trim();
  const matchesSubject = (s: string | null) =>
    s !== null && (s === row.subject || norm(s) === norm(row.subject) || (row.subject === "" && /^[(<]?no subject[)>]?$/i.test(s)));
  const ok = await until(async () => {
    const sel = await ax.get(row.h).catch(() => null);
    if (!sel?.selected) return false;
    const h = await headerInfo();
    if (h.draftSubject !== null) return matchesSubject(h.draftSubject);
    if (!matchesSubject(h.subject)) return false;
    // Conversation headers show the thread; single messages must also match the sender.
    if (row.conversationCount || !h.from) return true;
    return h.from.includes(row.sender.split(",")[0]!.trim());
  }, 6000, 150);
  if (!ok) throw new OutlookError(`Outlook did not show message "${row.subject}" in the reading pane.`);
}

// ---------------------------------------------------------------------------------------------
// Reading

export interface MessageView {
  ref: string;
  subject: string;
  from: string | null;
  recipients: string | null;
  sent: string | null;
  body: string;
  links: { text: string; url: string }[];
  attachments: string[];
  conversationMessages: number;
  unreadRestored?: boolean;
}

function cleanHeader(v: string | undefined, prefix: RegExp) {
  return v ? v.replace(prefix, "").replace(/,\s*Presence \w+/g, "").trim() : null;
}

async function readReadingPane(): Promise<Omit<MessageView, "ref">> {
  const header = await ax.findOne("main", { id: "HeaderMainContainer" }, { depth: 7 });
  if (!header) {
    // A draft opened for editing.
    const subj = await need("main", { id: "subjectTextField" }, "the reading pane", { depth: 9 });
    const to = await ax.findOne("main", { id: "toTextField" }, { depth: 10 });
    const area = await bodyEditor("main", 2000).catch(() => null);
    return {
      subject: subj.value ?? "",
      from: null,
      recipients: to?.value?.replace(/￼/g, "").trim() || null,
      sent: null,
      body: area ? (await ax.text(area.h)).text : "",
      links: [],
      attachments: [],
      conversationMessages: 1,
    };
  }
  const tree = await ax.tree(header.h, 3);
  const flat: AxNode[] = [];
  const walk = (n: AxNode) => {
    flat.push(n);
    n.children?.forEach(walk);
  };
  walk(tree);
  const subject = tree.children?.find((c) => c.role === "AXStaticText")?.value ?? "";
  const from = flat.find((n) => n.id === "messageHeaderFromContent")?.value;
  const rcpt = flat.find((n) => n.id === "messageHeaderRecipientsContent")?.value;
  const sent = flat.find((n) => n.id === "messageHeaderTimeStampContent")?.value ?? null;
  const panes = await ax.find("main", { role: "AXWebArea", desc: "Reading Pane" }, { depth: 12, limit: 50 });
  const parts: string[] = [];
  const links: { text: string; url: string }[] = [];
  for (const p of panes) {
    const t = await ax.text(p.h);
    parts.push(t.text.trim());
    links.push(...t.links);
  }
  const attachmentNodes = await ax.find("main", { idContains: "attachment" }, { depth: 12, limit: 50 });
  const attachments = new Set<string>();
  for (const a of attachmentNodes) {
    const kids = await ax.find(a.h, { role: "AXButton" }, { depth: 3, limit: 50 });
    for (const k of kids) {
      const label = k.desc ?? k.title ?? "";
      if (label && !/^(More|Download|Save|Preview|Show)/i.test(label)) attachments.add(label);
    }
  }
  return {
    subject,
    from: cleanHeader(from, /^From:\s*/),
    recipients: cleanHeader(rcpt, /^Recipients:\s*/),
    sent,
    body: parts.filter(Boolean).join("\n\n— — —\n\n"),
    links,
    attachments: [...attachments],
    conversationMessages: panes.length,
  };
}

async function rowUnread(row: RowInfo): Promise<boolean | null> {
  const t = await ax.tree(row.h, 1).catch(() => null);
  const desc = t?.children?.find((c) => c.role === "AXCell")?.desc;
  return desc ? (parseRowDesc(desc)?.unread ?? null) : null;
}

/** Outlook marks a previewed message read (sometimes after a delay); put the unread flag back. */
async function restoreUnread(row: RowInfo): Promise<boolean> {
  const becameRead = await until(async () => ((await rowUnread(row)) === false ? true : null), 2500, 250);
  if (!becameRead) return (await rowUnread(row)) !== false;
  await withFocus(async () => {
    await ax.menu(["Message", "Toggle Unread"]);
  });
  return !!(await until(async () => ((await rowUnread(row)) === true ? true : null), 3000, 250));
}

export async function readMessage(ref: string, opts: { markAsRead?: boolean } = {}): Promise<MessageView> {
  const { row } = await locate(ref);
  await guardConversation(row);
  await selectRow(row);
  await sleep(300);
  const view = await readReadingPane();
  const result: MessageView = { ref, ...view };
  if (row.unread && !opts.markAsRead) result.unreadRestored = await restoreUnread(row).catch(() => false);
  return result;
}

// ---------------------------------------------------------------------------------------------
// Save As .eml (full MIME source, attachments)

const DOWNLOADS = join(homedir(), "Downloads");

async function waitForStableFile(path: string, timeoutMs: number): Promise<boolean> {
  let lastSize = -1;
  const ok = await until(async () => {
    if (!existsSync(path)) return false;
    const size = statSync(path).size;
    const stable = size > 0 && size === lastSize;
    lastSize = size;
    return stable;
  }, timeoutMs, 400);
  return !!ok;
}

/** Close a sheet (dialog) left open by a failed step, pressing its Cancel button. */
async function cancelSheet(match: Record<string, string>) {
  const node = await ax.findOne("main", match, { depth: 3 }).catch(() => null);
  if (!node) return;
  const sheet = node.role === "AXSheet" ? node : await ax.call<AxNode>("parent", { h: node.h }).catch(() => null);
  const cancel = sheet ? await ax.findOne(sheet.h, { role: "AXButton", titleRegex: "^Cancel$" }, { depth: 3 }) : null;
  if (cancel) await ax.press(cancel.h).catch(() => {});
}

/**
 * Export a message as .eml via File > Save As. Outlook is sandboxed and writes where its save panel points;
 * we save into ~/Downloads under a unique name, then move the file to `outPath`.
 */
export async function exportEml(ref: string, outPath?: string): Promise<string> {
  const { row } = await locate(ref);
  await guardConversation(row);
  await selectRow(row);
  const token = `axlook-${process.pid}-${Date.now().toString(36)}`;
  const saved = join(DOWNLOADS, `${token}.eml`);
  await withFocus(async () => {
    try {
      await ax.menu(["File", "Save As..."]);
      const panel = await need("main", { id: "save-panel" }, "the Save As panel", { depth: 2, timeoutMs: 6000 });
      const name = await need(panel.h, { id: "saveAsNameTextField" }, "the file name field", { depth: 4 });
      const where = await ax.findOne(panel.h, { id: "where popup" }, { depth: 4 });
      if (where && where.value !== "Downloads") {
        const sidebarItem = await ax.findOne(panel.h, { role: "AXStaticText", value: "Downloads" }, { depth: 8 });
        if (sidebarItem) {
          const cell = await ax.call<AxNode>("parent", { h: sidebarItem.h });
          const rowEl = await ax.call<AxNode>("parent", { h: cell.h });
          await ax.set(rowEl.h, "AXSelected", true);
          await until(async () => (await ax.get(where.h)).value === "Downloads", 3000);
        }
        const w2 = await ax.get(where.h);
        if (w2.value !== "Downloads") throw new OutlookError(`Outlook's save panel points at "${w2.value}", and switching it to Downloads failed.`);
      }
      await ax.set(name.h, "AXValue", token);
      if (!(await until(async () => (await ax.get(name.h)).value === token, 2000))) throw new OutlookError("Could not set the file name in the save panel.");
      const okBtn = await need(panel.h, { id: "OKButton" }, "the Save button", { depth: 4 });
      await ax.press(okBtn.h);
      const gone = await ax.waitFor("main", { id: "save-panel" }, { gone: true, timeoutMs: 5000, depth: 2 });
      if (!gone.ok) throw new OutlookError("The save panel did not close.");
    } catch (e) {
      await cancelSheet({ id: "save-panel" });
      throw e;
    }
  });
  if (row.unread) await restoreUnread(row).catch(() => {});
  // Outlook writes the file asynchronously (it may need to download the MIME first).
  if (!(await waitForStableFile(saved, 60_000))) {
    throw new OutlookError("Outlook did not write the .eml file (timed out after 60 s).");
  }
  const target = outPath ?? join(tmpdir(), "axlook", `${token}.eml`);
  mkdirSync(join(target, ".."), { recursive: true });
  if (existsSync(target)) throw new OutlookError(`Refusing to overwrite existing file ${target}.`);
  renameSync(saved, target);
  return target;
}

// ---------------------------------------------------------------------------------------------
// Actions on a message

async function countMatches(data: RefData): Promise<number> {
  await waitListSettled();
  return (await readRows(await messageTable())).filter((r) => sameRow(r, data)).length;
}

/** Move a message to another folder via Message > Move > Choose Folder… */
export async function moveMessage(ref: string, destination: string) {
  const dest = await findFolderRow(destination);
  if (dest.section === "All Accounts" || dest.section === "Favorites") throw new OutlookError("Pick a destination inside a specific account (not 'All Accounts' or 'Favorites').");
  const { row, data, matches } = await locate(ref);
  await guardConversation(row);
  await selectRow(row);
  await withFocus(async () => {
    try {
      await ax.menu(["Message", "Move", "Choose Folder..."]);
      const box = await need("main", { role: "AXTextField", desc: "search box" }, "the Move dialog", { depth: 3, timeoutMs: 5000 });
      const sheet = await ax.call<AxNode>("parent", { h: box.h });
      await ax.set(box.h, "AXValue", dest.name);
      // Items are labelled "Name, account[/parent/…]".
      const parents = dest.path.split("/").slice(1, -1);
      const label = `${dest.name}, ${[dest.account, ...parents].join("/")}`;
      const item = await until(async () => {
        const items = await ax.find(sheet.h, { role: "AXGroup", descContains: dest.name }, { depth: 3, limit: 100 });
        return items.find((i) => (i.desc ?? "").toLowerCase() === label.toLowerCase()) ?? null;
      }, 4000);
      if (!item) throw new OutlookError(`The Move dialog does not offer "${label}".`);
      await ax.click(item.h);
      await sleep(300);
      const moveBtn = await need(sheet.h, { role: "AXButton", title: "Move" }, "the Move button", { depth: 1 });
      if ((await ax.get(moveBtn.h)).disabled) throw new OutlookError("The destination could not be selected in the Move dialog.");
      await ax.press(moveBtn.h);
      const closed = await ax.waitFor("main", { role: "AXTextField", desc: "search box" }, { gone: true, timeoutMs: 4000, depth: 3 });
      if (!closed.ok) throw new OutlookError("The Move dialog did not close; the move may not have happened.");
    } catch (e) {
      await cancelSheet({ role: "AXTextField", desc: "search box" });
      throw e;
    }
  });
  const verified = await until(async () => ((await countMatches(data)) === matches - 1 ? true : null), 8000, 500);
  return { moved: true, verified: !!verified, from: data.f, to: dest.path };
}

export async function setReadState(ref: string, read: boolean) {
  const { row } = await locate(ref);
  await guardConversation(row);
  if (row.unread === !read) return { changed: false, unread: row.unread };
  await selectRow(row);
  // Previewing may already mark it read.
  await sleep(800);
  if ((await rowUnread(row)) === !read) return { changed: true, unread: !read };
  await withFocus(async () => {
    await ax.menu(["Message", "Toggle Unread"]);
  });
  const now = await until(async () => ((await rowUnread(row)) === !read ? true : null), 3000, 250);
  if (!now) throw new OutlookError("Outlook did not change the read state.");
  return { changed: true, unread: !read };
}

async function rowFlagged(row: RowInfo): Promise<boolean | null> {
  const t = await ax.tree(row.h, 1).catch(() => null);
  const desc = t?.children?.find((c) => c.role === "AXCell")?.desc;
  return desc ? (parseRowDesc(desc)?.flagged ?? null) : null;
}

export async function setFlag(ref: string, flagged: boolean) {
  const { row } = await locate(ref);
  await guardConversation(row);
  if (row.flagged === flagged) return { changed: false, flagged };
  await selectRow(row);
  await withFocus(async () => {
    // The "flag options" toolbar menu offers "No Date" (flag) and "Clear Flag".
    const opts = await need("main", { role: "AXButton", help: "View Flag for follow up options" }, "the flag options button", { depth: 3 });
    await ax.press(opts.h);
    const title = flagged ? "No Date" : "Clear Flag";
    const item = await ax.waitFor("app", { role: "AXMenuItem", id: "handleRicherFlagMenuAction:", title }, { timeoutMs: 3000, depth: 8 });
    if (!item.ok || !item.node) {
      await ax.key(KEY.ESCAPE).catch(() => {});
      throw new OutlookError(`Outlook's flag menu has no "${title}" item.`);
    }
    await ax.press(item.node.h);
  });
  const ok = await until(async () => ((await rowFlagged(row)) === flagged ? true : null), 3000, 250);
  if (!ok) throw new OutlookError("Outlook did not change the flag.");
  return { changed: true, flagged };
}

export async function archiveMessage(ref: string) {
  const { row, data, matches } = await locate(ref);
  await guardConversation(row);
  await selectRow(row);
  await withFocus(async () => {
    await ax.menu(["Message", "Archive"]);
  });
  const verified = await until(async () => ((await countMatches(data)) === matches - 1 ? true : null), 6000, 500);
  return { archived: true, verified: !!verified, from: data.f };
}

/** Move to Deleted Items (Edit > Delete). Never permanently deletes. */
export async function deleteMessage(ref: string) {
  const { row, data, matches } = await locate(ref);
  await guardConversation(row);
  if (/deleted items|deleted messages|trash|bin/i.test(data.f)) {
    throw new OutlookError("Refusing to delete from a deleted-items folder: Outlook would delete it permanently.");
  }
  await selectRow(row);
  await withFocus(async () => {
    const isDraft = !!(await ax.findOne("main", { id: "subjectTextField" }, { depth: 10 }));
    if (!isDraft) {
      await ax.menu(["Edit", "Delete"]);
      return;
    }
    // A selected draft opens in edit mode: use Discard, which also moves it to Deleted Items.
    const discard = await need("main", { role: "AXButton", descRegex: "^Discard(,|$)" }, "the Discard button", { depth: 3 });
    await ax.press(discard.h);
    const sheet = await ax.waitFor("main", { role: "AXSheet" }, { timeoutMs: 3000, depth: 1 });
    if (sheet.ok && sheet.node) {
      const btn = await need(sheet.node.h, { role: "AXButton", title: "Discard" }, "the Discard confirmation", { depth: 2 });
      await ax.press(btn.h);
    }
  });
  const verified = await until(async () => ((await countMatches(data)) === matches - 1 ? true : null), 6000, 500);
  return { movedToDeletedItems: true, verified: !!verified, from: data.f };
}

export { decodeRef, locate, selectRow };
