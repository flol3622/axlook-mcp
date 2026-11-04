// Live end-to-end check of every tool through the MCP protocol. Touches only test targets:
// a test folder, test drafts and a test event (all cleaned up). Usage: bun scripts/e2e.ts [step…]
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { mkdtempSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

// Configure for your own mailbox (all required):
//   E2E_FOLDER        folder holding at least one disposable test message, e.g. "you@example.com/axlook-tests"
//   E2E_MOVE_TO       another folder of the same account to move it to and back
//   E2E_SELF          your own address (recipient of the forward draft)
//   E2E_ATTACH_SCOPE  folder containing a message with attachments, e.g. "you@example.com/Inbox"
//   E2E_CALENDAR      calendar for the temporary test event
// Optional: E2E_SENDER (preferred test message sender), E2E_DAY (YYYY-MM-DD for the test event),
//           E2E_DRAFT_FOLDERS (comma-separated Drafts folders to clean up; default: the Drafts of E2E_FOLDER's account)
const need = (k: string) => process.env[k] ?? (console.error(`Set ${k} (see the header of scripts/e2e.ts).`), process.exit(2));
const TEST_FOLDER = need("E2E_FOLDER");
const MOVE_TO = need("E2E_MOVE_TO");
const TEST_SENDER = process.env.E2E_SENDER ?? "";
const SELF = need("E2E_SELF");
const ATTACH_SCOPE = need("E2E_ATTACH_SCOPE");

const client = new Client({ name: "e2e", version: "0" });
await client.connect(new StdioClientTransport({ command: "bun", args: ["run", new URL("../src/server.ts", import.meta.url).pathname] }));

const results: { step: string; ok: boolean; ms: number; note: string }[] = [];
async function call(name: string, args: Record<string, unknown> = {}) {
  const res: any = await client.callTool({ name, arguments: args }, undefined, { timeout: 600_000 });
  const text = res.content?.[0]?.text ?? "";
  if (res.isError) throw new Error(`${name}: ${text}`);
  return JSON.parse(text);
}
async function step(name: string, fn: () => Promise<string>) {
  const only = process.argv.slice(2);
  if (only.length && !only.some((o) => name.startsWith(o))) return;
  const t0 = Date.now();
  try {
    const note = await fn();
    results.push({ step: name, ok: true, ms: Date.now() - t0, note });
    console.log(`✅ ${name} (${Date.now() - t0} ms) ${note}`);
  } catch (e: any) {
    results.push({ step: name, ok: false, ms: Date.now() - t0, note: e.message });
    console.log(`❌ ${name} (${Date.now() - t0} ms) ${e.message}`);
  }
}
const assert = (c: unknown, m: string) => {
  if (!c) throw new Error(m);
};

let msg: any;
const single = async () => {
  let list = await call("outlook_list_messages", { folder: TEST_FOLDER, limit: 10 });
  if (!list.messages.some((x: any) => !x.conversationCount && !x.folder)) {
    // Only conversation rows: a first call on one expands it (and returns guidance), then list again.
    const conv = list.messages.find((x: any) => x.conversationCount);
    if (conv) await client.callTool({ name: "outlook_read_message", arguments: { ref: conv.ref } }, undefined, { timeout: 600_000 });
    list = await call("outlook_list_messages", { folder: TEST_FOLDER, limit: 10 });
  }
  // Prefer the dedicated test message; Outlook may hide it inside an expanded conversation.
  const m = list.messages.find((x: any) => x.sender === TEST_SENDER && !x.conversationCount && !x.folder) ??
    list.messages.find((x: any) => !x.conversationCount && !x.folder);
  assert(m, "no single (non-conversation) message in test folder");
  return m;
};

await step("mail.folders", async () => {
  const f = await call("outlook_list_folders", {});
  assert(f.some((x: any) => x.path === TEST_FOLDER), "test folder missing");
  return `${f.length} folders`;
});
await step("mail.read", async () => {
  msg = await single();
  const v = await call("outlook_read_message", { ref: msg.ref });
  assert(v.subject === msg.subject && v.body.length > 50, "subject/body mismatch");
  return `${v.from} | body ${v.body.length} chars | links ${v.links.length}`;
});
await step("mail.source", async () => {
  msg ??= await single();
  const s = await call("outlook_get_message_source", { ref: msg.ref });
  assert(s.messageId && s.subject === msg.subject, "bad source");
  return `${s.messageId} attachments=${s.attachments.length}`;
});
await step("mail.unread", async () => {
  msg = await single();
  const a = await call("outlook_mark_read", { ref: msg.ref, read: false });
  assert(a.unread === true, "not unread");
  msg = await single();
  assert(msg.unread, "list does not show unread");
  const b = await call("outlook_mark_read", { ref: msg.ref, read: true });
  assert(b.unread === false, "not read again");
  return "unread → read";
});
await step("mail.flag", async () => {
  msg = await single();
  const a = await call("outlook_set_flag", { ref: msg.ref, flagged: true });
  assert(a.flagged, "not flagged");
  msg = await single();
  assert(msg.flagged, "list does not show flag");
  const b = await call("outlook_set_flag", { ref: msg.ref, flagged: false });
  assert(!b.flagged, "still flagged");
  return "flag on → off";
});
await step("mail.move", async () => {
  msg = await single();
  const a = await call("outlook_move_message", { ref: msg.ref, destination: MOVE_TO });
  assert(a.verified, "move not verified");
  let there = await call("outlook_list_messages", { folder: MOVE_TO, limit: 30 });
  let moved = there.messages.find((x: any) => x.subject === msg.subject && x.sender === msg.sender && !x.folder && !x.conversationCount);
  if (!moved) {
    // It may have joined a conversation in the destination: a first move attempt expands it.
    const conv = there.messages.find((x: any) => x.conversationCount && x.subject.replace(/^(Re|FW): /, "") === msg.subject.replace(/^(Re|FW): /, ""));
    if (conv) await client.callTool({ name: "outlook_move_message", arguments: { ref: conv.ref, destination: TEST_FOLDER } }, undefined, { timeout: 600_000 });
    there = await call("outlook_list_messages", { folder: MOVE_TO, limit: 30 });
    moved = there.messages.find((x: any) => x.subject === msg.subject && x.sender === msg.sender && !x.folder && !x.conversationCount);
  }
  assert(moved, "not found in destination");
  const b = await call("outlook_move_message", { ref: moved.ref, destination: TEST_FOLDER });
  assert(b.verified, "move back not verified");
  return `${TEST_FOLDER} → ${MOVE_TO} → back`;
});
await step("mail.attachments", async () => {
  const l = await call("outlook_list_messages", { folder: ATTACH_SCOPE, limit: 80 });
  const m = l.messages.find((x: any) => x.hasAttachments && !x.conversationCount);
  assert(m, "no single message with attachments found");
  const dir = mkdtempSync(join(tmpdir(), "outlook-e2e-"));
  const r = await call("outlook_save_attachments", { ref: m.ref, directory: dir, include_inline: true });
  const files = readdirSync(dir);
  assert(files.length === r.saved.length && files.length > 0, `saved ${files.length} files of ${r.attachmentsInMessage}`);
  const after = (await call("outlook_list_messages", { folder: ATTACH_SCOPE, limit: 80 })).messages.find((x: any) => x.subject === m.subject && x.sender === m.sender);
  assert(after && after.unread === m.unread, "unread state changed");
  return `${m.subject.slice(0, 40)}: ${files.length} file(s) → ${dir}`;
});
await step("draft.new", async () => {
  const r = await call("outlook_create_draft", { to: [SELF], subject: "MCP e2e draft (safe to delete)", body: "Line one\nLine two ✓ — é" });
  assert(r.saved, "not saved");
  return `from ${r.from}`;
});
await step("draft.reply", async () => {
  msg = await single();
  const r = await call("outlook_create_reply_draft", { ref: msg.ref, body: "MCP e2e reply text (safe to delete)" });
  assert(r.saved, "not saved");
  return r.subject;
});
await step("draft.forward", async () => {
  msg = await single();
  const r = await call("outlook_create_forward_draft", { ref: msg.ref, to: [SELF], body: "MCP e2e forward text (safe to delete)" });
  assert(r.saved, "not saved");
  return r.subject;
});

let ev: any;
const day = process.env.E2E_DAY ?? "2026-11-21";
await step("cal.list", async () => {
  const c = await call("outlook_list_calendars", {});
  assert(c.length > 0, "no calendars");
  return c.map((x: any) => `${x.path}${x.shown ? "" : " (hidden)"}`).slice(0, 6).join("; ");
});
await step("cal.events", async () => {
  const today = new Date();
  const ymd = (d: Date) => d.toISOString().slice(0, 10);
  const r = await call("outlook_list_events", { start: ymd(today), end: ymd(new Date(today.getTime() + 6 * 86400e3)) });
  return `${r.events.length} events: ${r.events.slice(0, 4).map((e: any) => `${e.date} ${e.start ?? "all-day"} ${e.subject}`).join(" | ")}`;
});
await step("cal.create", async () => {
  const r = await call("outlook_create_event", { subject: "MCP e2e event (safe to delete)", calendar: need("E2E_CALENDAR"), start: `${day}T10:00`, end: `${day}T11:15`, location: "Nowhere", body: "Created by the e2e test." });
  assert(r.created, "not created");
  // Outlook syncs new event details into its previews after a few seconds.
  for (let i = 0; i < 6; i++) {
    ev = (await call("outlook_list_events", { start: day })).events.find((e: any) => e.subject === "MCP e2e event (safe to delete)");
    if (ev) break;
    await Bun.sleep(5000);
  }
  assert(ev && ev.start === "10:00" && ev.end === "11:15", `event not found with right times: ${JSON.stringify(ev)}`);
  assert(ev.location === "Nowhere", `location is ${ev.location}`);
  return `${r.calendar} ${ev.date} ${ev.start}-${ev.end} @${ev.location}`;
});
const findTestEvent = async () => (await call("outlook_list_events", { start: day })).events.find((e: any) => e.subject.startsWith("MCP e2e event"));
await step("cal.get", async () => {
  ev ??= await findTestEvent();
  assert(ev, "no event");
  const g = await call("outlook_get_event", { ref: ev.ref, full: true });
  return `${g.calendar} | ${g.when} | body: ${(g.body ?? "").slice(0, 40)}`;
});
await step("cal.update", async () => {
  ev ??= await findTestEvent();
  assert(ev, "no event");
  const u = await call("outlook_update_event", { ref: ev.ref, subject: "MCP e2e event EDITED (safe to delete)", start: `${day}T14:00`, end: `${day}T14:30` });
  for (let i = 0; i < 6; i++) {
    ev = (await call("outlook_list_events", { start: day })).events.find((e: any) => e.subject === "MCP e2e event EDITED (safe to delete)");
    if (ev?.start === "14:00") break;
    await Bun.sleep(5000);
  }
  assert(ev && ev.start === "14:00" && ev.end === "14:30", `edited event wrong: ${JSON.stringify(ev)}`);
  return `${u.updated} → ${ev.start}-${ev.end}`;
});
await step("cal.delete", async () => {
  ev ??= await findTestEvent();
  assert(ev, "no event");
  const d = await call("outlook_delete_event", { ref: ev.ref });
  const l = await call("outlook_list_events", { start: day });
  assert(!l.events.some((e: any) => e.subject.startsWith("MCP e2e event")), "event still there");
  return JSON.stringify(d);
});

await step("cleanup.drafts", async () => {
  // Remove the drafts this run created (they go to Deleted Items).
  const isTest = (m: any) => /^\d{1,2}:\d{2}$/.test(m.date) && (/^MCP (e2e|draft)/.test(m.subject) || /^MCP e2e/.test(m.preview));
  let n = 0;
  for (const folder of [...new Set([`${TEST_FOLDER.split("/")[0]}/Drafts`, ...(process.env.E2E_DRAFT_FOLDERS ?? "").split(",").filter(Boolean)])]) {
    for (let i = 0; i < 10; i++) {
      const t = (await call("outlook_list_messages", { folder, limit: 30 })).messages.find(isTest);
      if (!t) break;
      await call("outlook_delete_message", { ref: t.ref });
      n++;
    }
  }
  return `${n} test drafts moved to Deleted Items`;
});

await client.close();
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} passed`);
process.exit(failed.length ? 1 : 0);
