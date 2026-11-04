// Drafts: new messages, replies and forwards are saved to Drafts and never sent.
import { KEY, type AxNode } from "../ax";
import { ax, bodyEditor, need, openApp, OutlookError, pasteInto, secondaryWindows, sleep, until, withFocus } from "./session";
import { locate, selectRow } from "./mail";

export interface DraftInput {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  html?: boolean; // treat body as HTML
  from?: string; // account e-mail to send from (must exist in Outlook)
}

const MAILTO_BODY_LIMIT = 3000;

function mailtoUrl(d: DraftInput, includeBody: boolean) {
  const enc = (s: string) => encodeURIComponent(s);
  const params: string[] = [];
  if (d.cc?.length) params.push(`cc=${d.cc.map(enc).join(",")}`);
  if (d.bcc?.length) params.push(`bcc=${d.bcc.map(enc).join(",")}`);
  if (d.subject) params.push(`subject=${enc(d.subject)}`);
  if (includeBody && d.body) params.push(`body=${enc(d.body.replace(/\r?\n/g, "\r\n"))}`);
  return `mailto:${(d.to ?? []).map(enc).join(",")}${params.length ? "?" + params.join("&") : ""}`;
}

/** Find the compose window or inline compose area: the element that contains #subjectTextField. */
async function composeRoot(win: AxNode) {
  return need(win.h, { id: "subjectTextField" }, "the compose form", { depth: 6, timeoutMs: 8000 });
}

async function chooseFromAccount(win: AxNode, from: string) {
  const picker = await need(win.h, { id: "accountPicker" }, "the From account picker", { depth: 6 });
  if ((picker.value ?? "").toLowerCase().includes(from.toLowerCase())) return;
  await ax.press(picker.h);
  const item = await until(async () => {
    const items = await ax.find("app", { role: "AXMenuItem" }, { depth: 6, limit: 200 });
    return items.find((i) => (i.title ?? "").toLowerCase().includes(from.toLowerCase())) ?? null;
  }, 3000);
  if (!item) {
    await ax.key(KEY.ESCAPE);
    throw new OutlookError(`No sending account matching "${from}" in the From picker.`);
  }
  await ax.press(item.h);
  await until(async () => ((await ax.get(picker.h)).value ?? "").toLowerCase().includes(from.toLowerCase()), 3000);
}

/** Put the caret at the top of the body editor and paste text/HTML there (above signature / quoted text). */
async function pasteIntoBody(win: AxNode, body: string, html: boolean) {
  const area = await bodyEditor(win.h);
  const text = html ? body.replace(/<[^>]+>/g, "") : body;
  const htmlPayload = html ? body : text.split(/\r?\n/).map((l) => (l ? escapeHtml(l) : "<br>")).map((l) => `<div>${l}</div>`).join("");
  await pasteInto(area, text + "\n", { html: htmlPayload + "<div><br></div>", mode: "start" });
}

function escapeHtml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Save (File > Save) and close a separate compose window, keeping the draft. Throws unless both happen. */
async function saveAndClose(win: AxNode) {
  await ax.menu(["File", "Save"]);
  const saved = await until(async () => !!(await ax.findOne(win.h, { role: "AXStaticText", valuePrefix: "Draft saved" }, { depth: 4 })), 10_000, 300);
  if (!saved) throw new OutlookError("Outlook did not confirm that the draft was saved; the compose window was left open.");
  const close = await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 });
  if (close) await ax.press(close.h);
  // A "save changes?" sheet can appear on some versions: keep the draft.
  const sheet = await ax.waitFor(win.h, { role: "AXSheet" }, { timeoutMs: 1200, depth: 1 }).catch(() => null);
  if (sheet?.ok && sheet.node) {
    const keep = await ax.findOne(sheet.node.h, { role: "AXButton", titleRegex: "^(Save|Keep)" }, { depth: 2 });
    if (keep) await ax.press(keep.h);
  }
  const closed = await until(async () => !(await ax.call<AxNode[]>("windows")).some((w) => w.h !== undefined && w.title === win.title && w.id !== "OutlookWindow"), 5000);
  if (!closed) throw new OutlookError("The draft was saved but its window did not close.");
}

/** The compose window that appeared after an action: the focused non-main window not present before. */
async function waitNewWindow(before: AxNode[]): Promise<AxNode> {
  const known = new Set(before.map((w) => `${w.title}|${JSON.stringify(w.frame)}`));
  const w = await until(async () => {
    const focused = await ax.get("focusedWindow", { frame: true }).catch(() => null);
    if (focused && focused.id !== "OutlookWindow" && focused.role === "AXWindow" && !known.has(`${focused.title}|${JSON.stringify(focused.frame)}`)) {
      return focused;
    }
    return (await secondaryWindows()).find((x) => !known.has(`${x.title}|${JSON.stringify(x.frame)}`)) ?? null;
  }, 10_000, 250);
  if (!w) throw new OutlookError("Outlook did not open a compose window.");
  return w;
}

/** On failure, close the compose window (Outlook keeps the autosaved draft). */
async function abandon(win: AxNode | null) {
  if (!win) return;
  const close = await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 }).catch(() => null);
  if (close) await ax.press(close.h).catch(() => {});
}

export async function createDraft(d: DraftInput) {
  if (!d.to?.length && !d.cc?.length && !d.bcc?.length && !d.subject && !d.body) {
    throw new OutlookError("A draft needs at least a recipient, a subject or a body.");
  }
  const bodyViaUrl = !!d.body && !d.html && d.body.length <= MAILTO_BODY_LIMIT;
  return withFocus(async () => {
    const before = await secondaryWindows();
    await openApp(["-a", "Microsoft Outlook", mailtoUrl(d, bodyViaUrl)]);
    let win: AxNode | null = null;
    try {
      win = await waitNewWindow(before);
      const subj = await composeRoot(win);
      if (d.subject && subj.value !== d.subject) throw new OutlookError("The compose window did not receive the subject.");
      if (d.from) await chooseFromAccount(win, d.from);
      if (d.body && !bodyViaUrl) await pasteIntoBody(win, d.body, !!d.html);
      const fromValue = (await ax.findOne(win.h, { id: "accountPicker" }, { depth: 6 }))?.value ?? null;
      await saveAndClose(win);
      return { saved: true, subject: d.subject ?? "", from: fromValue, to: d.to ?? [], cc: d.cc ?? [], bcc: d.bcc ?? [] };
    } catch (e) {
      await abandon(win);
      throw e;
    }
  });
}

/** Create a reply / reply-all / forward draft for a listed message. Nothing is sent. */
export async function respondDraft(ref: string, kind: "reply" | "replyAll" | "forward", body: string, opts: { to?: string[]; html?: boolean } = {}) {
  const { row } = await locate(ref);
  await selectRow(row);
  const menuItem = { reply: "Reply", replyAll: "Reply All", forward: "Forward" }[kind];
  return withFocus(async () => {
    const before = await secondaryWindows();
    await ax.menu(["Message", menuItem]);
    // New Outlook opens replies inline in the reading pane; pop them out into their own window.
    const inline = await ax.waitFor("main", { id: "fullMessageComposeButton" }, { timeoutMs: 6000, depth: 8 });
    if (inline.ok && inline.node) await ax.press(inline.node.h);
    let win: AxNode | null = null;
    try {
    win = await waitNewWindow(before);
    await composeRoot(win);
    if (opts.to?.length) {
      const to = await need(win.h, { id: "toTextField" }, "the To field", { depth: 6 });
      await ax.set(to.h, "AXFocused", true);
      await sleep(150);
      for (const addr of opts.to) {
        await ax.type(addr);
        await ax.key(KEY.RETURN);
        await sleep(300);
      }
      // Each resolved recipient becomes a token (U+FFFC) in the field.
      const w = win;
      const tokens = async () => ((await need(w.h, { id: "toTextField" }, "the To field", { depth: 6 })).value ?? "").split("\uFFFC").length - 1;
      if (!(await until(async () => (await tokens()) >= opts.to!.length, 3000, 250))) {
        throw new OutlookError("The forward recipients were not accepted by Outlook.");
      }
    }
    if (body) await pasteIntoBody(win, body, !!opts.html);
    const subject = (await ax.findOne(win.h, { id: "subjectTextField" }, { depth: 6 }))?.value ?? "";
    await saveAndClose(win);
    return { saved: true, kind, subject };
    } catch (e) {
      await abandon(win);
      throw e;
    }
  });
}
