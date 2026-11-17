// Drafts: new messages, replies and forwards are saved to Drafts and never sent.
import { KEY, type AxNode } from "../ax";
import { ax, bodyEditor, focusWindow, need, openApp, OutlookError, pasteInto, pressMenu, secondaryWindows, sleep, until, withFocus } from "./session";
import { openMessage } from "./mail";
import { htmlToText } from "../mime";
import { basename } from "node:path";

export interface DraftInput {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  html?: boolean; // treat body as HTML
  from?: string; // account e-mail to send from (must exist in Outlook)
  attachments?: string[]; // absolute paths of local files
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
  // The plain-text flavour, also used to check that the paste arrived: block tags become line breaks.
  const text = html ? htmlToText(body) : body;
  const htmlPayload = html ? body : text.split(/\r?\n/).map((l) => (l ? escapeHtml(l) : "<br>")).map((l) => `<div>${l}</div>`).join("");
  await pasteInto(area, text + "\n", { html: htmlPayload + "<div><br></div>", mode: "start" });
}

/**
 * Add local files by pasting them into the body, as a copy in Finder would: Outlook embeds images where the
 * caret is and attaches other files. Each file must show up as an attachment or an embedded image.
 */
async function attachFiles(win: AxNode, files: string[] | undefined, caretPlaced: boolean) {
  if (!files?.length) return;
  const present = async () => {
    const names = await ax.find(win.h, { role: "AXStaticText", desc: "Attachment name" }, { depth: 14, limit: 500 });
    const images = await ax.find(win.h, { role: "AXImage" }, { depth: 18, limit: 500 });
    return [...names.map((n) => n.value ?? ""), ...images.map((i) => i.desc ?? "")];
  };
  const count = (list: string[], name: string) => list.filter((x) => x === name).length;
  const before = await present();
  const area = await bodyEditor(win.h);
  const focused = await until(async () => {
    if ((await ax.get(area.h, { attrs: ["AXFocused"] })).attrs?.AXFocused === true) return true;
    await ax.set(area.h, "AXFocused", true);
    caretPlaced = false;
    return false;
  }, 4000, 300);
  if (!focused) throw new OutlookError("Could not focus the message body to add the attachments.");
  // Without text of ours, put images at the top (above the signature or the quoted message).
  if (!caretPlaced) await ax.key(KEY.UP, ["cmd"]);
  await sleep(150);
  await ax.pasteFiles(files);
  const wanted = new Map<string, number>();
  for (const f of files) wanted.set(basename(f), (wanted.get(basename(f)) ?? 0) + 1);
  const missing = async () => {
    const now = await present();
    return [...wanted].filter(([name, n]) => count(now, name) < count(before, name) + n).map(([name]) => name);
  };
  if (!(await until(async () => (await missing()).length === 0, 20_000, 500))) {
    throw new OutlookError(`Outlook did not add ${(await missing()).join(", ")} to the draft.`);
  }
}

function escapeHtml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Save (File > Save) and close a separate compose window, keeping the draft. Throws unless both happen. */
async function saveAndClose(win: AxNode) {
  // File > Save applies to the key window, and stays disabled while the main window is key.
  await focusWindow(win);
  await pressMenu(["File", "Save"], { why: "the compose window did not become the active window." });
  const saved = await until(async () => !!(await ax.findOne(win.h, { role: "AXStaticText", valuePrefix: "Draft saved" }, { depth: 4 })), 10_000, 300);
  if (!saved) throw new OutlookError("Outlook did not confirm that the draft was saved.");
  const close = await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 });
  if (close) await ax.press(close.h);
  // A "save changes?" sheet can appear on some versions: keep the draft.
  const sheet = await ax.waitFor(win.h, { role: "AXSheet" }, { timeoutMs: 1200, depth: 1 }).catch(() => null);
  if (sheet?.ok && sheet.node) {
    const keep = await ax.findOne(sheet.node.h, { role: "AXButton", titleRegex: "^(Save|Keep)" }, { depth: 2 });
    if (keep) await ax.press(keep.h);
  }
  if (!(await until(() => windowGone(win), 5000))) throw new OutlookError("The draft was saved but its window did not close.");
}

/** A closed window's element stops answering (titles cannot tell: another window may carry the same one). */
const windowGone = async (win: AxNode) => !(await ax.get(win.h).catch(() => null))?.role;

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

/** Press a compose form's Discard button and confirm. Returns whether the draft went away. */
async function discardIn(root: string, gone: () => Promise<boolean>): Promise<boolean> {
  for (let attempt = 0; attempt < 2 && !(await gone()); attempt++) {
    const discard = await ax.findOne(root, { role: "AXButton", descRegex: "^Discard(,|$)" }, { depth: 4 }).catch(() => null);
    if (!discard) return false;
    await ax.press(discard.h).catch(() => {});
    // "Discard this draft?": Continue Editing / Discard.
    const sheet = await ax.waitFor(root, { role: "AXSheet" }, { timeoutMs: 3000, depth: 1 }).catch(() => null);
    if (sheet?.ok && sheet.node) {
      const btn = await ax.findOne(sheet.node.h, { role: "AXButton", titleRegex: "^Discard" }, { depth: 3 }).catch(() => null);
      if (btn) await ax.press(btn.h).catch(() => {});
    }
    if (await until(gone, 4000, 200)) return true;
  }
  return gone();
}

/**
 * On failure, discard the draft Outlook already autosaved, so that a failed call leaves nothing behind and a
 * retry does not create a duplicate. `inline`: the draft is still open in the main window's reading pane.
 */
async function abandon(win: AxNode | null, inline = false): Promise<string> {
  if (win) {
    if (await discardIn(win.h, () => windowGone(win))) return "";
    const close = await ax.findOne(win.h, { subrole: "AXCloseButton" }, { depth: 1 }).catch(() => null);
    if (close) await ax.press(close.h).catch(() => {});
    return " The unfinished draft could not be discarded and may remain in Drafts.";
  }
  if (inline) {
    const open = async () => !!(await ax.findOne("main", { id: "subjectTextField" }, { depth: 10 }).catch(() => null));
    if (await discardIn("main", async () => !(await open()))) return "";
    return " An unfinished draft is still open in Outlook's reading pane.";
  }
  return "";
}

/** Rethrow `e` with a note about what was left behind. */
function withNote(e: unknown, note: string): never {
  if (!note) throw e;
  throw e instanceof OutlookError ? new OutlookError(e.message + note) : e;
}

export async function createDraft(d: DraftInput) {
  if (!d.to?.length && !d.cc?.length && !d.bcc?.length && !d.subject && !d.body) {
    throw new OutlookError("A draft needs at least a recipient, a subject or a body.");
  }
  // With attachments, paste the body too: images then land right after it.
  const bodyViaUrl = !!d.body && !d.html && d.body.length <= MAILTO_BODY_LIMIT && !d.attachments?.length;
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
      await attachFiles(win, d.attachments, !!d.body && !bodyViaUrl);
      const fromValue = (await ax.findOne(win.h, { id: "accountPicker" }, { depth: 6 }))?.value ?? null;
      await saveAndClose(win);
      return {
        saved: true,
        subject: d.subject ?? "",
        from: fromValue,
        to: d.to ?? [],
        cc: d.cc ?? [],
        bcc: d.bcc ?? [],
        ...(d.attachments?.length ? { attachments: d.attachments.map((f) => basename(f)) } : {}),
      };
    } catch (e) {
      withNote(e, await abandon(win));
    }
  });
}

/** Create a reply / reply-all / forward draft for a listed message. Nothing is sent. */
export async function respondDraft(
  ref: string,
  kind: "reply" | "replyAll" | "forward",
  body: string,
  opts: { to?: string[]; html?: boolean; attachments?: string[] } = {},
) {
  await openMessage(ref, { received: true });
  const menuItem = { reply: "Reply", replyAll: "Reply All", forward: "Forward" }[kind];
  return withFocus(async () => {
    const before = await secondaryWindows();
    await pressMenu(["Message", menuItem], {
      timeoutMs: 15_000,
      why: "the message may still be downloading, or Outlook does not allow this action on it. Retry in a moment.",
    });
    let win: AxNode | null = null;
    let inline = false;
    try {
      // New Outlook opens replies inline in the reading pane; pop them out into their own window.
      const pop = await ax.waitFor("main", { id: "fullMessageComposeButton" }, { timeoutMs: 6000, depth: 8 });
      if (pop.ok && pop.node) {
        inline = true;
        await ax.press(pop.node.h);
      }
      win = await waitNewWindow(before);
      inline = false;
      await focusWindow(win);
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
      await attachFiles(win, opts.attachments, !!body);
      const subject = (await ax.findOne(win.h, { id: "subjectTextField" }, { depth: 6 }))?.value ?? "";
      // Outlook answers from the account that received the original message.
      const from = (await ax.findOne(win.h, { id: "accountPicker" }, { depth: 6 }))?.value ?? null;
      await saveAndClose(win);
      return { saved: true, kind, subject, from, ...(opts.attachments?.length ? { attachments: opts.attachments.map((f) => basename(f)) } : {}) };
    } catch (e) {
      withNote(e, await abandon(win, inline));
    }
  });
}
