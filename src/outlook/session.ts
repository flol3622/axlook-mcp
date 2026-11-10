// Shared Outlook UI session: one Accessibility helper, one queue (the UI is a single shared resource),
// focus management and module (Mail / Calendar) switching.
import { execFile } from "node:child_process";
import { AxClient, KEY, type AxNode, type Match } from "../ax";

export class OutlookError extends Error {}

export const ax = new AxClient();

let queue: Promise<unknown> = Promise.resolve();

/** Run UI work strictly one call at a time. Handles from previous calls are released first. */
export function exclusive<T>(fn: () => Promise<T>): Promise<T> {
  const run = queue.then(async () => {
    await ax.call("release").catch(() => {});
    return fn();
  });
  queue = run.catch(() => {});
  return run;
}

export const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Run macOS `open` with the given arguments. */
export function openApp(args: string[]): Promise<void> {
  return new Promise((resolve, reject) => execFile("/usr/bin/open", args, (err) => (err ? reject(err) : resolve())));
}

/** Poll `fn` until it returns a truthy value or the timeout elapses. */
export async function until<T>(fn: () => Promise<T | null | undefined | false>, timeoutMs = 5000, everyMs = 120): Promise<T | null> {
  const deadline = Date.now() + timeoutMs;
  while (true) {
    const v = await fn().catch(() => null);
    if (v) return v as T;
    if (Date.now() >= deadline) return null;
    await sleep(everyMs);
  }
}

export interface Status {
  trusted: boolean;
  outlookRunning: boolean;
  version: string | null;
  pid: number;
}

export async function status(): Promise<Status> {
  return ax.call<Status>("ping");
}

export async function windows(): Promise<AxNode[]> {
  return ax.call<AxNode[]>("windows");
}

/** The main Outlook window (AXIdentifier "OutlookWindow"). Reopens one if Outlook has none. */
export async function mainWindow(): Promise<AxNode> {
  const st = await status();
  if (!st.trusted) {
    throw new OutlookError(
      "Accessibility permission missing. Grant it in System Settings > Privacy & Security > Accessibility to the app that launches this MCP server (e.g. Claude, Terminal, iTerm, VS Code), then restart it.",
    );
  }
  if (!st.outlookRunning) {
    await openApp(["-g", "-a", "Microsoft Outlook"]).catch(() => {});
    const ok = await until(async () => (await status()).outlookRunning, 30_000, 500);
    if (!ok) throw new OutlookError("Microsoft Outlook is not running and could not be started.");
  }
  const find = async () => {
    const ws = await windows();
    return ws.find((w) => w.id === "OutlookWindow" && w.main) ?? ws.find((w) => w.id === "OutlookWindow") ?? null;
  };
  let w = await find();
  if (w) return w;
  if ((await ax.call<{ bundleId?: string }>("frontmost")).bundleId === "com.apple.loginwindow") {
    throw new OutlookError("The Mac's screen is locked; macOS hides app windows from Accessibility until it is unlocked.");
  }
  // No main window (closed, or on a Space AX cannot see without AXMainWindow): ask Outlook to reopen one.
  const prev = await ax.call("frontmost");
  await openApp(["-a", "Microsoft Outlook"]).catch(() => {});
  w = await until(find, 8000, 250);
  if (prev?.pid) await ax.call("restore", { pid: prev.pid }).catch(() => {});
  if (!w) throw new OutlookError("Could not find Outlook's main window.");
  return w;
}

/**
 * Bring Outlook to the front for actions that need keyboard focus or menus, then give focus back to the
 * app the user was in. Nested calls only activate once.
 */
let focusDepth = 0;
export async function withFocus<T>(fn: () => Promise<T>): Promise<T> {
  if (focusDepth > 0) return fn();
  const front = await ax.call<{ bundleId?: string }>("frontmost");
  if (front.bundleId === "com.apple.loginwindow") {
    throw new OutlookError(
      "The Mac's screen is locked. This action needs Outlook in the foreground (menus, dialogs or typing); unlock the Mac and retry. Read-only tools still work while locked.",
    );
  }
  focusDepth++;
  let res: { previous: { pid?: number; bundleId?: string }; active: boolean };
  try {
    res = await ax.call("activate", { settleMs: 350 });
  } catch (e) {
    focusDepth--;
    throw e;
  }
  try {
    if (!res.active) throw new OutlookError("Could not bring Outlook to the front.");
    return await fn();
  } finally {
    focusDepth--;
    const prev = res.previous;
    if (prev?.pid && prev.bundleId !== "com.microsoft.Outlook" && process.env.AXLOOK_KEEP_FOCUS !== "1") {
      await ax.call("restore", { pid: prev.pid }).catch(() => {});
    }
  }
}

/**
 * Make Outlook frontmost again inside withFocus: another app (or the user) can take the focus back
 * mid-flow, and Outlook then ignores presses on some buttons.
 */
export async function bringToFront() {
  const res = await ax.call<{ active: boolean }>("activate", { settleMs: 150 });
  if (!res.active) throw new OutlookError("Could not bring Outlook to the front.");
}

export type Module = "Mail" | "Calendar";

/** Switch the main window to the Mail or Calendar module (works in the background). */
export async function switchModule(mod: Module): Promise<AxNode> {
  const w = await mainWindow();
  const isThere = async () => {
    if (mod === "Calendar") return ax.findOne("main", { id: "cal_grid_header_view_switcher" }, { depth: 2 });
    return ax.findOne("main", { role: "AXTable", desc: "Message List" }, { depth: 8 });
  };
  if (await isThere()) return w;
  const btn = await ax.findOne("main", { role: "AXButton", help: mod }, { depth: 3 });
  if (!btn) throw new OutlookError(`Could not find the ${mod} button in Outlook's sidebar.`);
  await ax.press(btn.h);
  const ok = await until(isThere, 6000);
  if (!ok) throw new OutlookError(`Outlook did not switch to ${mod}.`);
  return (await mainWindow());
}

export async function mainTitle(): Promise<string> {
  const w = await ax.get("main");
  return w.title ?? "";
}

/** Find a single element or throw a descriptive error. */
export async function need(h: string, match: Match, what: string, opts: { depth?: number; timeoutMs?: number } = {}): Promise<AxNode> {
  if (opts.timeoutMs) {
    const r = await ax.waitFor(h, match, { timeoutMs: opts.timeoutMs, depth: opts.depth });
    if (r.ok && r.node) return r.node;
  } else {
    const n = await ax.findOne(h, match, { depth: opts.depth });
    if (n) return n;
  }
  throw new OutlookError(`Could not find ${what} in Outlook's UI.`);
}

/** All open Outlook windows except the main one (compose, event editors, ...). */
export async function secondaryWindows(): Promise<AxNode[]> {
  return (await windows()).filter((w) => w.id !== "OutlookWindow" && w.role === "AXWindow" && w.subrole === "AXStandardWindow");
}

/**
 * The window that opened since `before` was read. Windows are told apart by title, and an older window
 * may carry the same title (e.g. a form left open): the newest one is listed first.
 */
export async function newWindow(before: AxNode[], accept: (w: AxNode) => boolean = () => true): Promise<AxNode | null> {
  const count = (ws: AxNode[], title?: string) => ws.filter((w) => w.title === title).length;
  const now = await secondaryWindows();
  return now.find((w) => accept(w) && count(now, w.title) > count(before, w.title)) ?? null;
}

/** The real body editor under `root`: Outlook also keeps tiny off-screen text areas, so take the largest. */
export async function bodyEditor(root: string, timeoutMs = 5000): Promise<AxNode> {
  const pick = async () => {
    const areas = await ax.find(root, { role: "AXTextArea" }, { depth: 14, limit: 20, frame: true });
    const visible = areas.filter((a) => a.frame && a.frame.w > 50 && a.frame.h >= 10 && a.frame.x > -1000);
    return visible.sort((a, b) => b.frame!.w * b.frame!.h - a.frame!.w * a.frame!.h)[0] ?? null;
  };
  const area = await until(pick, timeoutMs, 250);
  if (!area) throw new OutlookError("Could not find the message body editor in Outlook's UI.");
  return area;
}

/**
 * Focus a text area, optionally move the caret to the start or select everything, paste text (and HTML),
 * and verify the text arrived. Retries once; a freshly opened editor may not take focus immediately.
 */
export async function pasteInto(area: AxNode, text: string, opts: { html?: string; mode?: "start" | "replace" } = {}) {
  const probe = text.split(/\r?\n/).find((l) => l.trim().length > 3)?.trim().slice(0, 30);
  for (let attempt = 0; attempt < 2; attempt++) {
    const focused = await until(async () => {
      await ax.set(area.h, "AXFocused", true);
      return (await ax.get(area.h, { attrs: ["AXFocused"] })).attrs?.AXFocused === true;
    }, 4000, 300);
    if (!focused) continue;
    await sleep(250);
    if (opts.mode === "replace") await ax.menu(["Edit", "Select All"]).catch(() => {});
    else await ax.key(KEY.UP, ["cmd"]); // start of document (arrow keys are layout independent)
    await sleep(150);
    await ax.paste(text, opts.html);
    if (!probe) return;
    if (await until(async () => (await ax.text(area.h)).text.includes(probe), 4000, 250)) {
      await sleep(600); // let Outlook's editor commit the change
      return;
    }
  }
  throw new OutlookError("The text did not appear in Outlook's editor.");
}
