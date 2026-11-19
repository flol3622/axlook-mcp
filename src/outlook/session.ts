// Shared Outlook UI session: one Accessibility helper, one queue (the UI is a single shared resource),
// focus management and module (Mail / Calendar) switching.
import { AsyncLocalStorage } from "node:async_hooks";
import { execFile } from "node:child_process";
import { AxClient, KEY, type AxNode, type Match } from "../ax";

export class OutlookError extends Error {}

export const ax = new AxClient();

/** The tool call a piece of UI work belongs to; set `cancelled` once its deadline has passed. */
interface CallContext {
  label: string;
  cancelled: boolean;
}
const callContext = new AsyncLocalStorage<CallContext>();

/** Stop work whose call was abandoned, at its next helper command or poll. */
export function checkCancelled() {
  const c = callContext.getStore();
  if (c?.cancelled) throw new OutlookError(`${c.label} was cancelled after its deadline.`);
}
ax.guard = checkCancelled;

let queue: Promise<unknown> = Promise.resolve();
let queued = 0;
let current: { label: string; since: number } | null = null;

/** What the queue is doing, without waiting for it. */
export function busy(): { tool: string; runningForSeconds: number; queued: number } | null {
  return current ? { tool: current.label, runningForSeconds: Math.round((Date.now() - current.since) / 1000), queued } : null;
}

/** Default overall deadline of one tool call; AXLOOK_CALL_TIMEOUT (seconds) overrides it. */
export const CALL_TIMEOUT_MS = (Number(process.env.AXLOOK_CALL_TIMEOUT) || 90) * 1000;

/**
 * Run UI work strictly one call at a time. Handles from previous calls are released first.
 * Each call has an overall deadline: past it, the call fails, its remaining helper commands and polls are
 * refused, and the helper is restarted, so one stuck call cannot hold the queue (and Outlook) for minutes.
 */
export function exclusive<T>(fn: () => Promise<T>, opts: { label?: string; timeoutMs?: number } = {}): Promise<T> {
  const label = opts.label ?? "call";
  const timeoutMs = opts.timeoutMs ?? CALL_TIMEOUT_MS;
  queued++;
  const run = queue.then(async () => {
    queued--;
    const ctx: CallContext = { label, cancelled: false };
    current = { label, since: Date.now() };
    const work = callContext.run(ctx, async () => {
      await ax.call("release").catch(() => {});
      return fn();
    });
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timedOut = new OutlookError(
      `${label} timed out after ${Math.round(timeoutMs / 1000)} s; Outlook may be busy or showing something unexpected. Later calls start from a fresh state.`,
    );
    const deadline = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(timedOut), timeoutMs);
    });
    try {
      return await Promise.race([work, deadline]);
    } catch (e) {
      if (e === timedOut) {
        ctx.cancelled = true;
        ax.close(); // ends a helper command that is stuck inside Outlook
        await settled(work, 15_000);
        await resetUi().catch(() => {});
      }
      throw e;
    } finally {
      clearTimeout(timer);
      current = null;
    }
  });
  queue = run.catch(() => {});
  return run;
}

/** Whether `p` settles within `ms`. */
function settled(p: Promise<unknown>, ms: number): Promise<boolean> {
  return Promise.race([p.then(() => true, () => true), sleep(ms).then(() => false)]);
}

/** After an abandoned call: close a sheet or menu it may have left open in the main window. */
async function resetUi() {
  const sheet = await ax.findOne("main", { role: "AXSheet" }, { depth: 1 });
  const cancel = sheet ? await ax.findOne(sheet.h, { role: "AXButton", titleRegex: "^Cancel$" }, { depth: 3 }) : null;
  if (cancel) await ax.press(cancel.h).catch(() => {});
  const menu = await ax.findOne("app", { role: "AXMenu" }, { depth: 3 }).catch(() => null);
  if (menu) await ax.call("press", { h: menu.h, action: "AXCancel" }).catch(() => {});
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
    checkCancelled();
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
      // Give the focus back even when the call was cancelled meanwhile.
      await callContext.exit(() => ax.call("restore", { pid: prev.pid })).catch(() => {});
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

/** Whether the main window shows the Mail or Calendar module. */
async function inModule(mod: Module): Promise<boolean> {
  if (mod === "Calendar") {
    // The header's controls move into nested groups when the window is narrow or a pane is open.
    return !!(await ax.findOne("main", { idRegex: "^cal_grid_header_(view_switcher|today|date_range)$" }, { depth: 5 }));
  }
  return !!(await ax.findOne("main", { role: "AXTable", desc: "Message List" }, { depth: 8 }));
}

async function pressModuleButton(mod: Module) {
  const btn = await ax.findOne("main", { role: "AXButton", help: mod }, { depth: 3 });
  if (!btn) throw new OutlookError(`Could not find the ${mod} button in Outlook's sidebar.`);
  await ax.press(btn.h);
}

/** Switch the main window to the Mail or Calendar module (works in the background). */
export async function switchModule(mod: Module): Promise<AxNode> {
  const w = await mainWindow();
  if (await inModule(mod)) return w;
  await pressModuleButton(mod);
  if (!(await until(() => inModule(mod), 6000))) {
    // Pressing the button of the module already shown does nothing, and that module can be in a state the
    // check above does not recognise (e.g. an item open in it): go through the other module and back.
    const other: Module = mod === "Mail" ? "Calendar" : "Mail";
    await pressModuleButton(other);
    await until(() => inModule(other), 6000);
    await pressModuleButton(mod);
    if (!(await until(() => inModule(mod), 8000))) throw new OutlookError(`Outlook did not switch to ${mod}.`);
  }
  return mainWindow();
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

/**
 * The real body editor under `root`: Outlook also keeps tiny or off-screen text areas, so take the largest one
 * inside the root's frame. Frames are global screen coordinates (a display left of or above the main one has
 * negative ones), so compare with the root's own frame rather than with fixed bounds.
 */
export async function bodyEditor(root: string, timeoutMs = 5000): Promise<AxNode> {
  const pick = async () => {
    const box = (await ax.get(root, { frame: true }).catch(() => null))?.frame;
    const areas = await ax.find(root, { role: "AXTextArea" }, { depth: 16, limit: 20, frame: true });
    const inside = (f: NonNullable<AxNode["frame"]>) =>
      !box || (f.x + f.w > box.x && f.x < box.x + box.w && f.y + f.h > box.y && f.y < box.y + box.h);
    const visible = areas.filter((a) => a.frame && a.frame.w > 50 && a.frame.h >= 10 && inside(a.frame));
    return visible.sort((a, b) => b.frame!.w * b.frame!.h - a.frame!.w * a.frame!.h)[0] ?? null;
  };
  const area = await until(pick, timeoutMs, 250);
  if (!area) throw new OutlookError("Could not find the message body editor in Outlook's UI.");
  return area;
}

/**
 * Press a menu item, waiting for Outlook to enable it. Items stay disabled while a message is still being
 * downloaded, while a window that was just opened is not yet key, or while another app took the focus back.
 */
export async function pressMenu(path: string[], opts: { timeoutMs?: number; why?: string } = {}) {
  const deadline = Date.now() + (opts.timeoutMs ?? 8000);
  for (let attempt = 0; ; attempt++) {
    try {
      await ax.menu(path, true);
      break;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      if (!/menu item disabled/.test(msg)) throw e;
      if (Date.now() >= deadline) {
        throw new OutlookError(`Outlook kept ${path.join(" > ")} disabled${opts.why ? `: ${opts.why}` : "."}`);
      }
      if (attempt % 3 === 2) await bringToFront().catch(() => {});
      await sleep(400);
    }
  }
  return ax.menu(path);
}

/** Make `win` the key window (so menus and typing apply to it) and wait until it is. */
export async function focusWindow(win: AxNode, timeoutMs = 3000): Promise<boolean> {
  // Every lookup returns a new handle, so compare the windows by title and frame.
  const sig = (w: AxNode | null) => (w ? `${w.id ?? ""}|${w.title ?? ""}|${JSON.stringify(w.frame)}` : null);
  const isKey = async () => {
    const [f, w] = await Promise.all([ax.get("focusedWindow", { frame: true }).catch(() => null), ax.get(win.h, { frame: true }).catch(() => null)]);
    return !!f && !!w && sig(f) === sig(w);
  };
  if (await isKey()) return true;
  await ax.press(win.h, "AXRaise").catch(() => {});
  await ax.set(win.h, "AXMain", true).catch(() => {});
  return !!(await until(isKey, timeoutMs, 150));
}

/**
 * Focus a text area, optionally move the caret to the start or select everything, paste text (and HTML),
 * and verify the text arrived. Retries once; a freshly opened editor may not take focus immediately.
 */
export async function pasteInto(area: AxNode, text: string, opts: { html?: string; mode?: "start" | "replace" } = {}) {
  const probe = text.split(/\r?\n/).find((l) => l.trim().length > 3)?.trim().slice(0, 30);
  const arrived = async () => !!probe && (await ax.text(area.h)).text.replace(/\s+/g, " ").includes(probe.replace(/\s+/g, " "));
  for (let attempt = 0; attempt < 2; attempt++) {
    // A slow first paste may land late: never paste twice.
    if (attempt > 0 && (await until(arrived, 3000, 300))) return;
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
    if (await until(arrived, 4000, 250)) {
      await sleep(600); // let Outlook's editor commit the change
      return;
    }
  }
  throw new OutlookError("The text did not appear in Outlook's editor.");
}
