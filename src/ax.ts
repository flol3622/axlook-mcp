// Client for the Swift Accessibility helper (bin/axlook-helper), spoken to over JSON lines on stdio.
import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export interface AxNode {
  h: string;
  role?: string;
  subrole?: string;
  id?: string;
  title?: string;
  desc?: string;
  value?: string;
  help?: string;
  selected?: boolean;
  disabled?: boolean;
  frame?: { x: number; y: number; w: number; h: number };
  children?: AxNode[];
  more?: number;
  childCount?: number;
  x?: Record<string, unknown>;
  main?: boolean;
}

/** Matcher keys: role, subrole, id, title, desc, help, value, any; each with optional Prefix/Contains/Regex/Not suffix. */
export type Match = Record<string, string>;

const HERE = dirname(fileURLToPath(import.meta.url));
export const ROOT = resolve(HERE, "..");
/** AXLOOK_HELPER, else the helper next to the bundle (dist/), else the dev build (bin/). */
export const HELPER_PATH =
  process.env.AXLOOK_HELPER ??
  [join(HERE, "axlook-helper"), join(ROOT, "dist", "axlook-helper"), join(ROOT, "bin", "axlook-helper")].find((p) => existsSync(p)) ??
  join(ROOT, "bin", "axlook-helper");

type Pending = { resolve: (v: any) => void; reject: (e: Error) => void; timer: ReturnType<typeof setTimeout>; proc: ChildProcessWithoutNullStreams };

/** AX errors that mean the call certainly did nothing (stale element, AX disabled). Others are ambiguous. */
const FATAL_AX = new Map([
  [-25202, "the UI element no longer exists (Outlook's UI changed meanwhile)"],
  [-25211, "Accessibility API is disabled"],
]);

export class AxClient {
  private proc: ChildProcessWithoutNullStreams | null = null;
  private pending = new Map<number, Pending>();
  private nextId = 1;
  private buffer = "";

  private start() {
    if (!existsSync(HELPER_PATH)) {
      throw new Error(`Accessibility helper not found at ${HELPER_PATH}. Run: bun run build:helper`);
    }
    const proc = spawn(HELPER_PATH, [], { stdio: ["pipe", "pipe", "pipe"] });
    proc.stdout.setEncoding("utf8");
    proc.stdout.on("data", (chunk: string) => this.onData(chunk));
    proc.stderr.on("data", (d) => process.stderr.write(`[axlook-helper] ${d}`));
    const failAll = (why: string) => {
      if (this.proc === proc) this.proc = null;
      for (const [id, p] of this.pending) {
        if (p.proc !== proc) continue;
        clearTimeout(p.timer);
        this.pending.delete(id);
        p.reject(new Error(why));
      }
    };
    proc.on("exit", (code) => failAll(`Accessibility helper exited (code ${code})`));
    proc.on("error", (e) => failAll(`Accessibility helper failed: ${e.message}`));
    proc.stdin.on("error", (e) => failAll(`Accessibility helper pipe error: ${e.message}`));
    this.proc = proc;
  }

  private onData(chunk: string) {
    this.buffer += chunk;
    let nl: number;
    while ((nl = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, nl);
      this.buffer = this.buffer.slice(nl + 1);
      if (!line.trim()) continue;
      let msg: any;
      try {
        msg = JSON.parse(line);
      } catch {
        process.stderr.write(`[axlook-helper] bad line: ${line.slice(0, 200)}\n`);
        continue;
      }
      const p = this.pending.get(msg.id);
      if (!p) continue;
      this.pending.delete(msg.id);
      clearTimeout(p.timer);
      if (msg.ok) p.resolve(msg.result);
      else p.reject(new Error(msg.error ?? "helper error"));
    }
  }

  call<T = any>(cmd: string, args: Record<string, unknown> = {}, timeoutMs = 30_000): Promise<T> {
    if (!this.proc) this.start();
    const proc = this.proc!;
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Accessibility helper timed out on '${cmd}' (Outlook may be busy)`));
        // A stuck AX call usually means Outlook is busy; start a fresh helper for the next call.
        if (this.proc === proc) this.proc = null;
        proc.kill();
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer, proc });
      proc.stdin.write(JSON.stringify({ id, cmd, ...args }) + "\n");
    });
  }

  close() {
    const p = this.proc;
    this.proc = null;
    p?.kill();
  }

  // Typed conveniences -----------------------------------------------------

  find(h: string, match: Match, opts: { depth?: number; limit?: number; frame?: boolean; subtree?: number; extra?: string[] } = {}) {
    return this.call<AxNode[]>("find", { h, match, ...opts });
  }
  async findOne(h: string, match: Match, opts: { depth?: number; frame?: boolean; subtree?: number; extra?: string[] } = {}) {
    return (await this.find(h, match, { ...opts, limit: 1 }))[0] ?? null;
  }
  tree(h: string, depth = 4, opts: { maxChildren?: number; frame?: boolean; extra?: string[] } = {}) {
    return this.call<AxNode>("tree", { h, depth, ...opts });
  }
  children(h: string, opts: { frame?: boolean; extra?: string[] } = {}) {
    return this.call<AxNode[]>("children", { h, ...opts });
  }
  get(h: string, opts: { attrs?: string[]; actions?: boolean; frame?: boolean } = {}) {
    return this.call<AxNode & { attrs?: Record<string, any>; actions?: string[] }>("get", { h, ...opts });
  }
  async press(h: string, action?: string) {
    const r = await this.call<{ axError: number }>("press", { h, action });
    if (FATAL_AX.has(r.axError)) throw new Error(`Could not ${action ?? "press"}: ${FATAL_AX.get(r.axError)}`);
    return r;
  }
  async set(h: string, attr: string, value: string | boolean | number, type?: "string" | "bool" | "number" | "date") {
    const t = type ?? (typeof value === "boolean" ? "bool" : typeof value === "number" ? "number" : "string");
    const r = await this.call<{ axError: number }>("set", { h, attr, value, type: t });
    if (FATAL_AX.has(r.axError)) throw new Error(`Could not set ${attr}: ${FATAL_AX.get(r.axError)}`);
    return r;
  }
  waitFor(h: string, match: Match, opts: { timeoutMs?: number; depth?: number; gone?: boolean; frame?: boolean } = {}) {
    return this.call<{ ok: boolean; node?: AxNode | null }>("waitFor", { h, match, ...opts }, (opts.timeoutMs ?? 5000) + 10_000);
  }
  text(h: string) {
    return this.call<{ text: string; links: { text: string; url: string }[] }>("text", { h });
  }
  menu(path: string[], check = false) {
    return this.call("menu", { path, check });
  }
  key(code: number, mods: string[] = []) {
    return this.call("key", { code, mods });
  }
  type(text: string) {
    return this.call("type", { text }, 60_000);
  }
  /** Paste via Edit > Paste (Outlook must be frontmost); the user's clipboard is restored afterwards. */
  paste(text: string, html?: string, settleMs = 1200) {
    return this.call("paste", { text, html, settleMs });
  }
  click(h: string) {
    return this.call("click", { h });
  }
  sleep(ms: number) {
    return this.call("sleep", { ms }, ms + 10_000);
  }
}

/** Layout-independent key codes only. Letter shortcuts depend on the keyboard layout: use menus instead. */
export const KEY = { RETURN: 36, TAB: 48, DELETE: 51, ESCAPE: 53, FORWARD_DELETE: 117, HOME: 115, END: 119, DOWN: 125, UP: 126 } as const;
