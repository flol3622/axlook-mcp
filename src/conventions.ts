// Per-user conventions: a small Markdown file (with optional front matter) describing how this user
// organizes mail and calendars, e.g. which calendar professional vs. personal events go in.
// It lives outside the plugin so it survives updates: $AXLOOK_CONFIG_DIR or ~/.config/axlook.
import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

/** Front-matter keys the server applies itself when a tool call leaves them out. */
export const DEFAULT_KEYS = ["default_calendar", "default_from"] as const;
export type DefaultKey = (typeof DEFAULT_KEYS)[number];

export interface Conventions {
  path: string;
  exists: boolean;
  /** Front-matter values the server applies itself (default_calendar, default_from). */
  defaults: Partial<Record<DefaultKey, string>>;
  /** Any other front-matter `key: value` lines: guidance for Claude only. */
  settings: Record<string, string>;
  /** Markdown after the front matter. */
  body: string;
  /** The whole file as stored. */
  raw: string;
}

/** Longest conventions file accepted, to keep tool instructions small. */
export const MAX_CHARS = 8000;

export function conventionsPath(env: NodeJS.ProcessEnv = process.env): string {
  const dir = env.AXLOOK_CONFIG_DIR || join(env.XDG_CONFIG_HOME || join(homedir(), ".config"), "axlook");
  return join(dir, "conventions.md");
}

/** Split `---\nkey: value\n---\nbody` into front matter and body. Unknown or malformed lines are ignored. */
export function parseConventions(raw: string): { defaults: Record<string, string>; body: string } {
  const text = raw.replace(/^﻿/, "").replace(/\r\n/g, "\n");
  const m = text.match(/^---\n([\s\S]*?)\n---[ \t]*(?:\n|$)/);
  if (!m) return { defaults: {}, body: text.trim() };
  const defaults: Record<string, string> = {};
  for (const line of m[1]!.split("\n")) {
    const kv = line.match(/^\s*([A-Za-z_][\w-]*)\s*:\s*(.*?)\s*$/);
    if (!kv || line.trimStart().startsWith("#")) continue;
    let value = kv[2]!;
    if (/^(["']).*\1$/.test(value)) value = value.slice(1, -1);
    if (value) defaults[kv[1]!.toLowerCase()] = value;
  }
  return { defaults, body: text.slice(m[0].length).trim() };
}

export function readConventions(path = conventionsPath()): Conventions {
  if (!existsSync(path)) return { path, exists: false, defaults: {}, settings: {}, body: "", raw: "" };
  const raw = readFileSync(path, "utf8");
  const { defaults: all, body } = parseConventions(raw);
  const defaults: Partial<Record<DefaultKey, string>> = {};
  const settings: Record<string, string> = {};
  for (const [k, v] of Object.entries(all)) {
    if ((DEFAULT_KEYS as readonly string[]).includes(k)) defaults[k as DefaultKey] = v;
    else settings[k] = v;
  }
  return { path, exists: true, raw, defaults, settings, body };
}

/**
 * Save the conventions file. The previous version is copied to conventions.md.bak; saving identical
 * content changes nothing (so a repeated save cannot wipe the backup). The write is atomic.
 */
export function writeConventions(content: string, path = conventionsPath()): Conventions & { changed: boolean } {
  const text = content.replace(/\r\n/g, "\n").trim() + "\n";
  if (text.length > MAX_CHARS) {
    throw new Error(`Conventions are limited to ${MAX_CHARS} characters (got ${text.length}); keep them to the rules that matter.`);
  }
  mkdirSync(join(path, ".."), { recursive: true });
  if (existsSync(path)) {
    if (readFileSync(path, "utf8") === text) return { ...readConventions(path), changed: false };
    copyFileSync(path, `${path}.bak`);
  }
  const tmp = `${path}.tmp-${process.pid}`;
  writeFileSync(tmp, text, { encoding: "utf8", mode: 0o600 });
  renameSync(tmp, path);
  return { ...readConventions(path), changed: true };
}

/** Text appended to the MCP server instructions so every client sees the user's conventions. */
export function conventionsInstructions(c: Conventions): string {
  if (!c.exists || (!c.body && !Object.keys(c.defaults).length && !Object.keys(c.settings).length)) {
    return "No personal conventions are set. If the user states a lasting preference (e.g. which calendar certain events go in), offer to save it with outlook_set_conventions.";
  }
  const list = (o: Record<string, string>) => Object.entries(o).map(([k, v]) => `${k}: ${v}`).join("; ");
  const defaults = list(c.defaults);
  const settings = list(c.settings);
  return [
    "The user's personal conventions (follow them unless the user asks otherwise for a specific call; outlook_get_conventions returns the current version):",
    defaults ? `Applied by the server when a call leaves them out: ${defaults} (default_calendar → outlook_create_event, default_from → outlook_create_draft only).` : "",
    settings ? `Other settings (guidance only, not applied by the server): ${settings}.` : "",
    c.body.slice(0, MAX_CHARS),
  ].filter(Boolean).join("\n");
}
