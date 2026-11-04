import { describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { conventionsInstructions, conventionsPath, MAX_CHARS, parseConventions, readConventions, writeConventions } from "../src/conventions";

const SAMPLE = `---
default_calendar: me@example.com/My Calendars/Work
default_from: "me@example.com"
# a comment line
broken line without colon
---
# Calendars
- Work → \`me@example.com/My Calendars/Work\`
- Social with my partner → \`me@gmail.example/Shared\`
`;

describe("conventions path", () => {
  test("AXLOOK_CONFIG_DIR wins", () => {
    expect(conventionsPath({ AXLOOK_CONFIG_DIR: "/tmp/ax", XDG_CONFIG_HOME: "/x" })).toBe("/tmp/ax/conventions.md");
  });
  test("XDG_CONFIG_HOME next", () => {
    expect(conventionsPath({ XDG_CONFIG_HOME: "/x" })).toBe("/x/axlook/conventions.md");
  });
  test("~/.config by default", () => {
    expect(conventionsPath({})).toMatch(/\/\.config\/axlook\/conventions\.md$/);
  });
});

describe("parseConventions", () => {
  test("front matter keys, quotes stripped, junk ignored", () => {
    const { defaults, body } = parseConventions(SAMPLE);
    expect(defaults).toEqual({ default_calendar: "me@example.com/My Calendars/Work", default_from: "me@example.com" });
    expect(body.startsWith("# Calendars")).toBe(true);
  });
  test("no front matter: whole text is the body", () => {
    expect(parseConventions("# Just rules\n- a\n")).toEqual({ defaults: {}, body: "# Just rules\n- a" });
  });
  test("CRLF and BOM tolerated", () => {
    const { defaults, body } = parseConventions("﻿---\r\ndefault_from: a@b.c\r\n---\r\nRule\r\n");
    expect(defaults.default_from).toBe("a@b.c");
    expect(body).toBe("Rule");
  });
  test("empty values are dropped", () => {
    expect(parseConventions("---\ndefault_calendar:\n---\nx").defaults).toEqual({});
  });
});

describe("read and write", () => {
  const dir = mkdtempSync(join(tmpdir(), "axlook-conv-"));
  const path = join(dir, "nested", "conventions.md");

  test("missing file reads as empty", () => {
    const c = readConventions(path);
    expect(c.exists).toBe(false);
    expect(conventionsInstructions(c)).toContain("No personal conventions");
  });
  test("write creates folders, private mode, and keeps a backup", () => {
    writeConventions("first", path);
    const c = writeConventions(SAMPLE, path);
    expect(c.exists).toBe(true);
    expect(c.defaults.default_calendar).toBe("me@example.com/My Calendars/Work");
    expect(readFileSync(`${path}.bak`, "utf8")).toBe("first\n");
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });
  test("instructions include defaults and rules", () => {
    const text = conventionsInstructions(readConventions(path));
    expect(text).toContain("default_calendar: me@example.com/My Calendars/Work");
    expect(text).toContain("Social with my partner");
  });
  test("saving identical content is a no-op and keeps the backup", () => {
    const again = writeConventions(SAMPLE, path);
    expect(again.changed).toBe(false);
    expect(readFileSync(`${path}.bak`, "utf8")).toBe("first\n");
  });
  test("only default_calendar/default_from are server-applied; other keys are guidance", () => {
    writeConventions("---\ndefault_from: me@example.com\ntimezone: Europe/Brussels\n---\n- rule", path);
    const c = readConventions(path);
    expect(c.defaults).toEqual({ default_from: "me@example.com" });
    expect(c.settings).toEqual({ timezone: "Europe/Brussels" });
    const text = conventionsInstructions(c);
    expect(text).toContain("Applied by the server when a call leaves them out: default_from: me@example.com");
    expect(text).toContain("Other settings (guidance only, not applied by the server): timezone: Europe/Brussels");
    writeConventions(SAMPLE, path);
  });
  test("oversized content is refused and the file is untouched", () => {
    expect(() => writeConventions("x".repeat(MAX_CHARS + 1), path)).toThrow(/limited/);
    expect(readConventions(path).defaults.default_from).toBe("me@example.com");
  });
  test("hand-edited files are read as they are", () => {
    writeFileSync(path, "- only a rule\n");
    expect(readConventions(path).body).toBe("- only a rule");
  });
});
