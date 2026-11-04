#!/usr/bin/env bun
// MCP server exposing New Outlook for Mac (mail, drafts, calendars) by driving its UI through Accessibility.
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { rmSync } from "node:fs";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ax, exclusive, OutlookError, status, windows } from "./outlook/session";
import * as mail from "./outlook/mail";
import * as compose from "./outlook/compose";
import * as cal from "./outlook/calendar";
import { parseEml, saveAttachments } from "./mime";
import pkg from "../package.json";

const server = new McpServer(
  { name: "axlook", version: pkg.version },
  {
    instructions: [
      "Controls the user's running New Outlook for Mac (all accounts, including Exchange) through macOS Accessibility.",
      "Use these tools for every question about the user's email or calendars (reading, searching, drafts, events, availability), even when Outlook is not named, unless the user explicitly names another source.",
      "Workflow: outlook_list_folders → outlook_list_messages / outlook_search_messages (return `ref`s) → read/move/flag with those refs.",
      "Refs identify a row by folder + subject + sender + date; if a ref goes stale, list again.",
      "Calendar: outlook_list_calendars → outlook_list_events (YYYY-MM-DD range) → refs for get/update/delete.",
      "Nothing is ever sent: drafts are saved to Drafts for the user to review and send.",
      "Calls are serialized and take seconds (it is UI automation). Tools that need menus or typing briefly bring Outlook to the front and then return focus.",
    ].join("\n"),
  },
);

type Json = Record<string, unknown> | unknown[];

function ok(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
}

function fail(e: unknown) {
  const msg = e instanceof OutlookError ? e.message : e instanceof Error ? `Unexpected error: ${e.message}` : String(e);
  return { content: [{ type: "text" as const, text: msg }], isError: true };
}

/** Wrap a handler: serialize UI access and convert errors into tool errors. */
function run<A>(fn: (args: A) => Promise<unknown>) {
  return async (args: A) => {
    try {
      return ok(await exclusive(() => fn(args)));
    } catch (e) {
      return fail(e);
    }
  };
}

const ref = z.string().describe("Message ref returned by outlook_list_messages / outlook_search_messages");
const eventRef = z.string().describe("Event ref returned by outlook_list_events");
const RO = { readOnlyHint: true, openWorldHint: false } as const;
const WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: false } as const;

const expandHome = (p: string) => (p.startsWith("~") ? join(homedir(), p.slice(1)) : p);

// ------------------------------------------------------------------------------------------------
// Diagnostics

server.registerTool(
  "outlook_status",
  {
    description: "Check that Outlook is running, Accessibility permission is granted, and list Outlook's windows.",
    inputSchema: {},
    annotations: RO,
  },
  run(async () => {
    const st = await status();
    const front = await ax.call("frontmost");
    return {
      ...st,
      screenLocked: front?.bundleId === "com.apple.loginwindow",
      windows: st.trusted && st.outlookRunning ? (await windows()).map((w) => ({ title: w.title, id: w.id, main: !!w.main })) : [],
    };
  }),
);

server.registerTool(
  "outlook_ui_snapshot",
  {
    description:
      "Debug aid: dump part of Outlook's accessibility tree (roles, ids, labels). Use only when another tool fails because Outlook's UI changed.",
    inputSchema: {
      window: z.enum(["main", "focused"]).default("main"),
      depth: z.number().int().min(1).max(12).default(5),
      match: z.string().optional().describe("Only dump subtrees whose label/value/id contains this text"),
    },
    annotations: RO,
  },
  run(async ({ window, depth, match }: { window: "main" | "focused"; depth: number; match?: string }) => {
    const root = window === "main" ? "main" : "focusedWindow";
    if (match) return ax.find(root, { anyContains: match }, { depth: 20, limit: 10, subtree: depth });
    return ax.tree(root, depth, { maxChildren: 60 });
  }),
);

// ------------------------------------------------------------------------------------------------
// Mail

server.registerTool(
  "outlook_list_folders",
  {
    description:
      "List mail folders of every account as paths like 'account/Inbox' or 'account/Parent/Child', with unread and total counts. 'All Accounts/…' are unified views.",
    inputSchema: {
      expand: z.boolean().default(false).describe("Also reveal subfolders hidden under collapsed folders (slower)"),
    },
    annotations: RO,
  },
  run(async ({ expand }: { expand: boolean }) => mail.listFolders({ expand })),
);

server.registerTool(
  "outlook_list_messages",
  {
    description: "List messages in a folder (newest first) with sender, subject, date, preview and flags. Returns refs for other tools.",
    inputSchema: {
      folder: z.string().describe("Folder path from outlook_list_folders, e.g. 'you@example.com/Inbox' (a unique folder name also works)"),
      limit: z.number().int().min(1).max(200).default(25),
      offset: z.number().int().min(0).default(0),
      unread_only: z.boolean().default(false),
    },
    annotations: RO,
  },
  run(async (a: { folder: string; limit: number; offset: number; unread_only: boolean }) =>
    mail.listMessages({ folder: a.folder, limit: a.limit, offset: a.offset, unreadOnly: a.unread_only }),
  ),
);

server.registerTool(
  "outlook_search_messages",
  {
    description:
      "Search mail with Outlook's search box (supports Outlook query syntax such as from:alice subject:report hasattachments:yes). The scope is the whole mailbox of `folder`; default searches all accounts.",
    inputSchema: {
      query: z.string().min(1),
      folder: z.string().optional().describe("Any folder of the mailbox to search; default 'All Accounts/Inbox' (all accounts)"),
      limit: z.number().int().min(1).max(200).default(25),
    },
    annotations: RO,
  },
  run(async (a: { query: string; folder?: string; limit: number }) => mail.searchMessages(a)),
);

server.registerTool(
  "outlook_read_message",
  {
    description:
      "Read a message as displayed by Outlook: subject, from, recipients, date, body text, links and attachment names. Keeps it unread unless mark_as_read. For a conversation row, all visible messages are included.",
    inputSchema: { ref, mark_as_read: z.boolean().default(false) },
    annotations: RO,
  },
  run(async (a: { ref: string; mark_as_read: boolean }) => mail.readMessage(a.ref, { markAsRead: a.mark_as_read })),
);

server.registerTool(
  "outlook_get_message_source",
  {
    description:
      "Export the full message (File > Save As .eml) and return parsed headers (Message-ID, To/Cc, dates), plain text, optionally HTML, and attachment metadata. Briefly brings Outlook to the front.",
    inputSchema: {
      ref,
      include_html: z.boolean().default(false),
      save_eml_to: z.string().optional().describe("Optional path to keep the .eml file (otherwise it is deleted)"),
    },
    annotations: RO,
  },
  run(async (a: { ref: string; include_html: boolean; save_eml_to?: string }) => {
    const target = a.save_eml_to ? resolve(expandHome(a.save_eml_to)) : undefined;
    const path = await mail.exportEml(a.ref, target);
    try {
      const msg = await parseEml(path);
      return { ...msg, html: a.include_html ? msg.html : undefined, emlPath: target ?? null };
    } finally {
      if (!target) rmSync(path, { force: true });
    }
  }),
);

server.registerTool(
  "outlook_save_attachments",
  {
    description:
      "Download a message's attachments to a local folder (via the message's .eml export). By default saves all non-inline attachments; never overwrites files.",
    inputSchema: {
      ref,
      directory: z.string().default("~/Downloads/Outlook Attachments").describe("Destination folder (created if missing)"),
      names: z.array(z.string()).optional().describe("Only these attachment file names"),
      include_inline: z.boolean().default(false).describe("Also save inline images"),
    },
    annotations: WRITE,
  },
  run(async (a: { ref: string; directory: string; names?: string[]; include_inline: boolean }) => {
    const path = await mail.exportEml(a.ref);
    try {
      const msg = await parseEml(path);
      const saved = await saveAttachments(path, resolve(expandHome(a.directory)), { names: a.names, includeInline: a.include_inline });
      return { subject: msg.subject, attachmentsInMessage: msg.attachments.length, saved };
    } finally {
      rmSync(path, { force: true });
    }
  }),
);

server.registerTool(
  "outlook_move_message",
  {
    description: "Move a message to another folder (within any account). Briefly brings Outlook to the front.",
    inputSchema: { ref, destination: z.string().describe("Destination folder path from outlook_list_folders") },
    annotations: WRITE,
  },
  run(async (a: { ref: string; destination: string }) => mail.moveMessage(a.ref, a.destination)),
);

server.registerTool(
  "outlook_mark_read",
  {
    description: "Mark a message as read or unread.",
    inputSchema: { ref, read: z.boolean() },
    annotations: { ...WRITE, idempotentHint: true },
  },
  run(async (a: { ref: string; read: boolean }) => mail.setReadState(a.ref, a.read)),
);

server.registerTool(
  "outlook_set_flag",
  {
    description: "Flag or unflag a message for follow-up.",
    inputSchema: { ref, flagged: z.boolean() },
    annotations: { ...WRITE, idempotentHint: true },
  },
  run(async (a: { ref: string; flagged: boolean }) => mail.setFlag(a.ref, a.flagged)),
);

server.registerTool(
  "outlook_archive_message",
  {
    description: "Archive a message (Outlook's Archive action for its account).",
    inputSchema: { ref },
    annotations: WRITE,
  },
  run(async (a: { ref: string }) => mail.archiveMessage(a.ref)),
);

server.registerTool(
  "outlook_delete_message",
  {
    description: "Move a message to Deleted Items (never a permanent delete; refused inside Deleted Items).",
    inputSchema: { ref },
    annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
  },
  run(async (a: { ref: string }) => mail.deleteMessage(a.ref)),
);

// ------------------------------------------------------------------------------------------------
// Drafts (never sent)

const addresses = z.array(z.string()).optional();

server.registerTool(
  "outlook_create_draft",
  {
    description: "Create a new email draft in Outlook's Drafts folder. It is NOT sent; the user reviews and sends it.",
    inputSchema: {
      to: addresses,
      cc: addresses,
      bcc: addresses,
      subject: z.string().optional(),
      body: z.string().optional(),
      html: z.boolean().default(false).describe("Body is HTML"),
      from: z.string().optional().describe("Sending account e-mail address (defaults to Outlook's default account)"),
    },
    annotations: WRITE,
  },
  run(async (a: compose.DraftInput) => compose.createDraft(a)),
);

server.registerTool(
  "outlook_create_reply_draft",
  {
    description: "Create a reply (or reply-all) draft to a message, with your text above the quoted original. Saved to Drafts, NOT sent.",
    inputSchema: { ref, body: z.string(), reply_all: z.boolean().default(false), html: z.boolean().default(false) },
    annotations: WRITE,
  },
  run(async (a: { ref: string; body: string; reply_all: boolean; html: boolean }) =>
    compose.respondDraft(a.ref, a.reply_all ? "replyAll" : "reply", a.body, { html: a.html }),
  ),
);

server.registerTool(
  "outlook_create_forward_draft",
  {
    description: "Create a forward draft of a message (attachments included by Outlook). Saved to Drafts, NOT sent.",
    inputSchema: { ref, to: z.array(z.string()).default([]), body: z.string().default(""), html: z.boolean().default(false) },
    annotations: WRITE,
  },
  run(async (a: { ref: string; to: string[]; body: string; html: boolean }) => compose.respondDraft(a.ref, "forward", a.body, { to: a.to, html: a.html })),
);

// ------------------------------------------------------------------------------------------------
// Calendar

server.registerTool(
  "outlook_list_calendars",
  {
    description: "List calendars of all accounts as paths like 'account/My Calendars/Calendar', and whether each is shown.",
    inputSchema: {},
    annotations: RO,
  },
  run(async () => cal.listCalendars()),
);

server.registerTool(
  "outlook_list_events",
  {
    description:
      "List events between two dates (inclusive, YYYY-MM-DD) with times, location, recurrence, organizer, your response and attendee summary. By default covers the calendars currently shown in Outlook; pass `calendars` to read specific ones (temporarily toggled, then restored).",
    inputSchema: {
      start: z.string().describe("YYYY-MM-DD"),
      end: z.string().optional().describe("YYYY-MM-DD (default: same as start)"),
      calendars: z.array(z.string()).optional().describe("Calendar paths or unique names from outlook_list_calendars"),
    },
    annotations: RO,
  },
  run(async (a: { start: string; end?: string; calendars?: string[] }) => cal.listEvents(a)),
);

server.registerTool(
  "outlook_get_event",
  {
    description: "Get one event. full=true opens the event window for the calendar name, full attendee list and description (brings Outlook forward briefly).",
    inputSchema: { ref: eventRef, full: z.boolean().default(false) },
    annotations: RO,
  },
  run(async (a: { ref: string; full: boolean }) => cal.getEvent(a.ref, { full: a.full })),
);

const eventFields = {
  location: z.string().optional(),
  body: z.string().optional().describe("Description / notes"),
  all_day: z.boolean().optional(),
};

server.registerTool(
  "outlook_create_event",
  {
    description:
      "Add an event (appointment without attendees) to a calendar. Times are local wall-clock: start/end like 2026-11-20T10:00 (all-day: dates).",
    inputSchema: {
      subject: z.string().min(1),
      start: z.string(),
      end: z.string().optional().describe("Required unless all_day"),
      calendar: z.string().optional().describe("Calendar path or unique name (default: Outlook's default calendar)"),
      ...eventFields,
    },
    annotations: WRITE,
  },
  run(async (a: { subject: string; start: string; end?: string; calendar?: string; location?: string; body?: string; all_day?: boolean }) =>
    cal.createEvent({ subject: a.subject, start: a.start, end: a.end, calendar: a.calendar, location: a.location, body: a.body, allDay: a.all_day }),
  ),
);

server.registerTool(
  "outlook_update_event",
  {
    description: "Edit an event you own: change subject, times, all-day, location or description. Never sends updates to attendees (meetings with attendees are refused).",
    inputSchema: { ref: eventRef, subject: z.string().optional(), start: z.string().optional(), end: z.string().optional(), ...eventFields },
    annotations: WRITE,
  },
  run(async (a: { ref: string; subject?: string; start?: string; end?: string; location?: string; body?: string; all_day?: boolean }) =>
    cal.updateEvent(a.ref, { subject: a.subject, start: a.start, end: a.end, location: a.location, body: a.body, allDay: a.all_day }),
  ),
);

server.registerTool(
  "outlook_delete_event",
  {
    description: "Delete an event (single occurrence for recurring events). Refuses if Outlook would notify attendees.",
    inputSchema: { ref: eventRef },
    annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
  },
  run(async (a: { ref: string }) => cal.deleteEvent(a.ref)),
);

// ------------------------------------------------------------------------------------------------

const transport = new StdioServerTransport();
await server.connect(transport);
process.on("SIGINT", () => {
  ax.close();
  process.exit(0);
});
process.stdin.on("close", () => {
  ax.close();
  process.exit(0);
});

export type { Json };
