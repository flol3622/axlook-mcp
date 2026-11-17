import { describe, expect, test } from "bun:test";
import { expandedConversation, parseRowDesc, sameSubject, samePreview } from "../src/outlook/mail";
import { parseLongDate, parseRange, ymd } from "../src/outlook/calendar";

// Row descriptions captured from New Outlook 16.114.
describe("parseRowDesc", () => {
  test("plain message", () => {
    const r = parseRowDesc("     Sender: Alice Martin, Subject: Conf 2027 proceedings – indexing,      Yesterday,        Message preview: Dear Organising Committee, I am considering, ")!;
    expect(r.sender).toBe("Alice Martin");
    expect(r.subject).toBe("Conf 2027 proceedings – indexing");
    expect(r.date).toBe("Yesterday");
    expect(r.preview).toBe("Dear Organising Committee, I am considering");
    expect(r.unread).toBe(false);
    expect(r.flagged).toBe(false);
  });

  test("unread", () => {
    const r = parseRowDesc(" Unread,    Sender: Bob Ngata, Subject: Conf27 – Official Visa Invitation,      28/10/2026,        Message preview: Dear, ")!;
    expect(r.unread).toBe(true);
    expect(r.date).toBe("28/10/2026");
  });

  test("conversation with replied marker", () => {
    const r = parseRowDesc("   Conversation, 5 messages, Sender: Dirk Peeters, Subject: Abstract submission,   Replied Latest message:  23/09/2026,        Message preview: Thanks, ")!;
    expect(r.conversationCount).toBe(5);
    expect(r.replied).toBe(true);
    expect(r.subject).toBe("Abstract submission");
    expect(r.date).toBe("23/09/2026");
  });

  test("flagged", () => {
    const r = parseRowDesc("     Sender: Eva Lind, Subject: Questions regard potential technical papers,      24/08/2026,  Flagged,      Message preview: Hello, ")!;
    expect(r.flagged).toBe(true);
    expect(r.subject).toBe("Questions regard potential technical papers");
  });

  test("search result with folder", () => {
    const r = parseRowDesc("     Sender: Carl Jensen, Subject: Registered proceedings,      29/10/2026, Folder: Inbox,       Message preview: Dear Secretary, ")!;
    expect(r.folder).toBe("Inbox");
    expect(r.date).toBe("29/10/2026");
  });

  test("draft with several recipients as sender and time as date", () => {
    const r = parseRowDesc("     Sender: me@example.com, team, Subject: MCP draft test éà,      14:55,        Message preview: Line 1 Line 2 ✓, ")!;
    expect(r.sender).toBe("me@example.com, team");
    expect(r.subject).toBe("MCP draft test éà");
    expect(r.date).toBe("14:55");
  });

  test("no subject", () => {
    const r = parseRowDesc("     Sender: (no recipient), No subject,      30/10/2026,        Message preview: Sent from Outlook for Mac, ")!;
    expect(r.subject).toBe("");
    expect(r.sender).toBe("(no recipient)");
    expect(r.date).toBe("30/10/2026");
  });

  test("subject containing commas, a weekday and the word Flagged is not split", () => {
    const r = parseRowDesc("     Sender: A, Subject: Re: Meeting, Monday 3, Flagged items list,      Today,        Message preview: x, ")!;
    expect(r.subject).toBe("Re: Meeting, Monday 3, Flagged items list");
    expect(r.flagged).toBe(false);
    expect(r.date).toBe("Today");
  });

  test("attachments, meeting, completed and conversation unread markers", () => {
    const r = parseRowDesc(" 1 unread message,  Conversation, 5 messages, Sender: A, Subject: Plan,      Has files,  Meeting message, No Conflicts  Latest message:  30/10/2026,  Marked as completed,      Message preview: x, ")!;
    expect(r.hasAttachments).toBe(true);
    expect(r.meeting).toBe(true);
    expect(r.completed).toBe(true);
    expect(r.unread).toBe(true);
    expect(r.unreadInConversation).toBe(1);
    expect(r.conversationCount).toBe(5);
    expect(r.date).toBe("30/10/2026");
  });

  test("group header is not a message", () => {
    expect(parseRowDesc("Yesterday, Expanded")).toBeNull();
  });
});

describe("expandedConversation", () => {
  test("header of an expanded conversation", () => {
    const header = "  Expanded, Conversation, 3 messages,  Subject: Abstract submission,   Latest message:  26/06/2026, Folder: Inbox,  ";
    expect(expandedConversation(header)).toBe(3);
    expect(parseRowDesc(header)).toBeNull();
  });
  test("collapsed conversations and group headers are not", () => {
    expect(expandedConversation("   Conversation, 5 messages, Sender: Dirk Peeters, Subject: Abstract submission,   Latest message:  23/09/2026, ")).toBeNull();
    expect(expandedConversation("Top Results, Expanded")).toBeNull();
  });
});

describe("meeting messages", () => {
  test("invitation: the received date is the last date, not the meeting's", () => {
    const r = parseRowDesc("     Sender: Ann Lee, Subject: Planning talk,  Meeting message,Fri 27/11/2026, 09:00 (1 hour) No Conflicts    Yesterday,        Message preview: Hello, ")!;
    expect(r.subject).toBe("Planning talk");
    expect(r.meeting).toBe(true);
    expect(r.date).toBe("Yesterday");
  });
  test("response: a single space before the marker does not end up in the subject", () => {
    const r = parseRowDesc("     Sender: Ann Lee, Subject: Accepted: Planning talk, Meeting message,    12/11/2026,        Message preview: , ")!;
    expect(r.subject).toBe("Accepted: Planning talk");
    expect(r.meeting).toBe(true);
    expect(r.date).toBe("12/11/2026");
  });
  test("subject that merely mentions a meeting message is kept", () => {
    const r = parseRowDesc("     Sender: A, Subject: Re: Meeting messages archive,      Today,        Message preview: x, ")!;
    expect(r.subject).toBe("Re: Meeting messages archive");
    expect(r.meeting).toBe(false);
  });
});

describe("matching a row to the reading pane", () => {
  test("list tags and prefixes shown only in the pane", () => {
    expect(sameSubject("[list] FW: Invitation: seminar", "Invitation: seminar")).toBe(true);
    expect(sameSubject("RE: [EXTERNAL] Re: Budget 2027", "Re: Budget 2027")).toBe(true);
    expect(sameSubject("Accepted: Planning talk", "Accepted: Planning talk")).toBe(true);
    expect(sameSubject("(No subject)", "")).toBe(true);
    expect(sameSubject("Budget 2028", "Budget 2027")).toBe(false);
    expect(sameSubject(null, "x")).toBe(false);
  });
  test("previews differing only in punctuation and spacing", () => {
    expect(samePreview("From: A <a@x.org> Sent: Thursday, July 2", "From: A <a@x.org>. Sent: ")).toBe(true);
    expect(samePreview("Something else entirely", "From: A")).toBe(false);
  });
});

describe("calendar dates", () => {
  test("week header", () => {
    const r = parseRange("November 2 - November 8, 2026")!;
    expect(ymd(r.start)).toBe("2026-11-02");
    expect(ymd(r.end)).toBe("2026-11-08");
  });
  test("week header across years", () => {
    const r = parseRange("December 28, 2026 - January 3, 2027")!;
    expect(ymd(r.start)).toBe("2026-12-28");
    expect(ymd(r.end)).toBe("2027-01-03");
  });
  test("week header across years without start year", () => {
    const r = parseRange("December 28 - January 3, 2027")!;
    expect(ymd(r.start)).toBe("2026-12-28");
  });
  test("long dates", () => {
    expect(ymd(parseLongDate("Thursday, 5 November 2026")!)).toBe("2026-11-05");
    expect(ymd(parseLongDate("Thursday, November 5, 2026")!)).toBe("2026-11-05");
    expect(ymd(parseLongDate("Tue, 03 Nov 2026")!)).toBe("2026-11-03");
  });
});
