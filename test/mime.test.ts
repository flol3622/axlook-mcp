import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { htmlToText, parseEml, sanitizeFilename, saveAttachments } from "../src/mime.ts";

let dir: string;

const crlf = (s: string) => s.replace(/\r?\n/g, "\r\n");
async function fixture(name: string, body: string): Promise<string> {
  const p = join(dir, name);
  await writeFile(p, crlf(body));
  return p;
}

const PDF_BYTES = "%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\ntrailer\n%%EOF\n";
const PNG_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

const ALTERNATIVE = `From: "Dupont, Jean" <jean@example.org>
To: Team: a@example.org, "B Person" <b@example.org>;, c@example.org
Cc: d@example.org
Reply-To: replies@example.org
Subject: =?Windows-1252?Q?Conf_2027_proceedings_=96_indexing?=
Date: Tue, 3 Nov 2026 10:15:00 +0100
Message-ID: <msg-1@example.org>
In-Reply-To: <parent@example.org>
References: <root@example.org>
 <parent@example.org>
List-Id: Conference list <conf.example.org>
Importance: High
Thread-Topic: Conf 2027 proceedings
X-Unrelated: nope
MIME-Version: 1.0
Content-Type: multipart/alternative; boundary="ALT"

--ALT
Content-Type: text/plain; charset=windows-1252
Content-Transfer-Encoding: quoted-printable

Dear authors =96 the proceedings are indexed.
--ALT
Content-Type: text/html; charset=utf-8

<p>Dear authors &ndash; the proceedings are indexed.</p>
--ALT--
`;

const MIXED = `From: sender@example.org
To: you@example.org
Subject: Files
Date: Tue, 3 Nov 2026 11:00:00 +0000
Message-ID: <msg-2@example.org>
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="MIX"

--MIX
Content-Type: multipart/related; boundary="REL"

--REL
Content-Type: text/html; charset=utf-8

<p>See <img src="cid:logo@x"></p>
--REL
Content-Type: image/png
Content-ID: <logo@x>
Content-Disposition: inline; filename="logo.png"
Content-Transfer-Encoding: base64

${PNG_BYTES.toString("base64")}
--REL--
--MIX
Content-Type: application/pdf; name="report.pdf"
Content-Disposition: attachment; filename="report.pdf"
Content-Transfer-Encoding: base64

${Buffer.from(PDF_BYTES).toString("base64")}
--MIX
Content-Type: text/plain; name="../../evil.txt"
Content-Disposition: attachment; filename="../../evil.txt"

evil
--MIX
Content-Type: application/pdf
Content-Disposition: attachment
Content-Transfer-Encoding: base64

${Buffer.from(PDF_BYTES).toString("base64")}
--MIX
Content-Type: message/rfc822
Content-Disposition: attachment

From: inner@example.org
Subject: Inner message

Inner body.
--MIX--
`;

const HTML_ONLY = `From: news@example.org
To: you@example.org
Subject: Newsletter
Date: not a date
MIME-Version: 1.0
Content-Type: text/html; charset=utf-8

<html><head><style>p { color: red }</style><title>T</title></head>
<body><script>alert("x")</script>
<p>Hello&nbsp;&amp; welcome</p><p>Line one<br>Line two</p>
<ul><li>First</li><li>Second &#8211; dash</li></ul>
<table><tr><td>A</td><td>B</td></tr></table>
</body></html>
`;

let alternative: string, mixed: string, htmlOnly: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "mime-test-"));
  alternative = await fixture("alt.eml", ALTERNATIVE);
  mixed = await fixture("mixed.eml", MIXED);
  htmlOnly = await fixture("html.eml", HTML_ONLY);
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe("parseEml", () => {
  test("decodes Windows-1252 encoded subject and body", async () => {
    const m = await parseEml(alternative);
    expect(m.subject).toBe("Conf 2027 proceedings – indexing");
    expect(m.text).toContain("Dear authors – the proceedings");
    expect(m.html).toContain("<p>");
  });

  test("extracts addresses, ids, date and curated headers", async () => {
    const m = await parseEml(alternative);
    expect(m.from).toEqual({ name: "Dupont, Jean", address: "jean@example.org" });
    expect(m.to.map((a) => a.address)).toEqual(["a@example.org", "b@example.org", "c@example.org"]);
    expect(m.cc).toEqual([{ name: "", address: "d@example.org" }]);
    expect(m.replyTo[0]?.address).toBe("replies@example.org");
    expect(m.bcc).toEqual([]);
    expect(m.messageId).toBe("<msg-1@example.org>");
    expect(m.inReplyTo).toBe("<parent@example.org>");
    expect(m.references).toEqual(["<root@example.org>", "<parent@example.org>"]);
    expect(m.date).toBe("2026-11-03T09:15:00.000Z");
    expect(m.headers["list-id"]).toContain("conf.example.org");
    expect(m.headers["importance"]).toBe("High");
    expect(m.headers["thread-topic"]).toBe("Conf 2027 proceedings");
    expect(m.headers["content-type"]).toContain("multipart/alternative");
    expect(m.headers["x-unrelated"]).toBeUndefined();
    expect(m.attachments).toEqual([]);
  });

  test("lists attachments with inline flags and synthesized names", async () => {
    const m = await parseEml(mixed);
    const byName = Object.fromEntries(m.attachments.map((a) => [a.filename, a]));
    expect(byName["logo.png"]).toMatchObject({ inline: true, contentId: "<logo@x>", mimeType: "image/png", size: PNG_BYTES.length });
    expect(byName["report.pdf"]).toMatchObject({ inline: false, contentId: null, size: PDF_BYTES.length });
    expect(byName["../../evil.txt"]).toBeDefined();
    const unnamed = m.attachments.find((a) => a.filename.startsWith("attachment-") && a.mimeType === "application/pdf");
    expect(unnamed?.filename).toBe(`attachment-${unnamed?.index}.pdf`);
    const nested = m.attachments.find((a) => a.mimeType === "message/rfc822");
    expect(nested?.filename).toBe(`attachment-${nested?.index}.eml`);
    m.attachments.forEach((a, i) => expect(a.index).toBe(i));
  });

  test("derives text from HTML-only messages", async () => {
    const m = await parseEml(htmlOnly);
    expect(m.date).toBeNull();
    expect(m.text).not.toContain("alert");
    expect(m.text).not.toContain("color");
    expect(m.text).toContain("Hello & welcome");
    expect(m.text).toContain("Line one\nLine two");
    expect(m.text).toContain("- Second – dash");
    expect(m.text).not.toMatch(/\n{3,}/);
  });
});

describe("htmlToText", () => {
  test("handles entities, breaks and blank lines", () => {
    expect(htmlToText("<div>a&lt;b &#x2014; &quot;c&quot;</div>\n\n\n<p></p><p>d</p>")).toBe('a<b — "c"\n\nd');
  });
});

describe("sanitizeFilename", () => {
  test("strips directories, control chars and leading dots", () => {
    expect(sanitizeFilename("../../evil.txt")).toBe("evil.txt");
    expect(sanitizeFilename("..\\..\\win.txt")).toBe("win.txt");
    expect(sanitizeFilename(".hidden\u0000x\u001f.sh")).toBe("hiddenx.sh");
    expect(sanitizeFilename("...", "fb.bin")).toBe("fb.bin");
  });

  test("caps length and keeps the extension", () => {
    const s = sanitizeFilename("a".repeat(500) + ".pdf");
    expect(s.length).toBe(200);
    expect(s.endsWith(".pdf")).toBe(true);
  });
});

describe("saveAttachments", () => {
  test("saves non-inline attachments by default, safely", async () => {
    const out = join(dir, "out-default", "nested");
    const saved = await saveAttachments(mixed, out);
    const files = (await readdir(out)).sort();
    expect(saved.every((s) => !s.inline)).toBe(true);
    expect(files).toContain("report.pdf");
    expect(files).toContain("evil.txt");
    expect(files.some((f) => f.endsWith(".eml"))).toBe(true);
    expect(files).not.toContain("logo.png");
    for (const s of saved) expect(s.path.startsWith(out + "/")).toBe(true);
    const pdf = saved.find((s) => s.filename === "report.pdf")!;
    expect(await readFile(pdf.path, "utf8")).toBe(PDF_BYTES);
    const eml = saved.find((s) => s.mimeType === "message/rfc822")!;
    expect(await readFile(eml.path, "utf8")).toContain("Subject: Inner message");
    // Nothing escaped the temp dir.
    expect(await Bun.file(join(dir, "evil.txt")).exists()).toBe(false);
  });

  test("includeInline adds inline parts", async () => {
    const out = join(dir, "out-inline");
    const saved = await saveAttachments(mixed, out, { includeInline: true });
    expect(saved.map((s) => s.filename)).toContain("logo.png");
    expect(new Uint8Array(await readFile(join(out, "logo.png")))).toEqual(new Uint8Array(PNG_BYTES));
  });

  test("selects exactly by index or case-insensitive name", async () => {
    const out = join(dir, "out-select");
    const saved = await saveAttachments(mixed, out, { names: ["LOGO.PNG"], indexes: [] });
    expect(saved.map((s) => s.filename)).toEqual(["logo.png"]);
    const byIndex = await saveAttachments(mixed, join(dir, "out-index"), { indexes: [saved[0]!.index] });
    expect(byIndex.map((s) => s.filename)).toEqual(["logo.png"]);
  });

  test("never overwrites: second save gets ' (1)'", async () => {
    const out = join(dir, "out-dup");
    const first = await saveAttachments(mixed, out, { names: ["report.pdf"] });
    const second = await saveAttachments(mixed, out, { names: ["report.pdf"] });
    const third = await saveAttachments(mixed, out, { names: ["report.pdf"] });
    expect(first[0]!.path).toBe(join(out, "report.pdf"));
    expect(second[0]!.path).toBe(join(out, "report (1).pdf"));
    expect(third[0]!.path).toBe(join(out, "report (2).pdf"));
  });
});
