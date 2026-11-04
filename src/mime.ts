// Pure-file side of message retrieval: parse .eml files written by Outlook's
// "Save As" and extract their attachments safely.
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, extname, relative, resolve, isAbsolute } from "node:path";
import PostalMime from "postal-mime";
import type { Address as PmAddress, Attachment as PmAttachment, Email } from "postal-mime";

export type Address = { name: string; address: string };

export type AttachmentMeta = {
  index: number;
  filename: string;
  mimeType: string;
  size: number;
  inline: boolean;
  contentId: string | null;
};

export type SavedAttachment = AttachmentMeta & { path: string };

export type ParsedMessage = {
  messageId: string | null;
  subject: string | null;
  date: string | null;
  from: Address | null;
  to: Address[];
  cc: Address[];
  bcc: Address[];
  replyTo: Address[];
  inReplyTo: string | null;
  references: string[];
  headers: Record<string, string>;
  text: string;
  html: string | null;
  attachments: AttachmentMeta[];
};

export type SaveOptions = { indexes?: number[]; names?: string[]; includeInline?: boolean };

const CURATED_HEADERS = [
  "list-id",
  "list-unsubscribe",
  "importance",
  "x-priority",
  "thread-topic",
  "thread-index",
  "content-type",
];

const MAX_NAME_CHARS = 200;
// APFS limits a file name to 255 UTF-8 bytes; keep headroom for " (n)".
const MAX_NAME_BYTES = 240;

// ---------------------------------------------------------------------------
// Parsing

async function parseRaw(path: string): Promise<Email> {
  const raw = await readFile(path);
  return PostalMime.parse(raw, { attachmentEncoding: "arraybuffer" });
}

export async function parseEml(path: string): Promise<ParsedMessage> {
  const email = await parseRaw(path);

  const headers: Record<string, string> = {};
  for (const h of email.headers) {
    if (CURATED_HEADERS.includes(h.key) && !(h.key in headers)) headers[h.key] = h.value;
  }

  const html = email.html ?? null;
  let text = email.text ?? "";
  if (!text.trim() && html) text = htmlToText(html);

  return {
    messageId: email.messageId ?? null,
    subject: email.subject ?? null,
    date: toIsoDate(email.date),
    from: email.from ? (flattenAddresses([email.from])[0] ?? null) : null,
    to: flattenAddresses(email.to),
    cc: flattenAddresses(email.cc),
    bcc: flattenAddresses(email.bcc),
    replyTo: flattenAddresses(email.replyTo),
    inReplyTo: email.inReplyTo ?? null,
    references: splitReferences(email.references),
    headers,
    text,
    html,
    attachments: email.attachments.map(toMeta),
  };
}

function toIsoDate(value: string | undefined): string | null {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

function flattenAddresses(list: PmAddress[] | undefined): Address[] {
  const out: Address[] = [];
  for (const a of list ?? []) {
    if (a.group) {
      for (const m of a.group) out.push({ name: m.name ?? "", address: m.address ?? "" });
    } else {
      out.push({ name: a.name ?? "", address: a.address ?? "" });
    }
  }
  return out;
}

function splitReferences(value: string | undefined): string[] {
  if (!value) return [];
  const ids = value.match(/<[^<>\s]+>/g);
  return ids ?? value.split(/\s+/).filter(Boolean);
}

function contentBytes(att: PmAttachment): Uint8Array {
  const c = att.content;
  if (typeof c === "string") {
    return att.encoding === "base64" ? Buffer.from(c, "base64") : new TextEncoder().encode(c);
  }
  return c instanceof Uint8Array ? c : new Uint8Array(c);
}

function toMeta(att: PmAttachment, index: number): AttachmentMeta {
  const mimeType = att.mimeType || "application/octet-stream";
  const filename = att.filename?.trim() || `attachment-${index}.${extensionFor(mimeType)}`;
  return {
    index,
    filename,
    mimeType,
    size: contentBytes(att).byteLength,
    inline: att.disposition === "inline" || Boolean(att.contentId),
    contentId: att.contentId ?? null,
  };
}

const EXTENSIONS: Record<string, string> = {
  "message/rfc822": "eml",
  "application/pdf": "pdf",
  "application/zip": "zip",
  "application/json": "json",
  "application/msword": "doc",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  "application/vnd.ms-excel": "xls",
  "application/vnd.ms-powerpoint": "ppt",
  "application/octet-stream": "bin",
  "image/jpeg": "jpg",
  "image/png": "png",
  "image/gif": "gif",
  "image/svg+xml": "svg",
  "image/webp": "webp",
  "text/plain": "txt",
  "text/html": "html",
  "text/calendar": "ics",
  "text/csv": "csv",
};

function extensionFor(mimeType: string): string {
  const known = EXTENSIONS[mimeType.toLowerCase()];
  if (known) return known;
  // Fall back to the subtype ("image/heic" -> "heic") when it is a plain token.
  const sub = mimeType.split("/")[1]?.toLowerCase() ?? "";
  return /^[a-z0-9]{1,10}$/.test(sub) ? sub : "bin";
}

// ---------------------------------------------------------------------------
// Saving attachments

export async function saveAttachments(
  path: string,
  outDir: string,
  opts: SaveOptions = {},
): Promise<SavedAttachment[]> {
  const email = await parseRaw(path);
  const metas = email.attachments.map(toMeta);

  const explicit = opts.indexes !== undefined || opts.names !== undefined;
  const indexes = new Set(opts.indexes ?? []);
  const names = new Set((opts.names ?? []).map((n) => n.toLowerCase()));
  const selected = metas.filter((m) =>
    explicit
      ? indexes.has(m.index) || names.has(m.filename.toLowerCase())
      : opts.includeInline || !m.inline,
  );

  const root = resolve(outDir);
  await mkdir(root, { recursive: true });

  const saved: SavedAttachment[] = [];
  for (const meta of selected) {
    const att = email.attachments[meta.index]!;
    const safe = sanitizeFilename(meta.filename, `attachment-${meta.index}.${extensionFor(meta.mimeType)}`);
    const target = await writeUnique(root, safe, contentBytes(att));
    saved.push({ ...meta, path: target });
  }
  return saved;
}

/** Make an attachment name safe to use as a single path component. */
export function sanitizeFilename(name: string, fallback = "attachment.bin"): string {
  // Keep only the last component, whatever separator the sender used.
  let s = name.split(/[\\/]/).pop() ?? "";
  s = s.replace(/[\u0000-\u001f\u007f]/g, "").replace(/:/g, "_");
  s = s.replace(/^[\s.]+/, "").replace(/[\s.]+$/, "");
  if (!s) s = fallback;
  return truncateName(s);
}

function truncateName(name: string): string {
  const fits = (s: string) =>
    [...s].length <= MAX_NAME_CHARS && Buffer.byteLength(s, "utf8") <= MAX_NAME_BYTES;
  if (fits(name)) return name;

  let ext = extname(name);
  if ([...ext].length > 20) ext = "";
  const stem = [...name.slice(0, name.length - ext.length)];
  while (stem.length > 1 && !fits(stem.join("") + ext)) stem.pop();
  return stem.join("") + ext;
}

/** Write without ever overwriting: "a.pdf", then "a (1).pdf", "a (2).pdf", ... */
async function writeUnique(root: string, name: string, data: Uint8Array): Promise<string> {
  const ext = extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let n = 0; n < 10_000; n++) {
    const candidate = n === 0 ? name : `${stem} (${n})${ext}`;
    const target = assertInside(root, resolve(root, candidate));
    try {
      // "wx" fails with EEXIST instead of overwriting, so there is no check-then-write race.
      await writeFile(target, data, { flag: "wx" });
      return target;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err;
    }
  }
  throw new Error(`Could not find a free file name for ${name} in ${root}`);
}

function assertInside(root: string, target: string): string {
  const rel = relative(root, target);
  if (!rel || rel.startsWith("..") || isAbsolute(rel) || basename(target) !== rel) {
    throw new Error(`Refusing to write outside ${root}: ${target}`);
  }
  return target;
}

// ---------------------------------------------------------------------------
// HTML to text

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ",
  ndash: "–", mdash: "—", hellip: "…", bull: "•", middot: "·",
  lsquo: "‘", rsquo: "’", ldquo: "“", rdquo: "”", laquo: "«", raquo: "»",
  copy: "©", reg: "®", trade: "™", euro: "€", pound: "£", yen: "¥",
  deg: "°", times: "×", divide: "÷", shy: "", zwnj: "", zwj: "",
  eacute: "é", egrave: "è", ecirc: "ê", euml: "ë", aacute: "á", agrave: "à",
  acirc: "â", auml: "ä", ccedil: "ç", iuml: "ï", icirc: "î", ocirc: "ô",
  ouml: "ö", uuml: "ü", ucirc: "û", ugrave: "ù", szlig: "ß",
  Eacute: "É", Agrave: "À", Ccedil: "Ç",
};

function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z][a-z0-9]*);/gi, (match, body: string) => {
    if (body[0] === "#") {
      const code = body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
    }
    return NAMED_ENTITIES[body] ?? NAMED_ENTITIES[body.toLowerCase()] ?? match;
  });
}

export function htmlToText(html: string): string {
  let s = html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<(script|style|head|title)\b[\s\S]*?<\/\1\s*>/gi, "")
    .replace(/\r\n?/g, "\n");

  // Source whitespace is not significant in HTML; only tags produce line breaks.
  s = s.replace(/\s+/g, " ");
  s = s
    .replace(/<br\b[^>]*>/gi, "\n")
    .replace(/<li\b[^>]*>/gi, "\n- ")
    .replace(/<\/(p|div|tr|li|h[1-6]|blockquote|table|ul|ol|pre)\s*>/gi, "\n")
    .replace(/<(p|div|tr|h[1-6]|blockquote|table|ul|ol|pre)\b[^>]*>/gi, "\n")
    .replace(/<\/t[dh]\s*>/gi, "\t")
    .replace(/<hr\b[^>]*>/gi, "\n----\n")
    .replace(/<[^>]+>/g, "");

  s = decodeEntities(s).replace(/ /g, " ");

  return s
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
