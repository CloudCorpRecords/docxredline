/**
 * docxredline core: read, redline, and finalize .docx files, locally.
 *
 * A .docx is a ZIP of XML parts. This module carries its own minimal ZIP
 * reader (stored + deflated entries, inflated with node:zlib) and ZIP
 * writer (stored entries with locally computed CRC32), then works on
 * word/document.xml as text:
 *
 * - Reading splits the body into <w:p> paragraphs, concatenates the
 *   <w:t> text of each, and renders existing <w:ins> / <w:del> tracked
 *   changes as CriticMarkup ({++ ++} / {-- --}). The clean view resolves
 *   them instead: inserted text stays, deleted text goes.
 * - A modify edit finds its target inside one paragraph's concatenated
 *   <w:t> text and rewrites the containing run, wrapping the old text in
 *   a <w:del> and the new text in a <w:ins> so Word shows a native
 *   tracked change. Safety gate, mirroring the original tool: a
 *   non-regex target with match_mode "first" that matches more than one
 *   paragraph is refused as ambiguous rather than guessed at.
 * - Accept / reject address one tracked-change element by its w:id
 *   (carried in the change's target_text): accept unwraps <w:ins> and
 *   drops <w:del>; reject does the inverse.
 * - finalize empties the author fields of docProps/core.xml, strips
 *   w:rsid* attributes from the body, and can add a readOnly
 *   documentProtection element to word/settings.xml.
 *
 * Paragraphs are paged synthetically at 40 per page for read_docx's
 * `page` argument; .docx has no page concept of its own.
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { inflateRawSync } from "node:zlib";
import type {
  BatchInput,
  BatchResult,
  DocChange,
  FinalizeInput,
  FinalizeResult,
  ReadDocxOptions,
  ReadDocxResult,
  ZipEntry,
} from "./types.js";

export const VERSION = "0.1.0";

/** Paragraphs per synthetic page for read_docx's `page` argument. */
export const PARAGRAPHS_PER_PAGE = 40;

// ---------------------------------------------------------------------------
// CRC32 + ZIP (stored writer, stored/deflated reader)
// ---------------------------------------------------------------------------

const CRC_TABLE: number[] = (() => {
  const table: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table.push(c >>> 0);
  }
  return table;
})();

/** Standard CRC-32 (IEEE 802.3 polynomial) of a byte string. */
export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16le(n: number): number[] {
  return [n & 0xff, (n >>> 8) & 0xff];
}

function u32le(n: number): number[] {
  return [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff];
}

// DOS date pinned to 2026-10-26 so written archives are deterministic.
const DOS_TIME = 0;
const DOS_DATE = ((2026 - 1980) << 9) | (10 << 5) | 26;

/** Build a ZIP archive of stored (uncompressed) entries. */
export function writeZip(entries: ZipEntry[]): Buffer {
  const encoder = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  let offset = 0;
  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const crc = crc32(entry.data);
    const localHeader = [
      ...u32le(0x04034b50), // local file header signature
      ...u16le(20), // version needed
      ...u16le(0), // flags
      ...u16le(0), // method 0 = stored
      ...u16le(DOS_TIME),
      ...u16le(DOS_DATE),
      ...u32le(crc),
      ...u32le(entry.data.length), // compressed size
      ...u32le(entry.data.length), // uncompressed size
      ...u16le(nameBytes.length),
      ...u16le(0), // extra length
    ];
    chunks.push(...localHeader, ...nameBytes, ...entry.data);
    central.push(
      ...u32le(0x02014b50), // central directory signature
      ...u16le(20), // version made by
      ...u16le(20), // version needed
      ...u16le(0), // flags
      ...u16le(0), // method: stored
      ...u16le(DOS_TIME),
      ...u16le(DOS_DATE),
      ...u32le(crc),
      ...u32le(entry.data.length),
      ...u32le(entry.data.length),
      ...u16le(nameBytes.length),
      ...u16le(0), // extra length
      ...u16le(0), // comment length
      ...u16le(0), // disk number start
      ...u16le(0), // internal attributes
      ...u32le(0), // external attributes
      ...u32le(offset), // local header offset
      ...nameBytes,
    );
    offset += localHeader.length + nameBytes.length + entry.data.length;
  }
  const endRecord = [
    ...u32le(0x06054b50), // end of central directory signature
    ...u16le(0), // disk number
    ...u16le(0), // disk with central directory
    ...u16le(entries.length), // entries on this disk
    ...u16le(entries.length), // total entries
    ...u32le(central.length), // central directory size
    ...u32le(offset), // central directory offset
    ...u16le(0), // comment length
  ];
  return Buffer.from([...chunks, ...central, ...endRecord]);
}

/** Read a ZIP archive (stored or deflated entries) into ordered entries. */
export function readZip(buf: Buffer): ZipEntry[] {
  let eocd = -1;
  const earliest = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= earliest; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) {
    throw new Error("not a ZIP archive (end of central directory not found)");
  }
  const count = buf.readUInt16LE(eocd + 10);
  let pos = buf.readUInt32LE(eocd + 16);
  const entries: ZipEntry[] = [];
  for (let n = 0; n < count; n++) {
    if (pos + 46 > buf.length || buf.readUInt32LE(pos) !== 0x02014b50) {
      throw new Error("corrupt ZIP archive (bad central directory entry)");
    }
    const method = buf.readUInt16LE(pos + 10);
    const compSize = buf.readUInt32LE(pos + 20);
    const nameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    const localOffset = buf.readUInt32LE(pos + 42);
    const name = buf.toString("utf8", pos + 46, pos + 46 + nameLen);
    if (localOffset + 30 > buf.length || buf.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`corrupt ZIP archive (bad local header for ${name})`);
    }
    const lNameLen = buf.readUInt16LE(localOffset + 26);
    const lExtraLen = buf.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + lNameLen + lExtraLen;
    const raw = buf.subarray(dataStart, dataStart + compSize);
    let data: Buffer;
    if (method === 0) {
      data = Buffer.from(raw);
    } else if (method === 8) {
      data = inflateRawSync(raw);
    } else {
      throw new Error(`unsupported ZIP compression method ${method} for ${name}`);
    }
    entries.push({ name, data });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

const XML_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

/** Decode XML entities (&amp; &lt; &#65; &#x41; ...) in a text node. */
export function decodeXml(text: string): string {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body: string) => {
    if (body.startsWith("#x") || body.startsWith("#X")) {
      return String.fromCodePoint(parseInt(body.slice(2), 16));
    }
    if (body.startsWith("#")) {
      return String.fromCodePoint(parseInt(body.slice(1), 10));
    }
    return XML_ENTITIES[body] ?? whole;
  });
}

/** Escape a string for use as XML text content. */
export function encodeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Escape a string for use inside a double-quoted XML attribute. */
function encodeAttr(text: string): string {
  return encodeXml(text).replace(/"/g, "&quot;");
}

/** Paragraph XML spans of a document.xml string, in order. */
export function splitParagraphs(documentXml: string): string[] {
  const spans: string[] = [];
  const re = /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(documentXml)) !== null) spans.push(m[0]);
  return spans;
}

/** Replace the i-th paragraph span of document.xml (paragraph count is stable). */
function replaceParagraph(documentXml: string, index: number, replacement: string): string {
  const re = /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let i = 0;
  return documentXml.replace(re, (span) => (i++ === index ? replacement : span));
}

type SegmentKind = "normal" | "ins" | "del";
interface Segment {
  kind: SegmentKind;
  text: string;
}

/**
 * The text segments of one paragraph, in order, each tagged by whether it
 * sits inside a <w:ins>, inside a <w:del> (its text lives in <w:delText>),
 * or in ordinary body text.
 */
function paragraphSegments(paragraphXml: string): Segment[] {
  const segments: Segment[] = [];
  const tokenRe = /<[^>]+>|[^<]+/g;
  const context: SegmentKind[] = [];
  let capture: "t" | "delText" | null = null;
  let buffer = "";
  const kind = (): SegmentKind =>
    context.includes("del") ? "del" : context.includes("ins") ? "ins" : "normal";
  const flush = () => {
    if (capture !== null && buffer.length > 0) {
      segments.push({ kind: capture === "delText" ? "del" : kind(), text: decodeXml(buffer) });
    }
    buffer = "";
  };
  let tok: RegExpExecArray | null;
  while ((tok = tokenRe.exec(paragraphXml)) !== null) {
    const t = tok[0];
    if (t.startsWith("<")) {
      if (/^<w:ins\b/.test(t) && !t.endsWith("/>")) context.push("ins");
      else if (/^<\/w:ins\s*>/.test(t)) {
        const i = context.lastIndexOf("ins");
        if (i !== -1) context.splice(i, 1);
      } else if (/^<w:del\b/.test(t) && !t.endsWith("/>")) context.push("del");
      else if (/^<\/w:del\s*>/.test(t)) {
        const i = context.lastIndexOf("del");
        if (i !== -1) context.splice(i, 1);
      } else if (/^<w:t\b/.test(t) && !t.endsWith("/>")) {
        flush();
        capture = "t";
      } else if (/^<w:delText\b/.test(t) && !t.endsWith("/>")) {
        flush();
        capture = "delText";
      } else if (/^<\/w:(t|delText)\s*>/.test(t)) {
        flush();
        capture = null;
      }
    } else if (capture !== null) {
      buffer += t;
    }
  }
  flush();
  return segments;
}

/** Concatenated <w:t> text of a paragraph (inserted text in, deleted text out). */
function paragraphPlainText(paragraphXml: string): string {
  return paragraphSegments(paragraphXml)
    .filter((s) => s.kind !== "del")
    .map((s) => s.text)
    .join("");
}

/** Render one paragraph in markup (CriticMarkup) or clean view. */
function renderParagraph(paragraphXml: string, cleanView: boolean): string {
  const out: string[] = [];
  for (const seg of paragraphSegments(paragraphXml)) {
    if (seg.kind === "ins") out.push(cleanView ? seg.text : `{++${seg.text}++}`);
    else if (seg.kind === "del") {
      if (!cleanView) out.push(`{--${seg.text}--}`);
    } else out.push(seg.text);
  }
  return out.join("");
}

// ---------------------------------------------------------------------------
// .docx loading
// ---------------------------------------------------------------------------

interface LoadedDocx {
  resolvedPath: string;
  entries: ZipEntry[];
  documentXml: string;
}

function loadDocx(inputPath: string): LoadedDocx {
  const resolvedPath = path.resolve(inputPath);
  const buf = readFileSync(resolvedPath); // ENOENT etc. propagate as the real error
  const entries = readZip(buf);
  const doc = entries.find((e) => e.name === "word/document.xml");
  if (!doc) {
    throw new Error("not a .docx file (word/document.xml missing)");
  }
  return { resolvedPath, entries, documentXml: doc.data.toString("utf8") };
}

function saveDocx(entries: ZipEntry[], documentXml: string, outputPath: string): void {
  const next = entries.map((e) =>
    e.name === "word/document.xml" ? { name: e.name, data: Buffer.from(documentXml, "utf8") } : e,
  );
  writeFileSync(outputPath, writeZip(next));
}

function siblingOutput(inputPath: string, suffix: string): string {
  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath).replace(/\.docx$/i, "");
  return path.join(dir, `${base}.${suffix}.docx`);
}

// ---------------------------------------------------------------------------
// read_docx
// ---------------------------------------------------------------------------

export function readDocx(inputPath: string, options: ReadDocxOptions = {}): ReadDocxResult {
  const { resolvedPath, documentXml } = loadDocx(inputPath);
  const cleanView = options.cleanView ?? false;
  let paragraphs = splitParagraphs(documentXml).map((p) => renderParagraph(p, cleanView));
  const total = paragraphs.length;
  if (options.searchQuery !== undefined) {
    const q = options.searchQuery;
    paragraphs = paragraphs.filter((p) => p.includes(q));
  }
  const pageCount = Math.max(1, Math.ceil(total / PARAGRAPHS_PER_PAGE));
  const result: ReadDocxResult = {
    path: resolvedPath,
    text: "",
    paragraphs,
    paragraph_count: paragraphs.length,
    total_paragraphs: total,
    page_count: pageCount,
  };
  if (options.page !== undefined) {
    const start = (options.page - 1) * PARAGRAPHS_PER_PAGE;
    result.paragraphs = paragraphs.slice(start, start + PARAGRAPHS_PER_PAGE);
    result.paragraph_count = result.paragraphs.length;
    result.page = options.page;
  }
  result.text = result.paragraphs.join("\n\n");
  return result;
}

// ---------------------------------------------------------------------------
// process_document_batch
// ---------------------------------------------------------------------------

interface TextRunSpan {
  /** Char offset of this <w:t> element within the paragraph XML. */
  index: number;
  /** The whole element, e.g. <w:t xml:space="preserve">abc</w:t>. */
  element: string;
  /** The element's opening tag. */
  openTag: string;
  /** Decoded text of the element. */
  text: string;
}

function textRunSpans(paragraphXml: string): TextRunSpan[] {
  const spans: TextRunSpan[] = [];
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(paragraphXml)) !== null) {
    const openTag = m[0].slice(0, m[0].indexOf(">") + 1);
    spans.push({ index: m.index, element: m[0], openTag, text: decodeXml(m[1]) });
  }
  return spans;
}

interface MatchRange {
  start: number;
  end: number;
  matched: string;
}

function findMatches(plain: string, change: DocChange): MatchRange[] {
  const target = change.target_text ?? "";
  if (change.regex) {
    const re = new RegExp(target, "g");
    const out: MatchRange[] = [];
    let m: RegExpExecArray | null;
    while ((m = re.exec(plain)) !== null) {
      out.push({ start: m.index, end: m.index + m[0].length, matched: m[0] });
      if (m[0].length === 0) re.lastIndex++;
    }
    return out;
  }
  const out: MatchRange[] = [];
  let from = 0;
  for (;;) {
    const i = plain.indexOf(target, from);
    if (i === -1) break;
    out.push({ start: i, end: i + target.length, matched: target });
    from = i + Math.max(1, target.length);
  }
  return out;
}

/**
 * Rewrite one paragraph so the plain-text range [start, end) becomes a
 * tracked change: the old text wrapped in <w:del>, the new text in a
 * following <w:ins>. The range must lie inside a single <w:t> element.
 */
function applyModifyToParagraph(
  paragraphXml: string,
  range: MatchRange,
  newText: string,
  author: string,
  id: number,
  date: string,
): string {
  const spans = textRunSpans(paragraphXml);
  let plainStart = 0;
  for (const span of spans) {
    const spanEnd = plainStart + span.text.length;
    if (range.start >= plainStart && range.end <= spanEnd) {
      const localStart = range.start - plainStart;
      const localEnd = range.end - plainStart;
      const prefix = span.text.slice(0, localStart);
      const old = span.text.slice(localStart, localEnd);
      const suffix = span.text.slice(localEnd);
      const replacement =
        span.openTag +
        encodeXml(prefix) +
        `</w:t></w:r>` +
        `<w:del w:id="${id}" w:author="${encodeAttr(author)}" w:date="${date}">` +
        `<w:r><w:delText xml:space="preserve">${encodeXml(old)}</w:delText></w:r></w:del>` +
        `<w:ins w:id="${id + 1}" w:author="${encodeAttr(author)}" w:date="${date}">` +
        `<w:r><w:t xml:space="preserve">${encodeXml(newText)}</w:t></w:r></w:ins>` +
        `<w:r><w:t xml:space="preserve">${encodeXml(suffix)}</w:t>`;
      return (
        paragraphXml.slice(0, span.index) +
        replacement +
        paragraphXml.slice(span.index + span.element.length)
      );
    }
    plainStart = spanEnd;
  }
  throw new Error("target spans multiple runs; not supported in this subset");
}

function nextTrackedId(documentXml: string): number {
  let max = 0;
  const re = /w:id="(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(documentXml)) !== null) max = Math.max(max, parseInt(m[1], 10));
  return max + 1;
}

function attrValue(attrs: string, name: string): string | null {
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
}

/**
 * Accept or reject the tracked-change element carrying w:id === id.
 * Accept: unwrap <w:ins> (content stays), drop <w:del> (content goes).
 * Reject: drop <w:ins>, unwrap <w:del> restoring its text as <w:t>.
 * Returns the new XML, or null when no element carries that id.
 */
function applyAcceptReject(documentXml: string, id: string, accept: boolean): string | null {
  const insRe = /<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>/g;
  let m: RegExpExecArray | null;
  while ((m = insRe.exec(documentXml)) !== null) {
    if (attrValue(m[1], "w:id") === id) {
      const replacement = accept ? m[2] : "";
      return documentXml.slice(0, m.index) + replacement + documentXml.slice(m.index + m[0].length);
    }
  }
  const delRe = /<w:del\b([^>]*)>([\s\S]*?)<\/w:del>/g;
  while ((m = delRe.exec(documentXml)) !== null) {
    if (attrValue(m[1], "w:id") === id) {
      const restored = m[2]
        .replace(/<w:delText\b([^>]*)>/g, "<w:t$1>")
        .replace(/<\/w:delText\s*>/g, "</w:t>");
      const replacement = accept ? "" : restored;
      return documentXml.slice(0, m.index) + replacement + documentXml.slice(m.index + m[0].length);
    }
  }
  return null;
}

export function processDocumentBatch(input: BatchInput): BatchResult {
  const loaded = loadDocx(input.path);
  let documentXml = loaded.documentXml;
  const author = input.author ?? "docxredline";
  const failed: { index: number; error: string }[] = [];
  let applied = 0;
  let idCounter = nextTrackedId(documentXml);

  input.changes.forEach((change, index) => {
    try {
      if (change.type === "modify") {
        if (!change.target_text) throw new Error("modify requires target_text");
        const mode = change.match_mode ?? "first";
        if (mode !== "first" && mode !== "all") {
          throw new Error(`unknown match_mode: ${mode}`);
        }
        const paragraphs = splitParagraphs(documentXml);
        const plains = paragraphs.map(paragraphPlainText);
        const matchesPerPara = plains.map((p) => findMatches(p, change));
        const matchingParas = matchesPerPara
          .map((matches, i) => (matches.length > 0 ? i : -1))
          .filter((i) => i !== -1);
        if (matchingParas.length === 0) {
          throw new Error(`target not found: ${JSON.stringify(change.target_text)}`);
        }
        // Safety gate: an unanchored "first" edit must never guess between
        // paragraphs. (Regex targets and match_mode "all" are exempt: the
        // caller has opted into the broader behaviour explicitly.)
        if (!change.regex && mode === "first" && matchingParas.length > 1) {
          throw new Error(
            `ambiguous target: ${JSON.stringify(change.target_text)} matches ` +
              `${matchingParas.length} paragraphs; refusing a "first" edit ` +
              `(use match_mode "all" or a longer target)`,
          );
        }
        const date = new Date().toISOString();
        const newText = change.new_text ?? "";
        const targets =
          mode === "all"
            ? matchingParas.flatMap((pi) =>
                matchesPerPara[pi].map((range) => ({ pi, range })),
              )
            : [{ pi: matchingParas[0], range: matchesPerPara[matchingParas[0]][0] }];
        // Apply in reverse document order so earlier plain-text offsets
        // stay valid while the paragraph XML grows around later edits.
        targets.sort((a, b) => b.pi - a.pi || b.range.start - a.range.start);
        for (const { pi, range } of targets) {
          const current = splitParagraphs(documentXml)[pi];
          const updated = applyModifyToParagraph(current, range, newText, author, idCounter, date);
          idCounter += 2;
          documentXml = replaceParagraph(documentXml, pi, updated);
        }
        applied++;
      } else if (change.type === "accept" || change.type === "reject") {
        if (!change.target_text) {
          throw new Error(`${change.type} requires target_text (the w:id of the tracked change)`);
        }
        const updated = applyAcceptReject(documentXml, change.target_text, change.type === "accept");
        if (updated === null) {
          throw new Error(`no tracked change with w:id ${JSON.stringify(change.target_text)}`);
        }
        documentXml = updated;
        applied++;
      } else if (change.type === "reply") {
        throw new Error("comments not supported in this subset");
      } else {
        throw new Error(`unknown change type: ${JSON.stringify(change.type)}`);
      }
    } catch (err) {
      failed.push({ index, error: err instanceof Error ? err.message : String(err) });
    }
  });

  const result: BatchResult = { applied, failed };
  if (applied >= 1) {
    const outputPath = path.resolve(input.outputPath ?? siblingOutput(loaded.resolvedPath, "redlined"));
    saveDocx(loaded.entries, documentXml, outputPath);
    result.output_path = outputPath;
  }
  return result;
}

// ---------------------------------------------------------------------------
// finalize_document
// ---------------------------------------------------------------------------

export function finalizeDocument(input: FinalizeInput): FinalizeResult {
  const loaded = loadDocx(input.path);
  const entries = loaded.entries.map((e) => ({ name: e.name, data: e.data }));

  // Strip w:rsid* attributes (Word's internal revision-save IDs) from the body.
  let documentXml = loaded.documentXml;
  const rsidMatches = documentXml.match(/\s+w:rsid[A-Za-z]*="[^"]*"/g);
  const rsidRemoved = rsidMatches ? rsidMatches.length : 0;
  if (rsidRemoved > 0) {
    documentXml = documentXml.replace(/\s+w:rsid[A-Za-z]*="[^"]*"/g, "");
  }
  const docEntry = entries.find((e) => e.name === "word/document.xml");
  if (docEntry) docEntry.data = Buffer.from(documentXml, "utf8");

  // Empty the author fields of docProps/core.xml.
  let metadataStripped = false;
  const core = entries.find((e) => e.name === "docProps/core.xml");
  if (core) {
    const scrubbed = core.data
      .toString("utf8")
      .replace(/(<dc:creator\b[^>]*>)[\s\S]*?(<\/dc:creator>)/g, "$1$2")
      .replace(/(<cp:lastModifiedBy\b[^>]*>)[\s\S]*?(<\/cp:lastModifiedBy>)/g, "$1$2");
    core.data = Buffer.from(scrubbed, "utf8");
    metadataStripped = true;
  }

  // Optional document protection.
  let protection: string | null = null;
  if (input.protectionMode !== undefined) {
    if (input.protectionMode !== "read_only") {
      throw new Error(`unsupported protection_mode: ${JSON.stringify(input.protectionMode)}`);
    }
    const element = `<w:documentProtection w:edit="readOnly" w:enforcement="1"/>`;
    const settings = entries.find((e) => e.name === "word/settings.xml");
    if (settings) {
      let xml = settings.data.toString("utf8");
      if (/<w:documentProtection\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<w:documentProtection\b[^>]*\/>/, element);
      } else if (/<w:documentProtection\b[^>]*>[\s\S]*?<\/w:documentProtection>/.test(xml)) {
        xml = xml.replace(/<w:documentProtection\b[^>]*>[\s\S]*?<\/w:documentProtection>/, element);
      } else if (/<w:settings\b[^>]*>/.test(xml)) {
        xml = xml.replace(/(<w:settings\b[^>]*>)/, `$1${element}`);
      } else {
        throw new Error("word/settings.xml has no <w:settings> root element");
      }
      settings.data = Buffer.from(xml, "utf8");
    } else {
      // Minimal settings part. Note: a full .docx also needs a
      // [Content_Types].xml override and a relationship entry for this
      // part; see BUILD-REPORT gaps.
      const minimal =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `${element}</w:settings>`;
      entries.push({ name: "word/settings.xml", data: Buffer.from(minimal, "utf8") });
    }
    protection = "read_only";
  }

  const outputPath = path.resolve(input.outputPath ?? siblingOutput(loaded.resolvedPath, "scrubbed"));
  writeFileSync(outputPath, writeZip(entries));
  return {
    path: loaded.resolvedPath,
    output_path: outputPath,
    metadata_stripped: metadataStripped,
    rsid_attributes_removed: rsidRemoved,
    protection,
  };
}
