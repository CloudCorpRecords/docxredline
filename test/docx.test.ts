import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateRawSync } from "node:zlib";
import {
  crc32,
  finalizeDocument,
  processDocumentBatch,
  readDocx,
  readZip,
  writeZip,
} from "../src/docx.js";
import type { ZipEntry } from "../src/types.js";

// ---------------------------------------------------------------------------
// Fixture .docx generation (stored-ZIP writer from the core itself)
// ---------------------------------------------------------------------------

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:document xmlns:w="${W_NS}"><w:body>` +
  `<w:p w:rsidR="00AB12CD"><w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>Payment is due within thirty days of invoice.</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">The term is </w:t></w:r>` +
  `<w:del w:id="7" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>twelve</w:delText></w:r></w:del>` +
  `<w:ins w:id="8" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>twenty-four</w:t></w:r></w:ins>` +
  `<w:r><w:t xml:space="preserve"> months.</w:t></w:r></w:p>` +
  `</w:body></w:document>`;

const CORE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
  `<dc:creator>Jane Drafter</dc:creator><cp:lastModifiedBy>Jane Drafter</cp:lastModifiedBy>` +
  `<dc:title>Agreement</dc:title></cp:coreProperties>`;

const SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:settings xmlns:w="${W_NS}"><w:zoom w:percent="100"/></w:settings>`;

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

function makeDocx(
  documentXml: string = DOCUMENT_XML,
  opts: { core?: boolean; settings?: boolean } = {},
): string {
  const dir = mkdtempSync(path.join(tmpdir(), "docxredline-test-"));
  const file = path.join(dir, "sample.docx");
  const entries: ZipEntry[] = [
    { name: "[Content_Types].xml", data: Buffer.from(CONTENT_TYPES_XML, "utf8") },
    { name: "word/document.xml", data: Buffer.from(documentXml, "utf8") },
  ];
  if (opts.core !== false) {
    entries.push({ name: "docProps/core.xml", data: Buffer.from(CORE_XML, "utf8") });
  }
  if (opts.settings) {
    entries.push({ name: "word/settings.xml", data: Buffer.from(SETTINGS_XML, "utf8") });
  }
  writeFileSync(file, writeZip(entries));
  return file;
}

function entryText(file: string, name: string): string {
  const entries = readZip(readFileSync(file));
  const entry = entries.find((e) => e.name === name);
  assert.ok(entry, `entry ${name} present`);
  return entry.data.toString("utf8");
}

// ---------------------------------------------------------------------------
// ZIP layer
// ---------------------------------------------------------------------------

describe("zip", () => {
  it("crc32 matches the known check value for '123456789'", () => {
    assert.equal(crc32(new TextEncoder().encode("123456789")), 0xcbf43926);
  });

  it("round-trips stored entries", () => {
    const entries: ZipEntry[] = [
      { name: "a.txt", data: Buffer.from("hello world", "utf8") },
      { name: "dir/b.xml", data: Buffer.from("<x>1</x>", "utf8") },
    ];
    const back = readZip(writeZip(entries));
    assert.equal(back.length, 2);
    assert.equal(back[0].name, "a.txt");
    assert.equal(back[0].data.toString("utf8"), "hello world");
    assert.equal(back[1].name, "dir/b.xml");
    assert.equal(back[1].data.toString("utf8"), "<x>1</x>");
  });

  it("reads deflated entries", () => {
    const data = Buffer.from("deflate me, deflate me, deflate me", "utf8");
    const comp = deflateRawSync(data);
    const name = Buffer.from("word/document.xml", "utf8");
    const crc = crc32(data);
    const u16 = (n: number) => [n & 0xff, (n >>> 8) & 0xff];
    const u32 = (n: number) => [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff];
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(8), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(comp.length), ...u32(data.length), ...u16(name.length), ...u16(0),
    ];
    const localLen = local.length + name.length + comp.length;
    const central = [
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(8), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(comp.length), ...u32(data.length), ...u16(name.length), ...u16(0),
      ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(0),
    ];
    const end = [
      ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(1), ...u16(1),
      ...u32(central.length + name.length), ...u32(localLen), ...u16(0),
    ];
    const zip = Buffer.from([...local, ...name, ...comp, ...central, ...name, ...end]);
    const entries = readZip(zip);
    assert.equal(entries.length, 1);
    assert.equal(entries[0].data.toString("utf8"), data.toString("utf8"));
  });
});

// ---------------------------------------------------------------------------
// read_docx
// ---------------------------------------------------------------------------

describe("readDocx", () => {
  it("extracts paragraph text", () => {
    const r = readDocx(makeDocx());
    assert.equal(r.total_paragraphs, 3);
    assert.equal(r.paragraphs[0], "The quick brown fox jumps over the lazy dog.");
    assert.equal(r.paragraphs[1], "Payment is due within thirty days of invoice.");
  });

  it("renders existing tracked changes as CriticMarkup by default", () => {
    const r = readDocx(makeDocx());
    assert.equal(r.paragraphs[2], "The term is {--twelve--}{++twenty-four++} months.");
  });

  it("clean_view resolves tracked changes", () => {
    const r = readDocx(makeDocx(), { cleanView: true });
    assert.equal(r.paragraphs[2], "The term is twenty-four months.");
  });

  it("search_query filters paragraphs", () => {
    const r = readDocx(makeDocx(), { searchQuery: "Payment" });
    assert.deepEqual(r.paragraphs, ["Payment is due within thirty days of invoice."]);
    assert.equal(r.paragraph_count, 1);
    assert.equal(r.total_paragraphs, 3);
  });

  it("page groups paragraphs 40 per page", () => {
    const paras = Array.from(
      { length: 45 },
      (_, i) => `<w:p><w:r><w:t>Paragraph ${i + 1}</w:t></w:r></w:p>`,
    ).join("");
    const xml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>${paras}</w:body></w:document>`;
    const file = makeDocx(xml);
    const page1 = readDocx(file, { page: 1 });
    assert.equal(page1.paragraph_count, 40);
    assert.equal(page1.page_count, 2);
    const page2 = readDocx(file, { page: 2 });
    assert.equal(page2.paragraph_count, 5);
    assert.equal(page2.paragraphs[0], "Paragraph 41");
  });

  it("throws the real error for a missing file", () => {
    assert.throws(() => readDocx("/nonexistent/definitely-missing.docx"), /ENOENT/);
  });

  it("throws for a file that is not a ZIP", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "docxredline-test-"));
    const file = path.join(dir, "fake.docx");
    writeFileSync(file, "this is not a zip archive at all");
    assert.throws(() => readDocx(file), /not a ZIP archive/);
  });
});

// ---------------------------------------------------------------------------
// process_document_batch
// ---------------------------------------------------------------------------

describe("processDocumentBatch modify", () => {
  it("applies a modify as native tracked changes", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      author: "Reviewer",
      changes: [{ type: "modify", target_text: "thirty", new_text: "sixty" }],
    });
    assert.equal(r.applied, 1);
    assert.deepEqual(r.failed, []);
    assert.equal(r.output_path, out);
    const xml = entryText(out, "word/document.xml");
    assert.ok(xml.includes("<w:ins"), "output carries a w:ins element");
    assert.ok(xml.includes("<w:del"), "output carries a w:del element");
    assert.ok(xml.includes(">sixty</w:t>"), "inserted text present");
    assert.ok(xml.includes(">thirty</w:delText>"), "deleted text present");
    assert.ok(xml.includes('w:author="Reviewer"'), "author recorded");
  });

  it("the redlined output re-reads as CriticMarkup, and clean without the old text", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "modify", target_text: "thirty", new_text: "sixty" }],
    });
    const markup = readDocx(out);
    assert.equal(
      markup.paragraphs[1],
      "Payment is due within {--thirty--}{++sixty++} days of invoice.",
    );
    const clean = readDocx(out, { cleanView: true });
    assert.equal(clean.paragraphs[1], "Payment is due within sixty days of invoice.");
  });

  it("refuses an ambiguous non-regex 'first' target and writes nothing", () => {
    const xml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:r><w:t>Please review the first clause.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Please review the second clause.</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const file = makeDocx(xml);
    const r = processDocumentBatch({
      path: file,
      changes: [{ type: "modify", target_text: "review", new_text: "check" }],
    });
    assert.equal(r.applied, 0);
    assert.equal(r.failed.length, 1);
    assert.match(r.failed[0].error, /ambiguous target/);
    assert.equal(r.output_path, undefined);
    assert.equal(existsSync(path.join(path.dirname(file), "sample.redlined.docx")), false);
  });

  it("match_mode 'all' edits every matching paragraph", () => {
    const xml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:r><w:t>Please review the first clause.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Please review the second clause.</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const file = makeDocx(xml);
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "modify", target_text: "review", new_text: "check", match_mode: "all" }],
    });
    assert.equal(r.applied, 1);
    const clean = readDocx(out, { cleanView: true });
    assert.deepEqual(clean.paragraphs, [
      "Please check the first clause.",
      "Please check the second clause.",
    ]);
  });

  it("supports regex targets", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "modify", target_text: "thirt\\w+", new_text: "sixty", regex: true }],
    });
    assert.equal(r.applied, 1);
    const clean = readDocx(out, { cleanView: true });
    assert.equal(clean.paragraphs[1], "Payment is due within sixty days of invoice.");
  });

  it("reports a target that is not found", () => {
    const file = makeDocx();
    const r = processDocumentBatch({
      path: file,
      changes: [{ type: "modify", target_text: "nonexistent phrase", new_text: "x" }],
    });
    assert.equal(r.applied, 0);
    assert.match(r.failed[0].error, /target not found/);
  });

  it("refuses a target spanning multiple runs", () => {
    const xml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:r><w:t>Pay</w:t></w:r><w:r><w:t>ment is due.</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const file = makeDocx(xml);
    const r = processDocumentBatch({
      path: file,
      changes: [{ type: "modify", target_text: "Payment", new_text: "Billing" }],
    });
    assert.equal(r.applied, 0);
    assert.match(r.failed[0].error, /spans multiple runs/);
  });
});

describe("processDocumentBatch accept / reject / reply", () => {
  it("accept unwraps an insertion and keeps its text", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "accept", target_text: "8" }],
    });
    assert.equal(r.applied, 1);
    const xml = entryText(out, "word/document.xml");
    assert.ok(!xml.includes('w:id="8"'), "the accepted w:ins is gone");
    const clean = readDocx(out, { cleanView: true });
    assert.equal(clean.paragraphs[2], "The term is twenty-four months.");
  });

  it("accept drops a deletion entirely", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "accept", target_text: "7" }],
    });
    assert.equal(r.applied, 1);
    const markup = readDocx(out);
    assert.equal(markup.paragraphs[2], "The term is {++twenty-four++} months.");
  });

  it("reject restores a deletion as ordinary text", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "reject", target_text: "7" }],
    });
    assert.equal(r.applied, 1);
    const xml = entryText(out, "word/document.xml");
    assert.ok(xml.includes(">twelve</w:t>"), "deleted text restored as a w:t run");
    const markup = readDocx(out);
    assert.equal(markup.paragraphs[2], "The term is twelve{++twenty-four++} months.");
  });

  it("reject drops an insertion", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "edited.docx");
    const r = processDocumentBatch({
      path: file,
      outputPath: out,
      changes: [{ type: "reject", target_text: "8" }],
    });
    assert.equal(r.applied, 1);
    const markup = readDocx(out);
    assert.equal(markup.paragraphs[2], "The term is {--twelve--} months.");
  });

  it("reply returns an honest per-change error and applies nothing", () => {
    const file = makeDocx();
    const r = processDocumentBatch({
      path: file,
      changes: [{ type: "reply", target_text: "1", comment: "Agreed." }],
    });
    assert.equal(r.applied, 0);
    assert.equal(r.failed.length, 1);
    assert.equal(r.failed[0].error, "comments not supported in this subset");
    assert.equal(r.output_path, undefined);
  });

  it("an unknown tracked-change id is a per-change failure", () => {
    const file = makeDocx();
    const r = processDocumentBatch({
      path: file,
      changes: [{ type: "accept", target_text: "999" }],
    });
    assert.equal(r.applied, 0);
    assert.match(r.failed[0].error, /no tracked change with w:id/);
  });
});

// ---------------------------------------------------------------------------
// finalize_document
// ---------------------------------------------------------------------------

describe("finalizeDocument", () => {
  it("strips author metadata and rsid attributes", () => {
    const file = makeDocx();
    const out = path.join(path.dirname(file), "final.docx");
    const r = finalizeDocument({ path: file, outputPath: out });
    assert.equal(r.metadata_stripped, true);
    assert.ok(r.rsid_attributes_removed >= 1, "at least the fixture rsidR was removed");
    assert.equal(r.protection, null);
    const core = entryText(out, "docProps/core.xml");
    assert.ok(core.includes("<dc:creator></dc:creator>"));
    assert.ok(core.includes("<cp:lastModifiedBy></cp:lastModifiedBy>"));
    assert.ok(!core.includes("Jane Drafter"), "author name is gone");
    assert.ok(core.includes("<dc:title>Agreement</dc:title>"), "other metadata kept");
    const doc = entryText(out, "word/document.xml");
    assert.ok(!doc.includes("w:rsid"), "no rsid attributes remain");
  });

  it("read_only adds a documentProtection element to an existing settings.xml", () => {
    const file = makeDocx(DOCUMENT_XML, { settings: true });
    const out = path.join(path.dirname(file), "final.docx");
    const r = finalizeDocument({ path: file, outputPath: out, protectionMode: "read_only" });
    assert.equal(r.protection, "read_only");
    const settings = entryText(out, "word/settings.xml");
    assert.ok(
      settings.includes('<w:documentProtection w:edit="readOnly" w:enforcement="1"/>'),
      "protection element present",
    );
  });

  it("read_only creates a minimal settings.xml when the part is absent", () => {
    const file = makeDocx(DOCUMENT_XML, { settings: false });
    const out = path.join(path.dirname(file), "final.docx");
    const r = finalizeDocument({ path: file, outputPath: out, protectionMode: "read_only" });
    assert.equal(r.protection, "read_only");
    const settings = entryText(out, "word/settings.xml");
    assert.ok(settings.includes("<w:documentProtection"));
  });

  it("rejects an unknown protection_mode", () => {
    const file = makeDocx();
    assert.throws(
      () => finalizeDocument({ path: file, protectionMode: "encrypt" }),
      /unsupported protection_mode/,
    );
  });

  it("throws the real error for a missing file", () => {
    assert.throws(() => finalizeDocument({ path: "/nonexistent/nope.docx" }), /ENOENT/);
  });
});
