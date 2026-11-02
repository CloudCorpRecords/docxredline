# docxredline — Build Report

Tranche #101+ Batch 1, MCP Rebuild 1000. Clean-room rebuild of
dealfluence/adeu (candidate `ai.adeu/adeu`, rank 21, 168 stars) as a
**labelled core subset**: the 3 agent tools its public README names
(`read_docx`, `process_document_batch`, `finalize_document`), specified
only from `harness/fixtures/adeu-tools-original.json` (docs-derived from
the README's Agent Workflows section, retrieved 2026-10-26). Local build
only — nothing pushed or published.

Identity: repo `docxredline`, npm `@reneturcios/docxredline`,
MCP name `io.github.CloudCorpRecords/docxredline`, version `0.1.0`.

## Files created (19/19 per TEMPLATE.md checklist)

- `package.json` — bin `docxredline: dist/cli.js`, `mcpName` set,
  scripts build=`tsc` / dev=`tsx src/cli.ts` / test=`tsx --test test/*.test.ts`
- `server.json` — description "Read, redline, and finalize .docx files
  with native Word tracked changes." = **73 chars** (≤100 ✓); npm entry
  carries `version: 0.1.0`; OCI entry
  `ghcr.io/cloudcorprecords/docxredline:0.1.0` carries **no** `version`
  field (verified programmatically)
- `Dockerfile` — node:20-alpine two-stage, ENTRYPOINT
  `node /app/dist/cli.js`, CMD `mcp`, label
  `io.modelcontextprotocol.server.name="io.github.CloudCorpRecords/docxredline"`
- `tsconfig.json` (318 bytes, fleet-identical), `.gitignore`, `LICENSE`
  (MIT, Copyright (c) 2026 Rene Turcios), `README.md` (fresh text),
  `CHANGELOG.md` (single 0.1.0 entry, 2026-10-26)
- `src/index.ts`, `src/cli.ts` (`read` / `batch` / `finalize` / `mcp`,
  `--version`; exit 2 usage / 1 file-or-batch error), `src/mcp.ts`,
  `src/types.ts`, `src/docx.ts` (core logic, `VERSION = "0.1.0"`)
- `test/docx.test.ts`
- `.github/workflows/{ci,publish,publish-container,registry-publish}.yml`
  (fleet mechanics; image name and smoke-test server name adapted)

## Implementation notes

- All logic is local (`.docx` files on disk) — no network, no FetchFn.
  Pure Node stdlib beyond SDK + zod: a hand-rolled ZIP reader (locates
  the End Of Central Directory, walks the central directory, inflates
  with `node:zlib` `inflateRawSync`; stored + deflated entries) and a
  stored-entry ZIP writer with a locally implemented CRC32.
- `read_docx` splits `word/document.xml` on `<w:p>`, concatenates
  `<w:t>` text per paragraph, and renders existing `<w:ins>` / `<w:del>`
  as CriticMarkup `{++ ++}` / `{-- --}`; `clean_view` resolves them
  (inserted text kept, deleted text dropped). `search_query` filters
  paragraphs; `page` groups paragraphs **40 per page** — a synthetic
  local contract (`.docx` has no pages), documented in the README.
- `process_document_batch` `modify`: the target is found in one
  paragraph's concatenated `<w:t>` text and the containing run is split
  so the old text lands in a `<w:del>` (as `<w:delText>`) and the new
  text in a following `<w:ins>`, stamped with the batch author and
  fresh `w:id`s. **Safety gate from the original:** a non-regex,
  `match_mode: "first"` target matching more than one paragraph is
  refused with an ambiguity error entry. `match_mode: "all"` and
  `regex: true` opt into the broader behaviour explicitly. Batch result
  is `{applied, failed: [{index, error}]}`; the output file is written
  only when at least one change applied.
- `accept` / `reject`: the change's `target_text` carries the `w:id`
  string of the `<w:ins>` / `<w:del>` element (local contract,
  documented in the README). Accept unwraps `<w:ins>` / drops
  `<w:del>`; reject does the inverse, restoring deleted text as
  `<w:t>` runs.
- `reply`: returns the per-change error `"comments not supported in
  this subset"` — counted as failed, never faked.
- `finalize_document`: empties `dc:creator` / `cp:lastModifiedBy` in
  `docProps/core.xml`, strips `w:rsid*` attributes from
  `word/document.xml`, and with `protection_mode: "read_only"` adds or
  replaces `<w:documentProtection w:edit="readOnly" w:enforcement="1"/>`
  in `word/settings.xml` (created minimal when absent — see gaps).
- `src/mcp.ts` uses the SDK's low-level `Server` (zod still validates
  every `tools/call`) so `tools/list` serves the spec's exact input
  schemas; `McpServer`'s schema conversion would inject `$schema` /
  `additionalProperties` keys absent from the spec and break parity.

## Verification (exact outputs)

### 1. Build
```
$ ./node_modules/.bin/tsc ; echo exit=$?
tsc exit=0
```
`dist/` produced (cli.js with shebang, docx.js, mcp.js, index.js,
types.js, plus .d.ts/.js.map).

### 2. Tests
```
$ ./node_modules/.bin/tsx --test test/*.test.ts
ℹ tests 28
ℹ pass 28
ℹ fail 0
```
28 tests (≥13 required). Fixtures are generated in-test in a temp dir
with the core's own stored-ZIP writer: a 3-paragraph document carrying
an existing `<w:ins w:id="8">` / `<w:del w:id="7">` pair, a creator in
`docProps/core.xml`, and an `w:rsidR` attribute. Coverage: CRC32 check
value, ZIP round-trip, deflated-entry reading, text extraction,
CriticMarkup vs clean view, search filter, 40-per-page synthetic
paging (45-paragraph doc), modify apply (output `document.xml` carries
`<w:ins>` with the new text and `<w:del>` with the old) + re-read in
both views, ambiguous-target refusal with no output written,
`match_mode: "all"`, regex modify, target-not-found, run-spanning
target refusal, accept (ins and del), reject (ins and del), reply
honest error, unknown-ID failure, finalize metadata scrub + rsid
count, protection into existing settings.xml, protection creating a
minimal settings.xml, unknown protection mode, missing-file and
not-a-ZIP errors.

### 3. Parity (harness)
```
$ python3 ~/workspace/mcp-rebuild-100/harness/parity.py \
    --original ~/workspace/mcp-rebuild-100/harness/fixtures/adeu-tools-original.json \
    --rebuild "stdio:node /home/hatch/workspace/mcp-rebuild-100/builds/docxredline/dist/cli.js mcp" \
    --out ~/workspace/mcp-rebuild-100/builds/docxredline/parity-out
verdict=PASS parity_pct=100.0 original_tools=3 rebuild_tools=3 missing=[] extra=[] schema_diffs=[]
(exit 0)
```
(Original is tools-file mode, so the harness records
`fixture_results: SKIPPED-TOOLS-FILE-MODE` — expected semantics.)

### 4. Live checks (stdio, against the built server)

A sample .docx was generated with the built core's ZIP writer
(2 paragraphs, one carrying the ins/del pair), then driven over stdio:

- initialize → `serverInfo: {"name":"docxredline","version":"0.1.0"}`
- `tools/call read_docx {path: /tmp/docxredline-live/sample.docx}` →
  `"text": "Payment is due within thirty days of invoice.\n\nThe term is {--twelve--}{++twenty-four++} months."`,
  `paragraph_count: 2, total_paragraphs: 2, page_count: 1`
- `tools/call read_docx {path: /tmp/docxredline-live/nonexistent.docx}` →
  `{"content":[{"type":"text","text":"ENOENT: no such file or directory, open '/tmp/docxredline-live/nonexistent.docx'"}],"isError":true}`
- CLI: `--version` → `0.1.0`; `read --clean` on the sample →
  "The term is twenty-four months."; unknown command prints Usage and
  exits 2.

## Lockfile status

**Clean lockfile, generated via Route 3** (LOCKFILE-ROUTE.md):
`package.json` alone in fresh `/tmp/docxredline-lock/`,
`npm install --package-lock-only --ignore-scripts --offline` →
exit 0, no `node_modules` created; the produced `package-lock.json`
parses, `name` == `@reneturcios/docxredline`, version 0.1.0,
lockfileVersion 3, 128 package entries. Copied into the build dir as
the committed lockfile. (`node_modules` in the build dir is a `cp -a`
copy of clockport's — file counts verified identical at 3,690 files;
the copy logged ownership-preservation warnings only, and `tsc` /
`tsx` from it run correctly. It is gitignored and not part of the
deliverable.)

## Gaps vs the original (honest subset boundaries)

- **Run-splitting:** a `modify` target that spans multiple `<w:t>`
  runs inside its paragraph is refused as a per-change failure
  ("target spans multiple runs; not supported in this subset"); the
  original operates on a full run model. Replacements also reopen the
  suffix run without the original run properties (`w:rPr`).
- **Accept/reject ID model:** this subset addresses one tracked-change
  element by its `w:id` string only; the original's richer ID model
  (comments, ranges, nested change sets) is not reproduced.
- **Comments:** `reply` (and comment support generally) returns an
  honest per-change error; comments are neither read nor written.
- **Synthetic paging:** `page` groups paragraphs 40-per-page as a
  documented local contract; real pagination is layout-dependent and
  out of scope for a file-level subset.
- **settings.xml creation:** when `word/settings.xml` is absent and
  protection is requested, a minimal part is created, but
  `[Content_Types].xml` and the package relationships are not updated
  to register it (noted in the README); when the part exists, the
  protection element is added/replaced in place.
- **Omitted entirely** (per the fixture's `_source`): live MS Word
  integration (Windows/Python backend), MCP Apps UI rendering, and the
  original's SDK/CLI surfaces.
- Container build not exercised locally (no Docker run in this wave);
  the publish-container workflow smoke-tests the image at publish time.

## Blockers

None for the local build.
