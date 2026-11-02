# docxredline

docxredline works on Word `.docx` files the way an agent needs to:
read one as plain text you can reason over, edit one so Word itself
shows the changes as tracked redlines, and scrub one clean before it
leaves the building. Everything is local — documents never leave the
machine, and there is no server component beyond the stdio MCP face.

This is a labelled core subset: it covers the three document tools
(`read_docx`, `process_document_batch`, `finalize_document`) and nothing
else. It does not drive a live Word instance, render any UI, or handle
document comments (see the notes per tool below).

## Install

```bash
npx @reneturcios/docxredline read agreement.docx
```

Or run the container image:

```bash
docker run -i ghcr.io/cloudcorprecords/docxredline
```

The container defaults to the MCP server (see below). Pass CLI
arguments to use it as a CLI instead. Note the CLI reads and writes
files by path, so give the container a volume mount for your documents.

## CLI

```bash
# Read a document (tracked changes shown as CriticMarkup)
docxredline read agreement.docx

# Read the final clean version instead
docxredline read agreement.docx --clean

# Only paragraphs mentioning payment terms; synthetic page 2
docxredline read agreement.docx --search "payment"
docxredline read agreement.docx --page 2

# Apply a batch of edits from a JSON file
docxredline batch agreement.docx changes.json --author "Review Bot"

# Pre-send scrub, locked read-only
docxredline finalize agreement.docx --protection read_only

docxredline --version
```

Exit codes: `0` on success, `1` when a file cannot be read or no change
in a batch applies, `2` for usage errors. All commands accept `--json`.

A changes file is a JSON array of change objects:

```json
[
  { "type": "modify", "target_text": "thirty days", "new_text": "sixty days" },
  { "type": "modify", "target_text": "old clause", "new_text": "new clause", "match_mode": "all" },
  { "type": "accept", "target_text": "12" },
  { "type": "reject", "target_text": "13" }
]
```

## MCP server

One bin serves both faces. The `mcp` subcommand starts a stdio MCP
server:

```bash
npx -y @reneturcios/docxredline mcp
```

MCP client configuration:

```json
{
  "mcpServers": {
    "docxredline": {
      "command": "npx",
      "args": ["-y", "@reneturcios/docxredline", "mcp"]
    }
  }
}
```

### Tools

| Tool | Input | What it does |
| --- | --- | --- |
| `read_docx` | `path` (string, required); `clean_view` (boolean), `search_query` (string), `page` (number) optional | Returns the document as Markdown-like text, paragraph by paragraph. Existing tracked changes render as CriticMarkup — `{++inserted++}` and `{--deleted--}` — unless `clean_view` resolves them away. `search_query` keeps only paragraphs containing the text. |
| `process_document_batch` | `path` (string, required), `changes` (array, required); `output_path`, `author` optional | Applies a batch of changes and writes the result (default: `<name>.redlined.docx` beside the input). Returns how many changes applied and a per-change error list for the rest. The output is written only when at least one change applied. |
| `finalize_document` | `path` (string, required); `output_path`, `protection_mode` optional | Writes a distribution-ready copy (default: `<name>.scrubbed.docx`): author fields in `docProps/core.xml` emptied, `w:rsid*` attributes stripped from the body, and — with `protection_mode: "read_only"` — a `documentProtection` lock added. |

### How edits and IDs work in this subset

- A `modify` change finds `target_text` inside a single paragraph's
  text and rewrites the containing run: the old text is wrapped in a
  `<w:del>` element and the new text in a `<w:ins>` element, both
  stamped with the batch `author` and fresh `w:id`s, so Word displays
  a genuine tracked change.
- Safety gate: a non-regex `modify` with the default `match_mode`
  `"first"` whose target matches **more than one paragraph** is
  refused with an ambiguity error rather than applied to a guessed
  paragraph. Use a longer target, or `match_mode: "all"` to edit every
  occurrence in every matching paragraph. `regex: true` treats
  `target_text` as a regular expression.
- A target that spans multiple runs inside its paragraph cannot be
  rewritten by this subset and is reported as a per-change failure.
- `accept` / `reject` take the tracked change's `w:id` string as
  `target_text`. Accept keeps inserted text and drops deletions;
  reject does the inverse. (The original tool's richer ID model for
  comments and ranges is not reproduced here.)
- `reply` is not supported in this subset: it comes back as a
  per-change failure, `"comments not supported in this subset"`, and
  is never silently counted as applied.

### Synthetic paging

`.docx` has no page concept, so `page` is a local contract: paragraphs
are grouped 40 to a page, numbered from 1. `read_docx` reports the
document's total paragraph count and synthetic page count with every
result so callers can page deliberately.

### finalize notes

When the document has no `word/settings.xml` part and protection is
requested, docxredline creates a minimal one carrying only the
protection element. A complete package would also register that part
in `[Content_Types].xml` and the package relationships; this subset
records that as a known gap rather than rewriting those parts.

## Development

```bash
npm install
npm run build   # tsc -> dist/
npm test        # node test runner over test/*.test.ts, no network
npm run dev     # run the CLI from source via tsx
```

Tests generate their own minimal `.docx` fixtures in a temp directory
with the same stored-ZIP writer the tool ships, so the suite exercises
the real archive reader end to end.

## License

MIT — see `LICENSE`.
