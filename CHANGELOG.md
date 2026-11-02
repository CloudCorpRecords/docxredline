# Changelog

## 0.1.0 (2026-10-26)

- Initial release. A labelled core subset of the three agent tools in the
  public spec: `read_docx`, `process_document_batch`, `finalize_document`.
- CLI: `docxredline read <path>`, `docxredline batch <path> <changes.json>`,
  `docxredline finalize <path>`, plus `docxredline mcp` and `--version`.
- `read_docx` renders a .docx as Markdown-like text: existing tracked
  changes appear as CriticMarkup (`{++ ++}` / `{-- --}`), or are resolved
  away with `clean_view`. Paragraphs can be filtered with `search_query`
  and paged synthetically at 40 paragraphs per page.
- `process_document_batch` applies `modify` edits as native Word tracked
  changes (`<w:ins>` / `<w:del>`), and `accept` / `reject` by tracked-change
  ID. A non-regex `first` target that matches more than one paragraph is
  refused as ambiguous. `reply` returns an honest per-change error:
  comments are not supported in this subset.
- `finalize_document` strips author metadata from `docProps/core.xml`,
  removes `w:rsid*` attributes from the document body, and can lock the
  document with `protection_mode: "read_only"`.
- All logic is local and uses the Node standard library only (hand-rolled
  ZIP reader/writer, `node:zlib` inflate) beyond the MCP SDK + zod.
