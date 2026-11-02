/**
 * docxredline MCP server. Tools: read_docx, process_document_batch,
 * finalize_document — a labelled core subset of the public spec.
 *
 * Why the low-level Server instead of McpServer: the high-level helper
 * converts zod schemas to JSON Schema and injects two keys that are not
 * in this tool's public spec ("$schema" and "additionalProperties"),
 * which breaks exact tool-list/schema parity with the spec. This file
 * therefore serves tools/list with the spec's exact input schemas,
 * while zod still validates every tools/call argument set. Tool names,
 * parameters, and parameter descriptions are taken from the public spec.
 *
 * All logic is local: the tools read and write .docx files on disk.
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { finalizeDocument, processDocumentBatch, readDocx, VERSION } from "./docx.js";

const ReadDocxArgs = z.object({
  path: z.string().describe("Path to the .docx file."),
  clean_view: z
    .boolean()
    .optional()
    .describe("Read the final clean version (accept all tracked changes)."),
  search_query: z
    .string()
    .optional()
    .describe("Only return paragraphs containing this text."),
  page: z
    .number()
    .optional()
    .describe("Synthetic page number to return (paragraphs are grouped into pages)."),
});

const ChangeArgs = z.object({
  type: z.string().describe("modify, accept, reject, or reply."),
  target_text: z.string().optional(),
  new_text: z.string().optional(),
  comment: z.string().optional(),
  match_mode: z.string().optional().describe("first (default) or all."),
  regex: z.boolean().optional(),
});

const ProcessDocumentBatchArgs = z.object({
  path: z.string(),
  output_path: z
    .string()
    .optional()
    .describe("Where to write the edited .docx. Defaults to <name>.redlined.docx beside the input."),
  author: z.string().optional().describe("Author name recorded on the tracked changes."),
  changes: z.array(ChangeArgs).describe("The list of changes to apply atomically."),
});

const FinalizeDocumentArgs = z.object({
  path: z.string(),
  output_path: z.string().optional(),
  protection_mode: z.string().optional().describe("e.g. read_only to lock the document."),
});

const TOOLS = [
  {
    name: "read_docx",
    description:
      "Read a .docx file as clean Markdown-like text (CriticMarkup view of existing tracked changes), optionally filtered.",
    inputSchema: {
      type: "object" as const,
      properties: {
        path: { type: "string", description: "Path to the .docx file." },
        clean_view: {
          type: "boolean",
          description: "Read the final clean version (accept all tracked changes).",
        },
        search_query: {
          type: "string",
          description: "Only return paragraphs containing this text.",
        },
        page: {
          type: "number",
          description: "Synthetic page number to return (paragraphs are grouped into pages).",
        },
      },
      required: ["path"],
    },
  },
  {
    name: "process_document_batch",
    description:
      "Apply a batch of edits to a .docx as native Word tracked changes (redlines), or accept/reject/reply to existing tracked changes and comments by ID.",
    inputSchema: {
      type: "object" as const,
      properties: {
        path: { type: "string" },
        output_path: {
          type: "string",
          description:
            "Where to write the edited .docx. Defaults to <name>.redlined.docx beside the input.",
        },
        author: { type: "string", description: "Author name recorded on the tracked changes." },
        changes: {
          type: "array",
          items: {
            type: "object",
            properties: {
              type: { type: "string", description: "modify, accept, reject, or reply." },
              target_text: { type: "string" },
              new_text: { type: "string" },
              comment: { type: "string" },
              match_mode: { type: "string", description: "first (default) or all." },
              regex: { type: "boolean" },
            },
            required: ["type"],
          },
          description: "The list of changes to apply atomically.",
        },
      },
      required: ["path", "changes"],
    },
  },
  {
    name: "finalize_document",
    description:
      "Pre-send scrub: strip author metadata and internal tracking IDs from a .docx, optionally lock it, and prepare it for distribution.",
    inputSchema: {
      type: "object" as const,
      properties: {
        path: { type: "string" },
        output_path: { type: "string" },
        protection_mode: { type: "string", description: "e.g. read_only to lock the document." },
      },
      required: ["path"],
    },
  },
];

function textResult(value: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(value, null, 2) }] };
}

function errorResult(err: unknown) {
  return {
    content: [
      { type: "text" as const, text: err instanceof Error ? err.message : String(err) },
    ],
    isError: true,
  };
}

const server = new Server(
  { name: "docxredline", version: VERSION },
  { capabilities: { tools: {} } },
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  try {
    if (name === "read_docx") {
      const parsed = ReadDocxArgs.parse(args ?? {});
      return textResult(
        readDocx(parsed.path, {
          cleanView: parsed.clean_view,
          searchQuery: parsed.search_query,
          page: parsed.page,
        }),
      );
    }
    if (name === "process_document_batch") {
      const parsed = ProcessDocumentBatchArgs.parse(args ?? {});
      return textResult(
        processDocumentBatch({
          path: parsed.path,
          outputPath: parsed.output_path,
          author: parsed.author,
          changes: parsed.changes,
        }),
      );
    }
    if (name === "finalize_document") {
      const parsed = FinalizeDocumentArgs.parse(args ?? {});
      return textResult(
        finalizeDocument({
          path: parsed.path,
          outputPath: parsed.output_path,
          protectionMode: parsed.protection_mode,
        }),
      );
    }
  } catch (err) {
    return errorResult(err);
  }
  throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
});

async function main(): Promise<void> {
  await server.connect(new StdioServerTransport());
}

void main();
