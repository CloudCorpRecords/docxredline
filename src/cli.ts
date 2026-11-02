#!/usr/bin/env node
/**
 * docxredline CLI.
 *
 *   docxredline read <path.docx> [--clean] [--search <text>] [--page <n>] [--json]
 *   docxredline batch <path.docx> <changes.json> [--output <path>] [--author <name>] [--json]
 *   docxredline finalize <path.docx> [--output <path>] [--protection read_only] [--json]
 *   docxredline mcp
 *
 * `batch` reads its changes from a JSON file holding an array of change
 * objects ({type, target_text, new_text, ...}).
 *
 * Exit codes: 0 success, 1 file/edit error, 2 usage error.
 */
import { readFileSync } from "node:fs";
import { finalizeDocument, processDocumentBatch, readDocx, VERSION } from "./docx.js";
import type { DocChange } from "./types.js";

function usage(): never {
  console.error(`Usage:
  docxredline read <path.docx> [--clean] [--search <text>] [--page <n>] [--json]
  docxredline batch <path.docx> <changes.json> [--output <path>] [--author <name>] [--json]
  docxredline finalize <path.docx> [--output <path>] [--protection <mode>] [--json]
  docxredline mcp
  docxredline --version`);
  process.exit(2);
}

function flagValue(args: string[], flag: string): string | undefined {
  const i = args.indexOf(flag);
  return i !== -1 && i + 1 < args.length ? args[i + 1] : undefined;
}

function positional(args: string[], flagsWithValues: string[]): string[] {
  const out: string[] = [];
  for (let i = 0; i < args.length; i++) {
    if (flagsWithValues.includes(args[i])) i++;
    else if (!args[i].startsWith("--")) out.push(args[i]);
  }
  return out;
}

async function main(): Promise<void> {
  const [cmd, ...rest] = process.argv.slice(2);
  if (cmd === "--version" || cmd === "-v") {
    console.log(VERSION);
    return;
  }
  if (cmd === "mcp") {
    await import("./mcp.js");
    return;
  }
  if (!cmd) usage();
  const asJson = rest.includes("--json");
  try {
    if (cmd === "read") {
      const [input] = positional(rest, ["--search", "--page"]);
      if (!input) usage();
      const pageRaw = flagValue(rest, "--page");
      const result = readDocx(input, {
        cleanView: rest.includes("--clean"),
        searchQuery: flagValue(rest, "--search"),
        page: pageRaw !== undefined ? Number(pageRaw) : undefined,
      });
      if (asJson) console.log(JSON.stringify(result, null, 2));
      else console.log(result.text);
    } else if (cmd === "batch") {
      const [input, changesFile] = positional(rest, ["--output", "--author"]);
      if (!input || !changesFile) usage();
      const changes = JSON.parse(readFileSync(changesFile, "utf8")) as DocChange[];
      if (!Array.isArray(changes)) {
        console.error("error: changes file must hold a JSON array of changes");
        process.exit(1);
      }
      const result = processDocumentBatch({
        path: input,
        outputPath: flagValue(rest, "--output"),
        author: flagValue(rest, "--author"),
        changes,
      });
      if (asJson) console.log(JSON.stringify(result, null, 2));
      else {
        console.log(
          `applied ${result.applied}, failed ${result.failed.length}` +
            (result.output_path ? ` -> ${result.output_path}` : ""),
        );
        for (const f of result.failed) console.log(`  change ${f.index}: ${f.error}`);
      }
      if (result.applied === 0 && result.failed.length > 0) process.exit(1);
    } else if (cmd === "finalize") {
      const [input] = positional(rest, ["--output", "--protection"]);
      if (!input) usage();
      const result = finalizeDocument({
        path: input,
        outputPath: flagValue(rest, "--output"),
        protectionMode: flagValue(rest, "--protection"),
      });
      if (asJson) console.log(JSON.stringify(result, null, 2));
      else
        console.log(
          `scrubbed -> ${result.output_path} (metadata_stripped=${result.metadata_stripped}, ` +
            `rsid_removed=${result.rsid_attributes_removed}, protection=${result.protection})`,
        );
    } else {
      usage();
    }
  } catch (err) {
    console.error(`error: ${err instanceof Error ? err.message : String(err)}`);
    process.exit(1);
  }
}

void main();
