/**
 * docxredline types.
 *
 * Everything here is a local file contract: inputs name .docx files on
 * disk, outputs name the file that was written plus what happened inside
 * it. There is no network surface and no injectable dependency.
 */

export interface ReadDocxOptions {
  /** Resolve existing tracked changes away (accept-all view). */
  cleanView?: boolean;
  /** Only return paragraphs whose rendered text contains this string. */
  searchQuery?: string;
  /** Synthetic 1-based page to return; paragraphs are grouped 40 per page. */
  page?: number;
}

export interface ReadDocxResult {
  /** The resolved input path that was read. */
  path: string;
  /** The returned paragraphs joined with blank lines. */
  text: string;
  /** The returned paragraphs, in document order. */
  paragraphs: string[];
  /** How many paragraphs are in `paragraphs` (after filter / paging). */
  paragraph_count: number;
  /** How many paragraphs the document holds in total (before filter / paging). */
  total_paragraphs: number;
  /** Total synthetic pages at 40 paragraphs per page. */
  page_count: number;
  /** The page that was returned, when a page was requested. */
  page?: number;
}

export interface DocChange {
  /** modify, accept, reject, or reply. */
  type: string;
  /**
   * For modify: the text (or regex, when `regex` is set) to find.
   * For accept / reject: the `w:id` string of the tracked-change element.
   */
  target_text?: string;
  /** Replacement text for modify. */
  new_text?: string;
  /** Comment text; carried for shape compatibility, unused in this subset. */
  comment?: string;
  /** "first" (default) or "all". */
  match_mode?: string;
  /** Treat target_text as a regular expression for modify. */
  regex?: boolean;
}

export interface BatchInput {
  path: string;
  outputPath?: string;
  author?: string;
  changes: DocChange[];
}

export interface BatchFailure {
  /** Index of the change in the submitted batch. */
  index: number;
  error: string;
}

export interface BatchResult {
  /** How many changes were applied. */
  applied: number;
  /** One entry per change that could not be applied, with the reason. */
  failed: BatchFailure[];
  /** Where the edited document was written; present only when applied >= 1. */
  output_path?: string;
}

export interface FinalizeInput {
  path: string;
  outputPath?: string;
  /** e.g. "read_only" to lock the document. */
  protectionMode?: string;
}

export interface FinalizeResult {
  /** The resolved input path. */
  path: string;
  /** Where the scrubbed document was written. */
  output_path: string;
  /** True when docProps/core.xml was present and its author fields emptied. */
  metadata_stripped: boolean;
  /** How many w:rsid* attributes were removed from word/document.xml. */
  rsid_attributes_removed: number;
  /** The protection mode that was applied, or null when none was requested. */
  protection: string | null;
}

/** One entry of a ZIP archive, in archive order. */
export interface ZipEntry {
  name: string;
  data: Buffer;
}
