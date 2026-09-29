/** Plain-text building blocks shared by the export writers. Dependency-free
 * on purpose: importable headlessly (unit tests) without the Obsidian API. */

/** Inline formatting for a slice of text. */
export interface TextRunModel {
  text: string;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  code: boolean;
}

/** Block-level content of a note. Text-only by design (v1 skips images,
 * tables, math and embeds). */
export type Block =
  | { kind: "heading"; level: 1 | 2 | 3 | 4 | 5 | 6; runs: TextRunModel[] }
  | { kind: "paragraph"; runs: TextRunModel[] }
  | { kind: "list"; ordered: boolean; items: TextRunModel[][] }
  | { kind: "break" };

/** One export unit: a note's title plus its parsed blocks. A chapter export
 * is a single entry; a novel export is one entry per note in the folder, in
 * reading order. Writers reset their per-note conventions (flush first
 * paragraph, drop cap, page break) at every chapter boundary. */
export interface Chapter {
  title: string;
  blocks: Block[];
}

/** A level-1 title heading: the chapter's file name, prepended to the note so
 * every exported chapter is labelled in DOCX, EPUB and PDF alike. */
export function titleHeading(text: string): Block {
  return {
    kind: "heading",
    level: 1,
    runs: [{ text, bold: false, italic: false, strike: false, code: false }],
  };
}

/** Case- and whitespace-insensitive form used to compare a chapter title with
 * a note's own first heading. */
function normalizeTitle(text: string): string {
  return text.trim().toLowerCase().replace(/\s+/g, " ");
}

/** Whether the chapter-title heading still needs to be prepended: only when
 * the note doesn't already open with a heading whose text is the chapter
 * name (never duplicate a title the author already wrote). */
export function chapterTitleNeeded(title: string, blocks: Block[]): boolean {
  const first = blocks[0];
  if (!first || first.kind !== "heading") return true;
  const text = first.runs.map((run) => run.text).join("");
  return normalizeTitle(text) !== normalizeTitle(title);
}

/** Split runs into segments at soft line breaks ("\n" inside text).
 * Each segment keeps its run flags; empty segments are kept and filtered
 * by the caller via hasText. Pure and unit-tested. */
export function splitOnBreaks(runs: TextRunModel[]): TextRunModel[][] {
  const segments: TextRunModel[][] = [[]];
  const pushText = (run: TextRunModel, text: string) => {
    if (text) segments[segments.length - 1].push({ ...run, text });
  };
  for (const run of runs) {
    const parts = run.text.split("\n");
    pushText(run, parts[0]);
    for (const part of parts.slice(1)) {
      segments.push([]);
      pushText(run, part);
    }
  }
  return segments;
}
