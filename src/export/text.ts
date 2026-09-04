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
