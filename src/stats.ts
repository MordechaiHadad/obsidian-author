/** Word → print-page estimate. Dependency-free on purpose: importable
 * headlessly (unit tests) without the Obsidian API. */

export const WORDS_PER_PAGE = 250;

/** Strip frontmatter and fenced code blocks so they don't inflate the count. */
function stripNonProse(text: string): string {
  let out = text ?? "";
  out = out.replace(/^---\n[\s\S]*?\n---\n?/, "");
  out = out.replace(/```[\s\S]*?(?:```|$)/g, " ");
  out = out.replace(/~~~[\s\S]*?(?:~~~|$)/g, " ");
  return out;
}

/** Count words: runs of letters/numbers (apostrophes inside words count once).
 * Markdown syntax (#, *, >, -, links) contributes no words by construction. */
export function countWords(text: string): number {
  const clean = stripNonProse(text);
  const matches = clean.match(/[\p{L}\p{N}]+(?:['’][\p{L}\p{N}]+)*/gu);
  return matches ? matches.length : 0;
}

/** Standard print pages: 250 words per page, rounded up. Empty note → 0. */
export function wordsToPages(words: number): number {
  if (!Number.isFinite(words) || words <= 0) return 0;
  return Math.max(1, Math.ceil(words / WORDS_PER_PAGE));
}

/** "~48 print pages" / "~1 print page" / "0 print pages". */
export function formatPrintPages(pages: number): string {
  if (!Number.isFinite(pages) || pages <= 0) return "0 print pages";
  const n = Math.floor(pages);
  if (n <= 1) return "~1 print page";
  return `~${n.toLocaleString("en-US")} print pages`;
}
