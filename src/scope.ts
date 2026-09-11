/// <reference lib="dom" />
/** Folder scoping + manuscript DOM marking + print-CSS builder.
 * Dependency-free on purpose: importable headlessly (unit tests) without
 * the Obsidian API. DOM helpers need a DOM at runtime (browser/Obsidian). */

/** Per-section marker: lives on rendered content nodes (not leaf chrome). */
export const PP_CLASS = "author-pp";
/** First paragraph of the note: always flush left. */
export const PP_FIRST_CLASS = "author-pp-first";
/** Added per section only when the flush-after-heading toggle is on. */
export const PP_FLUSH_CLASS = "author-pp-flush";

export interface ManuscriptMarkOptions {
  indent: string;
  lineHeight: string;
  flushAfterHeading: boolean;
}

/** Normalize "Novels/", "/Novels", " novels " -> "Novels". "" = disabled. */
export function normalizeFolder(input: string | null | undefined): string {
  return (input ?? "")
    .trim()
    .replace(/^\/+|\/+$/g, "")
    .replace(/\/{2,}/g, "/");
}

/** Split a sibling-node list at <br> elements (any depth), preserving
 * inline formatting ancestors by cloning wrappers per segment. Nodes are
 * moved, not cloned, and <br> markers are dropped. Needs a DOM at runtime;
 * kept here (not main.ts) so it stays importable without the Obsidian API. */
export function collectInlineSegments(nodes: Node[]): Node[][] {
  const segments: Node[][] = [[]];
  const current = (): Node[] => segments[segments.length - 1];
  for (const node of nodes) {
    if (node instanceof HTMLElement && node.tagName === "BR") {
      segments.push([]);
    } else if (node instanceof HTMLElement && node.querySelector("br")) {
      const inner = collectInlineSegments(Array.from(node.childNodes));
      inner.forEach((piece, idx) => {
        const wrapper = node.cloneNode(false) as HTMLElement;
        wrapper.append(...piece);
        if (idx === 0) current().push(wrapper);
        else segments.push([wrapper]);
      });
    } else {
      current().push(node);
    }
  }
  return segments;
}

/** True when a segment holds visible content (prose or an image). */
export function segmentHasContent(seg: Node[]): boolean {
  return seg.some((n) => {
    if (n instanceof HTMLElement && n.querySelector("img")) return true;
    return (n.textContent ?? "").trim() !== "";
  });
}

/** Split one rendered <p> at soft breaks into sibling <p> elements, one
 * per source line. Moves nodes; inline ancestors are cloned per segment.
 * No-op when there are no breaks. Idempotent. */
export function splitElementParagraph(p: HTMLParagraphElement): void {
  if (!p.querySelector("br")) return;
  const segments = collectInlineSegments(Array.from(p.childNodes))
    .filter(segmentHasContent);
  if (segments.length < 2) return;
  const parent = p.parentElement;
  if (!parent) return;
  const doc = p.ownerDocument;
  p.replaceChildren(...segments[0]);
  let anchor: Node = p;
  for (const seg of segments.slice(1)) {
    const next = doc.createElement("p");
    next.replaceChildren(...seg);
    parent.insertBefore(next, anchor.nextSibling);
    anchor = next;
  }
}

/** Tag one rendered section (post-processor `el`, or a top-level child of
 * a hidden render root): marker classes, per-section CSS variables, first-
 * paragraph flush, and soft-break splitting. Idempotent. */
export function markManuscriptSection(
  el: HTMLElement,
  isFirst: boolean,
  opts: ManuscriptMarkOptions,
): void {
  el.classList.add(PP_CLASS);
  if (opts.flushAfterHeading) el.classList.add(PP_FLUSH_CLASS);
  el.style.setProperty("--author-indent", opts.indent);
  el.style.setProperty("--author-line-height", opts.lineHeight);
  if (isFirst) el.classList.add(PP_FIRST_CLASS);
  const target = el.tagName === "P"
    ? el as HTMLParagraphElement
    : el.querySelector(":scope > p");
  if (target instanceof HTMLParagraphElement) splitElementParagraph(target);
}

/** Tag every top-level section of a rendered container (e.g. a hidden
 * `MarkdownRenderer.render` host). First *paragraph block* gets the flush
 * marker; headings before it don't consume it. */
export function markManuscriptRoot(
  container: HTMLElement,
  opts: ManuscriptMarkOptions,
): void {
  let firstSeen = false;
  for (const child of Array.from(container.children)) {
    if (!(child instanceof HTMLElement)) continue;
    const isFirst = !firstSeen && child.matches("div.el-p");
    if (child.matches("div.el-p")) firstSeen = true;
    markManuscriptSection(child, isFirst, opts);
  }
}

/** True when a vault-relative path is a Markdown note inside the folder.
 * Subfolders included. `folder` must already be normalized (may be ""). */
export function isManuscriptPath(
  filePath: string | null | undefined,
  folder: string,
): boolean {
  if (!filePath || !folder) return false;
  if (!filePath.endsWith(".md")) return false;
  return filePath.startsWith(folder + "/");
}

/** Build a self-contained `@media print` stylesheet with literal values
 * (no CSS variables, no dependency on JS-added leaf classes).
 * The `author-pp` marker class is added per rendered section by a
 * `registerMarkdownPostProcessor` hook, which — unlike leaf-container
 * classes — lives on content nodes that the PDF print render keeps.
 * `!important` beats theme print resets.
 * NOTE: this and buildEpubCss are the single source of truth for
 * generated CSS. styles.css (screen, CSS variables) must stay in sync
 * with them by hand — it can't import from here. */
export function buildPrintCss(
  indent: string,
  lineHeight: string,
  flushAfterHeading: boolean,
  enableIndent = true,
): string {
  const bodyIndent = enableIndent ? indent : "0";
  const lines = [
    "@media print {",
    `  .author-pp p { text-indent: ${bodyIndent} !important; margin-block-start: 0 !important; margin-block-end: 0 !important; line-height: ${lineHeight} !important; }`,
    `  p.author-pp { text-indent: ${bodyIndent} !important; }`,
    flushAfterHeading
      ? `  .author-pp-first > p:first-child, h1 + .author-pp-flush > p:first-child, h2 + .author-pp-flush > p:first-child, h3 + .author-pp-flush > p:first-child, h4 + .author-pp-flush > p:first-child, h5 + .author-pp-flush > p:first-child, h6 + .author-pp-flush > p:first-child, hr + .author-pp-flush > p:first-child { text-indent: 0 !important; }`
      : `  .author-pp-first > p:first-child { text-indent: 0 !important; }`,
    `  li.author-pp p, li .author-pp p, blockquote.author-pp p, blockquote .author-pp p, table .author-pp p, pre .author-pp p { text-indent: 0 !important; }`,
    "}",
  ];
  return lines.join("\n");
}

/** Build the standalone EPUB stylesheet with literal values.
 * Lives here (not epub.ts) so both generated stylesheets share one home:
 * fix spacing/indent here once and every export target picks it up.
 * Tight manuscript rhythm — the indent (not vertical gaps) separates
 * paragraphs. !important + block-start/end + padding beat reader UA
 * defaults (e.g. p { margin: 1em 0 }) that otherwise double-space
 * single-newline source lines. */
export function buildEpubCss(
  indent: string,
  lineHeight: string,
  enableIndent = true,
): string {
  const bodyIndent = enableIndent ? indent : "0";
  return [
    `p { text-indent: ${bodyIndent} !important; margin: 0 !important; padding: 0; margin-block-start: 0 !important; margin-block-end: 0 !important; line-height: ${lineHeight}; }`,
    `p.flush { text-indent: 0 !important; }`,
    `.scene { text-indent: 0 !important; text-align: center; margin: 1em 0 !important; }`,
    `h1, h2, h3, h4, h5, h6 { line-height: 1.3; margin: 1em 0 0.5em; font-weight: bold; }`,
    `ol, ul { margin: 0; padding-left: 1.5em; }`,
    `li { text-indent: 0; margin: 0; padding: 0; }`,
    ``,
  ].join("\n");
}
