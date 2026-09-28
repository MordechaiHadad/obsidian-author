/// <reference lib="dom" />
/** Folder scoping + manuscript DOM marking + print-CSS builder.
 * Dependency-free on purpose: importable headlessly (unit tests) without
 * the Obsidian API. DOM helpers need a DOM at runtime (browser/Obsidian). */

/** Per-section marker: lives on rendered content nodes (not leaf chrome). */
export const PP_CLASS = "author-pp";
/** First paragraph of the note: always flush left. */
export const PP_FIRST_CLASS = "author-pp-first";
/** Direct first-paragraph marker for isolated PDF renders. */
export const PP_FIRST_PARAGRAPH_CLASS = "author-pp-first-paragraph";
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
    const isParagraph = child.matches("div.el-p, p");
    const isFirst = !firstSeen && isParagraph;
    if (isParagraph) firstSeen = true;
    markManuscriptSection(child, isFirst, opts);
  }
  // Also mark the actual prose paragraph, independently of whether the
  // renderer wrapped it in a top-level `div.el-p` (isolated PDF renders can
  // differ from Reading view's DOM structure).
  const firstProseParagraph = Array.from(container.querySelectorAll("p"))
    .find((p) => !p.closest("li, blockquote, table, pre, .callout, .footnote"));
  firstProseParagraph?.classList.add(PP_FIRST_PARAGRAPH_CLASS);
}

/** True when a reading-view sibling is frontmatter/properties chrome rather
 * than content. The Markdown post-processor only sees content sections, but
 * the first `div.el-p` can have such a node as `previousElementSibling`
 * when the note has frontmatter — it must still count as "first". */
export function isFrontmatterContainer(el: Element | null): boolean {
  if (!(el instanceof HTMLElement)) return false;
  const cls = Array.from(el.classList).join(" ").toLowerCase();
  return cls.includes("frontmatter") || cls.includes("metadata") ||
    cls.includes("yaml") || cls.includes("properties");
}

/** True when `el` is the first paragraph section of the note: a `div.el-p`
 * with no previous paragraph/heading/rule/list/quote/table/pre sibling.
 * Frontmatter/properties containers are skipped. */
export function isFirstParagraphSection(el: HTMLElement): boolean {
  if (!el.matches("div.el-p")) return false;
  let sib: Element | null = el.previousElementSibling;
  while (sib) {
    if (isFrontmatterContainer(sib)) {
      sib = sib.previousElementSibling;
      continue;
    }
    // Any other rendered sibling means we are not the first section.
    // (Headings/rules before the first paragraph don't consume the
    // manuscript "first paragraph flush" — markManuscriptRoot handles that
    // case separately — but in the streaming post-processor the safest
    // signal for "top of note" is no previous content sibling at all.)
    return false;
  }
  return true;
}

/** Reconcile first-paragraph markers within one rendered Reading view.
 * The Markdown post-processor can fire while `el` is still detached (bulk
 * render), so `isFirstParagraphSection` may mark several `div.el-p` as
 * first. The rendered document order is authoritative: keep the marker on
 * the first match, drop it from the rest. Idempotent. */
export function reconcileFirstParagraphMarkers(preview: HTMLElement): void {
  const marked = Array.from(
    preview.querySelectorAll(`div.el-p.${PP_FIRST_CLASS}`),
  );
  for (const el of marked.slice(1)) {
    if (el instanceof HTMLElement) el.classList.remove(PP_FIRST_CLASS);
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

/** Flush-left selectors for the first paragraph of the note. The manuscript
 * marker can land either on a wrapper section (Reading view: `div.el-p
 * .author-pp-first > p`) or — as the isolated PDF render does — on the
 * `<p>` itself. Both shapes must match: the drop cap already has a
 * marker-on-the-paragraph selector, so without one here the cap renders
 * while `text-indent` keeps its first line indented. */
const FIRST_FLUSH_SELECTORS = [
  ".author-pp-first > p:first-child",
  "p.author-pp-first",
  ".author-pp p.author-pp-first-paragraph",
  "p.author-pp-first-paragraph",
].join(", ");

/** Same two shapes for the paragraph after a heading/rule: the marker is
 * either the wrapper the heading is a sibling of, or the `<p>` itself. */
const HEADING_FLUSH_SELECTORS = ["h1", "h2", "h3", "h4", "h5", "h6", "hr"]
  .map((tag) =>
    `${tag} + .author-pp-flush > p:first-child, ${tag} + p.author-pp-flush`
  )
  .join(", ");

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
  enableDropCap = false,
): string {
  const bodyIndent = enableIndent ? indent : "0";
  const lines = [
    "@media print {",
    `  .author-pp p { text-indent: ${bodyIndent} !important; margin-block-start: 0 !important; margin-block-end: 0 !important; line-height: ${lineHeight} !important; }`,
    `  p.author-pp { text-indent: ${bodyIndent} !important; }`,
    enableDropCap
      ? `  .author-pp-first > p:first-child::first-letter, .author-pp-first-paragraph::first-letter { float: left !important; font-size: 3em !important; line-height: 1 !important; padding-right: 0.1em !important; font-weight: 600 !important; }`
      : "",
    flushAfterHeading
      ? `  ${FIRST_FLUSH_SELECTORS}, ${HEADING_FLUSH_SELECTORS} { text-indent: 0 !important; }`
      : `  ${FIRST_FLUSH_SELECTORS} { text-indent: 0 !important; }`,
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
  enableDropCap = false,
): string {
  const bodyIndent = enableIndent ? indent : "0";
  return [
    `p { text-indent: ${bodyIndent} !important; margin: 0 !important; padding: 0; margin-block-start: 0 !important; margin-block-end: 0 !important; line-height: ${lineHeight}; }`,
    `p.flush { text-indent: 0 !important; }`,
    ...(enableDropCap
      ? [
        `p.dropcap::first-letter { float: left !important; font-size: 3em !important; line-height: 1 !important; padding-right: 0.1em !important; font-weight: 600 !important; }`,
      ]
      : []),
    `.scene { text-indent: 0 !important; text-align: center; margin: 1em 0 !important; }`,
    `h1, h2, h3, h4, h5, h6 { line-height: 1.3; margin: 1em 0 0.5em; font-weight: bold; }`,
    `ol, ul { margin: 0; padding-left: 1.5em; }`,
    `li { text-indent: 0; margin: 0; padding: 0; }`,
    ``,
  ].join("\n");
}
