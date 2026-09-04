/// <reference lib="dom" />
import { App, Component, MarkdownRenderer, TFile } from "obsidian";

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

const blankRun = (text: string): TextRunModel => ({
  text,
  bold: false,
  italic: false,
  strike: false,
  code: false,
});

/** Render a note's markdown through Obsidian and extract a clean block model.
 * Frontmatter is stripped first; whitespace-only paragraphs are dropped. */
export async function noteToBlocks(
  app: App,
  file: TFile,
  content: string,
): Promise<Block[]> {
  const stripped = content.replace(/^---\n[\s\S]*?\n---\n?/, "");
  const host = document.createElement("div");
  const component = new Component();
  component.load();
  try {
    await MarkdownRenderer.render(app, stripped, host, file.path, component);
    return extractBlocks(host);
  } finally {
    component.unload();
  }
}

function extractBlocks(host: HTMLElement): Block[] {
  const blocks: Block[] = [];
  for (const child of Array.from(host.children)) {
    const el = child as HTMLElement;
    const tag = el.tagName;
    if (/^H[1-6]$/.test(tag)) {
      const runs = extractRuns(el);
      if (hasText(runs)) {
        blocks.push({
          kind: "heading",
          level: Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6,
          runs,
        });
      }
    } else if (tag === "P") {
      const runs = extractRuns(el);
      if (hasText(runs)) blocks.push({ kind: "paragraph", runs });
    } else if (tag === "UL" || tag === "OL") {
      const items: TextRunModel[][] = [];
      for (const li of Array.from(el.querySelectorAll(":scope > li"))) {
        const runs = extractRuns(li as HTMLElement);
        if (hasText(runs)) items.push(runs);
      }
      if (items.length > 0) {
        blocks.push({ kind: "list", ordered: tag === "OL", items });
      }
    } else if (tag === "HR") {
      blocks.push({ kind: "break" });
    } else if (tag === "PRE") {
      // Code block: one (monospace) paragraph per line.
      for (const line of (el.textContent ?? "").split("\n")) {
        if (line.trim()) {
          blocks.push({
            kind: "paragraph",
            runs: [{ ...blankRun(line), code: true }],
          });
        }
      }
    } else if (tag === "BLOCKQUOTE") {
      // Quotes flatten to plain paragraphs in v1.
      for (const p of Array.from(el.querySelectorAll("p"))) {
        const runs = extractRuns(p as HTMLElement);
        if (hasText(runs)) blocks.push({ kind: "paragraph", runs });
      }
    }
    // Tables, math, embeds and anything else are skipped in v1.
  }
  return blocks;
}

function hasText(runs: TextRunModel[]): boolean {
  return runs.some((r) => r.text.trim() !== "");
}

/** Collect inline runs, merging child formatting flags downwards. */
function extractRuns(
  el: HTMLElement,
  inherited?: Omit<TextRunModel, "text">,
): TextRunModel[] {
  const base = inherited ?? {
    bold: false,
    italic: false,
    strike: false,
    code: false,
  };
  const runs: TextRunModel[] = [];
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      runs.push({ ...base, text: node.textContent ?? "" });
    } else if (node instanceof HTMLElement) {
      if (node.matches(".math, script, style")) continue;
      const tag = node.tagName;
      const flags = {
        bold: base.bold || tag === "STRONG" || tag === "B",
        italic: base.italic || tag === "EM" || tag === "I",
        strike: base.strike || tag === "DEL" || tag === "S",
        code: base.code || tag === "CODE",
      };
      if (tag === "BR") {
        runs.push({ ...base, text: " " });
      } else if (tag === "IMG") {
        continue; // Images are v1.1.
      } else {
        runs.push(...extractRuns(node, flags));
      }
    }
  }
  return runs;
}
