/// <reference lib="dom" />
import { App, Component, MarkdownRenderer, TFile } from "obsidian";
import type { Block, TextRunModel } from "./text.ts";
import { splitOnBreaks } from "./text.ts";

const blankRun = (text: string): TextRunModel => ({
  text,
  bold: false,
  italic: false,
  strike: false,
  code: false,
});

/** Render a note's markdown through Obsidian and extract a clean block model.
 * Frontmatter is stripped first; whitespace-only paragraphs are dropped.
 * Manuscript convention: every source line is its own paragraph, so rendered
 * paragraphs are additionally split on soft breaks (literal newlines in text
 * and <br> elements). This deliberately deviates from strict Markdown, where
 * single newlines stay inside one paragraph. */
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
      const runs = flattenBreaks(extractRuns(el));
      if (hasText(runs)) {
        blocks.push({
          kind: "heading",
          level: Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6,
          runs,
        });
      }
    } else if (tag === "P") {
      for (const runs of splitOnBreaks(extractRuns(el))) {
        if (hasText(runs)) blocks.push({ kind: "paragraph", runs });
      }
    } else if (tag === "UL" || tag === "OL") {
      const items: TextRunModel[][] = [];
      for (const li of Array.from(el.querySelectorAll(":scope > li"))) {
        const runs = flattenBreaks(extractRuns(li as HTMLElement));
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
        for (const runs of splitOnBreaks(extractRuns(p as HTMLElement))) {
          if (hasText(runs)) blocks.push({ kind: "paragraph", runs });
        }
      }
    }
    // Tables, math, embeds and anything else are skipped in v1.
  }
  return blocks;
}

function hasText(runs: TextRunModel[]): boolean {
  return runs.some((r) => r.text.trim() !== "");
}

/** Collapse soft breaks to spaces (headings, list items). */
function flattenBreaks(runs: TextRunModel[]): TextRunModel[] {
  return runs.map((r) =>
    r.text.includes("\n") ? { ...r, text: r.text.replace(/\n/g, " ") } : r
  );
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
      if (tag === "BR") runs.push({ ...base, text: "\n" });
      else if (tag === "IMG") continue; // Images are v1.1.
      else runs.push(...extractRuns(node, flags));
    }
  }
  return runs;
}
