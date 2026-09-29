import JSZip from "jszip";
import { buildEpubCss } from "../scope.ts";
import type { Block, Chapter, TextRunModel } from "./text.ts";

export interface EpubOptions {
  /** Book/chapter title (note name). */
  title: string;
  /** Sanitized CSS indent, e.g. "2em". */
  indent: string;
  /** Sanitized CSS line height, e.g. "1.7". */
  lineHeight: string;
  /** When false, no first-line indent anywhere. Defaults to true. */
  enableIndent?: boolean;
  /** Flush-left paragraph after heading/break. Defaults to true. */
  flushAfterHeading?: boolean;
  /** Add a two-line drop cap to the first body paragraph. Defaults to false. */
  enableDropCap?: boolean;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runsToXhtml(runs: TextRunModel[]): string {
  return runs
    .map((r) => {
      let text = escapeXml(r.text);
      if (r.code) text = `<code>${text}</code>`;
      if (r.strike) text = `<s>${text}</s>`;
      if (r.italic) text = `<em>${text}</em>`;
      if (r.bold) text = `<strong>${text}</strong>`;
      return text;
    })
    .join("");
}

function blocksToXhtml(
  blocks: Block[],
  opts?: Pick<
    EpubOptions,
    "enableIndent" | "flushAfterHeading" | "enableDropCap"
  >,
): string {
  const enableIndent = opts?.enableIndent ?? true;
  const flushAfterHeading = opts?.flushAfterHeading ?? true;
  const enableDropCap = opts?.enableDropCap ?? false;
  // Manuscript convention (mirrors DOCX indentedYet): the first body
  // paragraph is flush left; headings/lists/breaks don't consume it.
  let seenFirstParagraph = false;
  let afterHeadingOrBreak = false;
  return blocks
    .map((block) => {
      switch (block.kind) {
        case "heading":
          afterHeadingOrBreak = true;
          return `<h${block.level}>${
            runsToXhtml(block.runs)
          }</h${block.level}>`;
        case "paragraph": {
          const isFirstParagraph = !seenFirstParagraph;
          const flush = !enableIndent || !seenFirstParagraph ||
            (flushAfterHeading && afterHeadingOrBreak);
          seenFirstParagraph = true;
          afterHeadingOrBreak = false;
          const classes = [
            flush && "flush",
            enableDropCap && isFirstParagraph && "dropcap",
          ].filter(Boolean).join(" ");
          return classes
            ? `<p class="${classes}">${runsToXhtml(block.runs)}</p>`
            : `<p>${runsToXhtml(block.runs)}</p>`;
        }
        case "list": {
          afterHeadingOrBreak = false;
          const tag = block.ordered ? "ol" : "ul";
          const items = block.items
            .map((runs) => `<li>${runsToXhtml(runs)}</li>`)
            .join("");
          return `<${tag}>${items}</${tag}>`;
        }
        case "break":
          afterHeadingOrBreak = true;
          return `<p class="scene">* * *</p>`;
      }
    })
    .join("\n");
}

/** Build an EPUB3 buffer from one note's blocks (chapter export). */
export async function blocksToEpubBuffer(
  blocks: Block[],
  opts: EpubOptions,
): Promise<ArrayBuffer> {
  return await chaptersToEpubBuffer([{ title: opts.title, blocks }], opts);
}

/** File name of chapter `index` inside OEBPS/. The first one keeps the
 * plain `chapter.xhtml` name (single-chapter archives stay identical). */
function chapterHref(index: number): string {
  return index === 0 ? "chapter.xhtml" : `chapter-${index + 1}.xhtml`;
}

/** Build an EPUB3 buffer from chapters (novel export): one xhtml file per
 * chapter with its own flush/drop-cap state, a manifest + spine entry each,
 * and a nav TOC listing every chapter. `opts.title` is the book title.
 * The stylesheet comes from buildEpubCss() in scope.ts — the single source
 * of truth for generated CSS shared with the print/PDF path. */
export async function chaptersToEpubBuffer(
  chapters: Chapter[],
  opts: EpubOptions,
): Promise<ArrayBuffer> {
  const title = escapeXml(opts.title);
  const contents = chapters.map((chapter, index) => ({
    id: index === 0 ? "chapter" : `chapter-${index + 1}`,
    href: chapterHref(index),
    title: escapeXml(chapter.title || opts.title),
    body: blocksToXhtml(chapter.blocks, opts),
  }));
  const modified = new Date().toISOString().split(".")[0] + "Z";

  const container = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n` +
    `  <rootfiles>\n` +
    `    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>\n` +
    `  </rootfiles>\n</container>\n`;

  const manifest = contents
    .map((c) =>
      `    <item id="${c.id}" href="${c.href}" media-type="application/xhtml+xml"/>`
    )
    .join("\n");
  const spine = contents
    .map((c) => `    <itemref idref="${c.id}"/>`)
    .join("\n");

  const opf = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="bookid" version="3.0">\n` +
    `  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">\n` +
    `    <dc:identifier id="bookid">urn:obsidian-author:${title}</dc:identifier>\n` +
    `    <dc:title>${title}</dc:title>\n` +
    `    <dc:language>en</dc:language>\n` +
    `    <meta property="dcterms:modified">${modified}</meta>\n` +
    `  </metadata>\n` +
    `  <manifest>\n` +
    `${manifest}\n` +
    `    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>\n` +
    `    <item id="style" href="style.css" media-type="text/css"/>\n` +
    `  </manifest>\n` +
    `  <spine>\n` +
    `${spine}\n` +
    `  </spine>\n</package>\n`;

  const xhtmlFiles = contents.map((c) => ({
    href: c.href,
    content: `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<!DOCTYPE html>\n` +
      `<html xmlns="http://www.w3.org/1999/xhtml">\n` +
      `<head><title>${c.title}</title><link rel="stylesheet" type="text/css" href="style.css"/></head>\n` +
      `<body>\n${c.body}\n</body>\n</html>\n`,
  }));

  const navItems = contents
    .map((c) => `<li><a href="${c.href}">${c.title}</a></li>`)
    .join("");
  const nav = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<!DOCTYPE html>\n` +
    `<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">\n` +
    `<head><title>${title}</title></head>\n` +
    `<body><nav epub:type="toc"><ol>${navItems}</ol></nav></body>\n` +
    `</html>\n`;

  // mimetype must be the first entry, stored uncompressed.
  const zip = new JSZip();
  zip.file("mimetype", "application/epub+zip", { compression: "STORE" });
  zip.file("META-INF/container.xml", container);
  zip.file("OEBPS/content.opf", opf);
  for (const file of xhtmlFiles) zip.file(`OEBPS/${file.href}`, file.content);
  zip.file("OEBPS/nav.xhtml", nav);
  zip.file(
    "OEBPS/style.css",
    buildEpubCss(
      opts.indent,
      opts.lineHeight,
      opts.enableIndent ?? true,
      opts.enableDropCap ?? false,
    ),
  );

  const out = await zip.generateAsync({
    type: "uint8array",
    mimeType: "application/epub+zip",
  });
  return out.slice().buffer as ArrayBuffer;
}
