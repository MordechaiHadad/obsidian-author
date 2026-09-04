import JSZip from "jszip";
import type { Block, TextRunModel } from "./text.ts";

export interface EpubOptions {
  /** Book/chapter title (note name). */
  title: string;
  /** Sanitized CSS indent, e.g. "2em". */
  indent: string;
  /** Sanitized CSS line height, e.g. "1.7". */
  lineHeight: string;
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

function blocksToXhtml(blocks: Block[]): string {
  return blocks
    .map((block) => {
      switch (block.kind) {
        case "heading":
          return `<h${block.level}>${
            runsToXhtml(block.runs)
          }</h${block.level}>`;
        case "paragraph":
          return `<p>${runsToXhtml(block.runs)}</p>`;
        case "list": {
          const tag = block.ordered ? "ol" : "ul";
          const items = block.items
            .map((runs) => `<li>${runsToXhtml(runs)}</li>`)
            .join("");
          return `<${tag}>${items}</${tag}>`;
        }
        case "break":
          return `<p class="scene">* * *</p>`;
      }
    })
    .join("\n");
}

function stylesheet(indent: string, lineHeight: string): string {
  return [
    `p { text-indent: ${indent}; margin: 0; line-height: ${lineHeight}; }`,
    `body > p:first-of-type { text-indent: 0; }`,
    `.scene { text-indent: 0; text-align: center; margin: 1em 0; }`,
    `h1, h2, h3, h4, h5, h6 { line-height: 1.3; }`,
    ``,
  ].join("\n");
}

/** Build an EPUB3 buffer from blocks: one chapter, our own indent stylesheet. */
export async function blocksToEpubBuffer(
  blocks: Block[],
  opts: EpubOptions,
): Promise<ArrayBuffer> {
  const title = escapeXml(opts.title);
  const body = blocksToXhtml(blocks);
  const modified = new Date().toISOString().split(".")[0] + "Z";

  const container = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n` +
    `  <rootfiles>\n` +
    `    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>\n` +
    `  </rootfiles>\n</container>\n`;

  const opf = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="bookid" version="3.0">\n` +
    `  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">\n` +
    `    <dc:identifier id="bookid">urn:obsidian-author:${title}</dc:identifier>\n` +
    `    <dc:title>${title}</dc:title>\n` +
    `    <dc:language>en</dc:language>\n` +
    `    <meta property="dcterms:modified">${modified}</meta>\n` +
    `  </metadata>\n` +
    `  <manifest>\n` +
    `    <item id="chapter" href="chapter.xhtml" media-type="application/xhtml+xml"/>\n` +
    `    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>\n` +
    `    <item id="style" href="style.css" media-type="text/css"/>\n` +
    `  </manifest>\n` +
    `  <spine>\n` +
    `    <itemref idref="chapter"/>\n` +
    `  </spine>\n</package>\n`;

  const chapter = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<!DOCTYPE html>\n` +
    `<html xmlns="http://www.w3.org/1999/xhtml">\n` +
    `<head><title>${title}</title><link rel="stylesheet" type="text/css" href="style.css"/></head>\n` +
    `<body>\n${body}\n</body>\n</html>\n`;

  const nav = `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<!DOCTYPE html>\n` +
    `<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">\n` +
    `<head><title>${title}</title></head>\n` +
    `<body><nav epub:type="toc"><ol><li><a href="chapter.xhtml">${title}</a></li></ol></nav></body>\n` +
    `</html>\n`;

  // mimetype must be the first entry, stored uncompressed.
  const zip = new JSZip();
  zip.file("mimetype", "application/epub+zip", { compression: "STORE" });
  zip.file("META-INF/container.xml", container);
  zip.file("OEBPS/content.opf", opf);
  zip.file("OEBPS/chapter.xhtml", chapter);
  zip.file("OEBPS/nav.xhtml", nav);
  zip.file("OEBPS/style.css", stylesheet(opts.indent, opts.lineHeight));

  const out = await zip.generateAsync({
    type: "uint8array",
    mimeType: "application/epub+zip",
  });
  return out.slice().buffer as ArrayBuffer;
}
