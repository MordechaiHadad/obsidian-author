import JSZip from "jszip";
import { blocksToDocxBuffer } from "../src/export/docx.ts";
import { blocksToEpubBuffer } from "../src/export/epub.ts";
import type { Block } from "../src/export/model.ts";

const blocks: Block[] = [
  {
    kind: "heading",
    level: 1,
    runs: [{
      text: "Chapter One",
      bold: false,
      italic: false,
      strike: false,
      code: false,
    }],
  },
  {
    kind: "paragraph",
    runs: [
      {
        text: "It was a ",
        bold: false,
        italic: false,
        strike: false,
        code: false,
      },
      { text: "dark", bold: true, italic: false, strike: false, code: false },
      { text: " and ", bold: false, italic: false, strike: false, code: false },
      { text: "stormy", bold: false, italic: true, strike: false, code: false },
      {
        text: " night with <xml> & trouble.",
        bold: false,
        italic: false,
        strike: false,
        code: false,
      },
    ],
  },
  {
    kind: "paragraph",
    runs: [{
      text: "Second paragraph here.",
      bold: false,
      italic: false,
      strike: false,
      code: false,
    }],
  },
  {
    kind: "list",
    ordered: false,
    items: [[{
      text: "an item",
      bold: false,
      italic: false,
      strike: false,
      code: false,
    }]],
  },
  { kind: "break" },
];

let failures = 0;
const check = (name: string, cond: boolean) => {
  console.log(`${cond ? "PASS" : "FAIL"} ${name}`);
  if (!cond) failures++;
};

// --- DOCX ---
const docxBuf = await blocksToDocxBuffer(blocks, "2em");
const docxZip = await JSZip.loadAsync(docxBuf);
const docXml = await docxZip.file("word/document.xml")?.async("string") ?? "";
check("docx: document.xml exists", docXml.length > 0);
check(
  "docx: first-line indent 480 twips (2em @12pt)",
  docXml.includes('w:firstLine="480"'),
);
check(
  "docx: first paragraph flush (val=0)",
  docXml.includes('w:firstLine="0"'),
);
check("docx: heading style present", docXml.includes('w:val="Heading1"'));
check("docx: bold run present", docXml.includes("<w:b/>"));
check("docx: list numbering present", docXml.includes("<w:numPr>"));
check("docx: scene break centered", docXml.includes("* * *"));

// --- EPUB ---
const epubBuf = await blocksToEpubBuffer(blocks, {
  title: "Test & <Chapter>",
  indent: "2em",
  lineHeight: "1.7",
});
const epubZip = await JSZip.loadAsync(epubBuf);
const names = Object.keys(epubZip.files);
check("epub: mimetype is first entry", names[0] === "mimetype");
const raw = new Uint8Array(epubBuf);
check(
  "epub: mimetype stored uncompressed (method 0x0000)",
  raw[8] === 0 && raw[9] === 0,
);
const opf = await epubZip.file("OEBPS/content.opf")?.async("string") ?? "";
check(
  "epub: opf has title",
  opf.includes("<dc:title>Test &amp; &lt;Chapter&gt;</dc:title>"),
);
const chapter = await epubZip.file("OEBPS/chapter.xhtml")?.async("string") ??
  "";
check("epub: chapter has h1", chapter.includes("<h1>Chapter One</h1>"));
check(
  "epub: body text escaped",
  chapter.includes("night with &lt;xml&gt; &amp; trouble."),
);
check(
  "epub: bold/italic runs",
  chapter.includes("<strong>dark</strong>") &&
    chapter.includes("<em>stormy</em>"),
);
check("epub: scene break", chapter.includes('<p class="scene">* * *</p>'));
const css = await epubZip.file("OEBPS/style.css")?.async("string") ?? "";
check("epub: css has 2em indent", css.includes("text-indent: 2em;"));
check("epub: nav exists", !!epubZip.file("OEBPS/nav.xhtml"));

if (failures > 0) {
  throw new Error(`${failures} export test(s) failed`);
}
console.log("ALL EXPORT TESTS PASS");
