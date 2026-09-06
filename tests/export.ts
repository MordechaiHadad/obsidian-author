import JSZip from "jszip";
import { blocksToDocxBuffer, lengthToTwips } from "../src/export/docx.ts";
import { blocksToEpubBuffer } from "../src/export/epub.ts";
import {
  countWords,
  formatPrintPages,
  WORDS_PER_PAGE,
  wordsToPages,
} from "../src/stats.ts";
import type { Block, TextRunModel } from "../src/export/text.ts";
import { splitOnBreaks } from "../src/export/text.ts";

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
const docxBuf = await blocksToDocxBuffer(blocks, "2em", "1.7");
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
check("docx: no gap between paragraphs", docXml.includes('w:after="0"'));
check(
  "docx: line height 1.7 mapped (w:line=408)",
  docXml.includes('w:line="408"'),
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

// --- Pure units ---
check("twips: 2em -> 480", lengthToTwips("2em") === 480);
check("twips: bare number -> em", lengthToTwips("1.5") === 360);
check("twips: px @96dpi", lengthToTwips("32px") === 480);
check("twips: garbage -> default 480", lengthToTwips("banana") === 480);

const plain = (text: string): TextRunModel => ({
  text,
  bold: false,
  italic: false,
  strike: false,
  code: false,
});
const split = splitOnBreaks([
  plain("one\ntwo"),
  { ...plain("three\nfour"), bold: true },
]);
check("split: 3 segments", split.length === 3);
check(
  "split: flags survive breaks",
  split[1][0].text === "two" && split[1][1].text === "three" &&
    split[1][1].bold && split[2][0].text === "four" && split[2][0].bold,
);
check(
  "split: no breaks untouched",
  splitOnBreaks([plain("abc")]).length === 1,
);

// --- Print pages (250 words = 1 standard page) ---
check("pages: standard is 250 wpp", WORDS_PER_PAGE === 250);
check("words: empty -> 0", countWords("") === 0);
check(
  "words: basic count",
  countWords("It was a dark and stormy night.") === 7,
);
check(
  "words: markdown syntax adds nothing",
  countWords("# Chapter\n\nHello *world* — ok.") === 4,
);
check(
  "words: frontmatter ignored",
  countWords("---\ntitle: Foo\n---\nHello world") === 2,
);
check(
  "words: fenced code ignored",
  countWords("Hello\n```js\nconst a = 1;\n```\nworld") === 2,
);
check("pages: 0 words -> 0", wordsToPages(0) === 0);
check("pages: 1 word -> 1", wordsToPages(1) === 1);
check("pages: 250 words -> 1", wordsToPages(250) === 1);
check("pages: 251 words -> 2", wordsToPages(251) === 2);
check("pages: 500 words -> 2", wordsToPages(500) === 2);
check("format: 0 -> 0 print pages", formatPrintPages(0) === "0 print pages");
check("format: 1 -> singular", formatPrintPages(1) === "~1 print page");
check(
  "format: 48 -> plural",
  formatPrintPages(48) === "~48 print pages",
);

if (failures > 0) throw new Error(`${failures} export test(s) failed`);
console.log("ALL EXPORT TESTS PASS");
