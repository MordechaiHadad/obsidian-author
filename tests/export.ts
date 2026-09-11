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
import {
  buildEpubCss,
  buildPrintCss,
  isManuscriptPath,
  normalizeFolder,
} from "../src/scope.ts";
import { toAbsolutePath } from "../src/export/save-dialog.ts";

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
check("epub: css has 2em indent", css.includes("text-indent: 2em !important;"));
check(
  "epub: indent survives reader defaults (!important like the margin reset)",
  css.includes("p { text-indent: 2em !important;") &&
    css.includes("p.flush { text-indent: 0 !important; }"),
);
check(
  "epub: css resets reader paragraph gap (!important, no double spacing)",
  css.includes("margin: 0 !important") &&
    css.includes("margin-block-start: 0 !important") &&
    css.includes("margin-block-end: 0 !important") &&
    css.includes("padding: 0"),
);
check(
  "epub: no fragile first-of-type flush rule",
  !css.includes("first-of-type"),
);
check(
  "epub: first body paragraph flush (manuscript convention)",
  chapter.includes('<p class="flush">It was a '),
);
check(
  "epub: second paragraph keeps indent",
  chapter.includes("<p>Second paragraph here.</p>"),
);
check("epub: nav exists", !!epubZip.file("OEBPS/nav.xhtml"));

// --- EPUB manuscript conventions: flush after heading/break, toggle ---
const para = (text: string): Block => ({
  kind: "paragraph",
  runs: [{
    text,
    bold: false,
    italic: false,
    strike: false,
    code: false,
  }],
});
const headingBlock: Block = {
  kind: "heading",
  level: 1,
  runs: [{
    text: "Title",
    bold: false,
    italic: false,
    strike: false,
    code: false,
  }],
};
const readChapter = async (bs: Block[], opts: Record<string, unknown>) => {
  const buf = await blocksToEpubBuffer(bs, {
    title: "T",
    indent: "2em",
    lineHeight: "1.7",
    ...opts,
  });
  const zip = await JSZip.loadAsync(buf);
  return {
    chapter: await zip.file("OEBPS/chapter.xhtml")?.async("string") ?? "",
    css: await zip.file("OEBPS/style.css")?.async("string") ?? "",
  };
};
const afterHeading = await readChapter(
  [headingBlock, para("First"), para("Second")],
  { flushAfterHeading: true },
);
check(
  "epub: paragraph after heading flush when toggle on",
  afterHeading.chapter.includes("<h1>Title</h1>\n<p class=\"flush\">First</p>\n<p>Second</p>"),
);
const afterHeadingOff = await readChapter(
  [para("Intro"), headingBlock, para("After"), para("Later")],
  { flushAfterHeading: false },
);
check(
  "epub: no flush after heading when toggle off (only first para flush)",
  afterHeadingOff.chapter.includes(
    '<p class="flush">Intro</p>\n<h1>Title</h1>\n<p>After</p>\n<p>Later</p>',
  ),
);
const afterBreak = await readChapter(
  [para("Before"), { kind: "break" }, para("After"), para("Later")],
  { flushAfterHeading: true },
);
check(
  "epub: paragraph after break flush when toggle on",
  afterBreak.chapter.includes('<p class="scene">* * *</p>\n<p class="flush">After</p>\n<p>Later</p>'),
);
const noIndent = await readChapter([para("One"), para("Two")], {
  enableIndent: false,
});
check(
  "epub: toggle off removes all indents",
  noIndent.css.includes("p { text-indent: 0 !important;") &&
    noIndent.chapter.includes('<p class="flush">One</p>') &&
    noIndent.chapter.includes('<p class="flush">Two</p>'),
);

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

// --- Manuscript scope + print CSS (PDF export path) ---
check("scope: normalize slashes", normalizeFolder("/Novels/") === "Novels");
check("scope: normalize empty", normalizeFolder("  ") === "");
check(
  "scope: note in folder",
  isManuscriptPath("Manuscript/ch1.md", "Manuscript") === true,
);
check(
  "scope: note in subfolder",
  isManuscriptPath("Manuscript/act1/ch1.md", "Manuscript") === true,
);
check(
  "scope: outside folder",
  isManuscriptPath("Notes/ch1.md", "Manuscript") === false,
);
check(
  "scope: non-md ignored",
  isManuscriptPath("Manuscript/ch1.pdf", "Manuscript") === false,
);
check(
  "scope: empty folder disabled",
  isManuscriptPath("Manuscript/ch1.md", "") === false,
);
const printCss = buildPrintCss("2em", "1.7", true);
check("print: @media print", printCss.includes("@media print"));
check(
  "print: literal indent !important",
  printCss.includes("text-indent: 2em !important"),
);
check(
  "print: literal line-height !important",
  printCss.includes("line-height: 1.7 !important"),
);
check("print: marker class", printCss.includes(".author-pp p"));
check(
  "print: flush after heading",
  printCss.includes("h1 + .author-pp-flush > p:first-child"),
);
check(
  "print: first paragraph flush",
  printCss.includes(".author-pp-first > p:first-child"),
);
check(
  "print: no flush rules when disabled",
  !buildPrintCss("2em", "1.7", false).includes(".author-pp-flush"),
);
check(
  "print: toggle off removes indent (overrides styles.css fallback)",
  buildPrintCss("2em", "1.7", true, false).includes(
    ".author-pp p { text-indent: 0 !important;",
  ) && !buildPrintCss("2em", "1.7", true, false).includes("text-indent: 2em"),
);
check(
  "css: single source of truth (epub zip css === buildEpubCss)",
  css === buildEpubCss("2em", "1.7", true),
);
check(
  "css: buildEpubCss toggle off removes indent",
  buildEpubCss("2em", "1.7", false).includes("p { text-indent: 0 !important;"),
);

// --- Save-dialog path join ---
check(
  "reveal: join base + vault path",
  toAbsolutePath("/vault", "Manuscript/ch1.pdf") === "/vault/Manuscript/ch1.pdf",
);
check(
  "reveal: trailing/leading slashes",
  toAbsolutePath("/vault/", "/Manuscript/ch1.pdf") === "/vault/Manuscript/ch1.pdf",
);
check(
  "reveal: empty vault path -> base",
  toAbsolutePath("/vault", "") === "/vault",
);

if (failures > 0) throw new Error(`${failures} export test(s) failed`);
console.log("ALL EXPORT TESTS PASS");
