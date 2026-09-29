import JSZip from "jszip";
import {
  blocksToDocxBuffer,
  chaptersToDocxBuffer,
  lengthToTwips,
} from "../src/export/docx.ts";
import {
  blocksToEpubBuffer,
  chaptersToEpubBuffer,
} from "../src/export/epub.ts";
import {
  countWords,
  formatPrintPages,
  WORDS_PER_PAGE,
  wordsToPages,
} from "../src/stats.ts";
import type { Block, Chapter, TextRunModel } from "../src/export/text.ts";
import {
  chapterTitleNeeded,
  splitOnBreaks,
  titleHeading,
} from "../src/export/text.ts";
import { compareChapterPaths } from "../src/export/target.ts";
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
const dropcapDocxBuf = await blocksToDocxBuffer(blocks, "2em", "1.7", true);
const dropcapDocxZip = await JSZip.loadAsync(dropcapDocxBuf);
const dropcapDocXml = await dropcapDocxZip.file("word/document.xml")
  ?.async("string") ?? "";
check(
  "docx: two-line first-paragraph drop cap",
  dropcapDocXml.includes('w:dropCap="drop"') &&
    dropcapDocXml.includes('w:lines="2"'),
);

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
const dropcapEpub = await readChapter(
  [para("Once upon a time"), para("Next")],
  { enableDropCap: true },
);
check(
  "epub: first paragraph gets drop cap class",
  dropcapEpub.chapter.includes('<p class="flush dropcap">Once upon a time</p>'),
);
check(
  "epub: drop cap styling spans two lines",
  dropcapEpub.css.includes("p.dropcap::first-letter") &&
    dropcapEpub.css.includes("font-size: 3em !important"),
);
const afterHeading = await readChapter(
  [headingBlock, para("First"), para("Second")],
  { flushAfterHeading: true },
);
check(
  "epub: paragraph after heading flush when toggle on",
  afterHeading.chapter.includes(
    '<h1>Title</h1>\n<p class="flush">First</p>\n<p>Second</p>',
  ),
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
  afterBreak.chapter.includes(
    '<p class="scene">* * *</p>\n<p class="flush">After</p>\n<p>Later</p>',
  ),
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

// --- Novel export: one section per chapter ---
const count = (haystack: string, needle: string) =>
  haystack.split(needle).length - 1;
const novelChapters: Chapter[] = [
  { title: "Chapter One", blocks },
  { title: "Chapter Two", blocks: [headingBlock, para("Second chapter.")] },
];

const novelDocxBuf = await chaptersToDocxBuffer(novelChapters, "2em", "1.7");
const novelDocxZip = await JSZip.loadAsync(novelDocxBuf);
const novelDocxXml = await novelDocxZip.file("word/document.xml")
  ?.async("string") ?? "";
check(
  "docx: chapter 2 starts on a new page",
  novelDocxXml.includes(
    "<w:pageBreakBefore/>",
  ),
);
check(
  "docx: exactly one page break for two chapters",
  count(novelDocxXml, "<w:pageBreakBefore/>") === 1,
);
check(
  "docx: single-chapter export never breaks the page",
  !docXml.includes("<w:pageBreakBefore/>"),
);
check(
  "docx: every chapter's first paragraph is flush",
  count(novelDocxXml, 'w:firstLine="0"') === 2,
);
const novelDropcapBuf = await chaptersToDocxBuffer(
  novelChapters,
  "2em",
  "1.7",
  true,
);
const novelDropcapXml = await (await JSZip.loadAsync(novelDropcapBuf)).file(
  "word/document.xml",
)?.async("string") ?? "";
check(
  "docx: drop cap on each chapter's first paragraph",
  count(novelDropcapXml, 'w:dropCap="drop"') === 2,
);

const novelEpubBuf = await chaptersToEpubBuffer(novelChapters, {
  title: "My Novel",
  indent: "2em",
  lineHeight: "1.7",
});
const novelEpubZip = await JSZip.loadAsync(novelEpubBuf);
check(
  "epub: one xhtml file per chapter",
  !!novelEpubZip.file("OEBPS/chapter.xhtml") &&
    !!novelEpubZip.file("OEBPS/chapter-2.xhtml"),
);
const novelOpf =
  await novelEpubZip.file("OEBPS/content.opf")?.async("string") ??
    "";
check(
  "epub: book title in metadata",
  novelOpf.includes("<dc:title>My Novel</dc:title>"),
);
check(
  "epub: manifest + spine carry every chapter",
  novelOpf.includes(
    '<item id="chapter-2" href="chapter-2.xhtml" media-type="application/xhtml+xml"/>',
  ) && count(novelOpf, "<itemref ") === 2,
);
const novelNav = await novelEpubZip.file("OEBPS/nav.xhtml")?.async("string") ??
  "";
const navOne = novelNav.indexOf('<a href="chapter.xhtml">Chapter One</a>');
const navTwo = novelNav.indexOf('<a href="chapter-2.xhtml">Chapter Two</a>');
check(
  "epub: nav TOC lists chapters in reading order",
  navOne !== -1 && navTwo > navOne,
);
const novelChapter2 = await novelEpubZip.file("OEBPS/chapter-2.xhtml")
  ?.async("string") ?? "";
check(
  "epub: chapter 2 keeps its own first-paragraph state (flush)",
  novelChapter2.includes(
    '<h1>Title</h1>\n<p class="flush">Second chapter.</p>',
  ),
);

// --- Chapter titles: the file name labels every exported chapter ---
check(
  "title: needed when the note opens with a paragraph",
  chapterTitleNeeded("1 - Awakening", [para("Once upon a time.")]),
);
check(
  "title: not needed when the note already opens with that heading",
  !chapterTitleNeeded("1 - Awakening", [titleHeading("1 - Awakening")]),
);
check(
  "title: comparison ignores case and repeated whitespace",
  !chapterTitleNeeded("Ch 1", [titleHeading("  CH   1  ")]),
);
check(
  "title: a different opening heading still gets the chapter name",
  chapterTitleNeeded("1 - Awakening", [titleHeading("Prologue")]),
);
check(
  "title: heading block is level 1 and keeps markdown punctuation literal",
  (() => {
    const heading = titleHeading('Ch*1: [Act]"');
    return heading.kind === "heading" && heading.level === 1 &&
      heading.runs[0].text === 'Ch*1: [Act]"';
  })(),
);

const titledChapters: Chapter[] = [
  {
    title: "1 - Awakening",
    blocks: [titleHeading("1 - Awakening"), para("The story begins.")],
  },
  {
    title: "2 - Return",
    blocks: [titleHeading("2 - Return"), para("And continues.")],
  },
];

const titledDocxBuf = await chaptersToDocxBuffer(titledChapters, "2em", "1.7");
const titledDocxXml = await (await JSZip.loadAsync(titledDocxBuf)).file(
  "word/document.xml",
)?.async("string") ?? "";
check(
  "docx: chapter titles render as level-1 headings",
  count(titledDocxXml, 'w:pStyle w:val="Heading1"') === 2 &&
    titledDocxXml.includes("1 - Awakening"),
);
check(
  "docx: the page break lands on chapter 2's title",
  count(titledDocxXml, "<w:pageBreakBefore/>") === 1,
);
const titledDropcapXml = await (await JSZip.loadAsync(
  await chaptersToDocxBuffer(titledChapters, "2em", "1.7", true),
)).file("word/document.xml")?.async("string") ?? "";
check(
  "docx: the title heading doesn't consume the drop cap",
  count(titledDropcapXml, 'w:dropCap="drop"') === 2,
);

const titledEpubBuf = await chaptersToEpubBuffer(titledChapters, {
  title: "My Novel",
  indent: "2em",
  lineHeight: "1.7",
  enableDropCap: true,
});
const titledEpubZip = await JSZip.loadAsync(titledEpubBuf);
const titledChapterOne = await titledEpubZip.file("OEBPS/chapter.xhtml")
  ?.async("string") ?? "";
check(
  "epub: chapter title heading sits at the top of the chapter body",
  titledChapterOne.includes(
    '<h1>1 - Awakening</h1>\n<p class="flush dropcap">The story begins.</p>',
  ),
);
const titledChapterTwo = await titledEpubZip.file("OEBPS/chapter-2.xhtml")
  ?.async("string") ?? "";
check(
  "epub: each chapter opens with its own title heading",
  titledChapterTwo.includes(
    '<h1>2 - Return</h1>\n<p class="flush dropcap">And continues.</p>',
  ),
);

// --- Novel reading order (1 file = chapter, folder = novel) ---
check(
  "order: natural sort puts Ch2 before Ch10",
  compareChapterPaths("Manuscript/Ch2.md", "Manuscript/Ch10.md") < 0 &&
    compareChapterPaths("Manuscript/Ch10.md", "Manuscript/Ch2.md") > 0,
);
check(
  "order: case-insensitive tie-break",
  compareChapterPaths("Manuscript/ch1.md", "Manuscript/Ch1.md") === 0,
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
  printCss.includes(".author-pp-first > p:first-child") &&
    printCss.includes(".author-pp p.author-pp-first-paragraph"),
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
  "print: optional two-line drop cap rule",
  buildPrintCss("2em", "1.7", true, true, true).includes(
    ".author-pp-first-paragraph::first-letter",
  ) && !buildPrintCss("2em", "1.7", true).includes("::first-letter"),
);
// The isolated PDF render can put the manuscript marker on the <p> itself
// instead of a wrapper section. The drop-cap rule already matches that
// shape, so the flush rule must too — otherwise the cap renders while its
// first line keeps the 2em indent from `p.author-pp`.
const firstFlush = [
  ".author-pp-first > p:first-child",
  "p.author-pp-first",
  ".author-pp p.author-pp-first-paragraph",
  "p.author-pp-first-paragraph",
].join(", ");
check(
  "print: first paragraph flushes when the marker sits on the <p> itself",
  printCss.includes(firstFlush) &&
    buildPrintCss("2em", "1.7", false, true, true).includes(firstFlush),
);
check(
  "print: flush rule outranks the indent rule by source order (same specificity)",
  printCss.indexOf(firstFlush) >
      printCss.indexOf("p.author-pp { text-indent: 2em !important; }") &&
    printCss.indexOf(firstFlush) !== -1,
);
check(
  "print: flush after heading covers a top-level <p> section",
  printCss.includes("h1 + p.author-pp-flush") &&
    printCss.includes("hr + p.author-pp-flush"),
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
  toAbsolutePath("/vault", "Manuscript/ch1.pdf") ===
    "/vault/Manuscript/ch1.pdf",
);
check(
  "reveal: trailing/leading slashes",
  toAbsolutePath("/vault/", "/Manuscript/ch1.pdf") ===
    "/vault/Manuscript/ch1.pdf",
);
check(
  "reveal: empty vault path -> base",
  toAbsolutePath("/vault", "") === "/vault",
);

// --- Live Preview: the first line of the file must stay flush ---
// The CM6 marker (src/dropcap.ts) is scoped to the manuscript folder only —
// never to the drop-cap toggle — so this holds with the cap on or off.
const screenCss = await Deno.readTextFile(
  new URL("../src/styles.css", import.meta.url),
);
const dropcapSource = await Deno.readTextFile(
  new URL("../src/dropcap.ts", import.meta.url),
);
check(
  "dropcap: live preview first-line marker ignores the drop-cap toggle",
  !dropcapSource.includes("enableDropCap"),
);
check(
  "styles: live preview indent skips the first-line marker",
  /:not\(\.author-dropcap-line\)/.test(screenCss),
);
check(
  "styles: live preview first-line marker is forced to indent 0",
  /\.author-dropcap-line\s*\{\s*text-indent:\s*0;\s*\}/.test(screenCss),
);
check(
  "styles: live preview indent skips quote/callout lines (Reading view flushes them)",
  /:not\(\.HyperMD-quote\)/.test(screenCss),
);

if (failures > 0) throw new Error(`${failures} export test(s) failed`);
console.log("ALL EXPORT TESTS PASS");
