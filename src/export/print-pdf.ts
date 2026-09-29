/// <reference lib="dom" />
import { Component, MarkdownRenderer } from "obsidian";
import type { App, TFile } from "obsidian";
import { buildPrintCss, markManuscriptRoot } from "../scope.ts";

export interface ManuscriptPdfOptions {
  indent: string;
  lineHeight: string;
  flushAfterHeading: boolean;
  enableIndent: boolean;
  enableDropCap: boolean;
}

/** Minimal Electron surface we need. Acquired at runtime via Obsidian's
 * `window.electron.remote` (desktop only); never imported, so mobile and
 * bundling stay safe. Mirrors the proven export-readview-pdf technique. */
interface PrintWindow {
  loadURL(url: string): Promise<void>;
  webContents: {
    executeJavaScript(code: string, userGesture?: boolean): Promise<unknown>;
    printToPDF(options: Record<string, unknown>): Promise<unknown>;
    /** Optional: emulate print media so the probe lays the snapshot out the
     * way printToPDF will (and catches rules that hide content only when
     * printing). Older Electron builds don't expose it; probe then runs on
     * screen media, which still catches structural hiding. */
    setEmulatedMedia?(options: { media?: string }): void;
  };
  destroy(): void;
  isDestroyed(): boolean;
}

function getBrowserWindowCtor():
  | (new (
    opts: Record<string, unknown>,
  ) => PrintWindow)
  | null {
  const w = window as unknown as Record<string, unknown>;
  const electron = w["electron"] as
    | { remote?: { BrowserWindow?: unknown } }
    | undefined;
  const fromObsidian = electron?.remote?.BrowserWindow;
  if (typeof fromObsidian === "function") {
    return fromObsidian as new (
      opts: Record<string, unknown>,
    ) => PrintWindow;
  }
  const req = w["require"] as unknown;
  if (typeof req === "function") {
    try {
      const ctor = (req as (id: string) => { BrowserWindow?: unknown })(
        "@electron/remote",
      )?.BrowserWindow;
      if (typeof ctor === "function") {
        return ctor as new (
          opts: Record<string, unknown>,
        ) => PrintWindow;
      }
    } catch {
      // Obsidian-provided remote is the primary path; ignore fallback errors.
    }
  }
  return null;
}

/** Collect all document stylesheets as text so the print window renders
 * with the same theme + plugin CSS. Cross-origin sheets are skipped. */
function collectDocumentCss(): string {
  const parts: string[] = [];
  const pushSheet = (sheet: CSSStyleSheet): void => {
    try {
      let text = "";
      for (const rule of Array.from(sheet.cssRules)) {
        text += rule.cssText + "\n";
      }
      if (text) parts.push(text);
    } catch {
      // Cross-origin / inaccessible sheet: skip.
    }
  };
  for (const sheet of Array.from(document.styleSheets)) pushSheet(sheet);
  const adopted = (document as Document & {
    adoptedStyleSheets?: CSSStyleSheet[];
  }).adoptedStyleSheets;
  if (Array.isArray(adopted)) { for (const sheet of adopted) pushSheet(sheet); }
  return parts.join("\n");
}

/** Best-effort: inline images as data URLs so the sandboxed print window
 * (blob URL origin) can load vault-local assets. Failures keep the
 * original src. Manuscripts are usually text-only; this is a bonus. */
async function inlineLocalImages(root: HTMLElement): Promise<void> {
  for (const img of Array.from(root.querySelectorAll("img"))) {
    const src = img.getAttribute("src");
    if (!src || src.startsWith("data:")) continue;
    try {
      const res = await fetch(src);
      if (!res.ok) continue;
      const blob = await res.blob();
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = () => reject(new Error("read failed"));
        reader.readAsDataURL(blob);
      });
      img.setAttribute("src", dataUrl);
      img.removeAttribute("srcset");
    } catch {
      // Keep original src.
    }
  }
}

/** Last-resort overrides: the copied app/theme CSS was written for the
 * workspace layout, not a standalone page — make sure nothing in it can
 * hide our snapshot (a blank PDF must be impossible, not silent). */
const VISIBILITY_OVERRIDES = [
  ".export-manuscript-document { background: #ffffff !important; }",
  ".export-manuscript-document .markdown-preview-view { display: block !important; visibility: visible !important; opacity: 1 !important; color: #000000 !important; background: #ffffff !important; max-width: none !important; }",
  ".export-manuscript-document .markdown-preview-view * { visibility: visible !important; opacity: 1 !important; }",
].join("\n");

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toVerifiedArrayBuffer(value: unknown): ArrayBuffer {
  let buffer: ArrayBuffer;
  if (value instanceof ArrayBuffer) buffer = value.slice(0);
  else if (ArrayBuffer.isView(value)) {
    buffer = Uint8Array.from(value as Uint8Array).buffer as ArrayBuffer;
  } else throw new Error("Electron returned invalid PDF data");
  const bytes = new Uint8Array(buffer);
  if (
    bytes.byteLength < 5 ||
    String.fromCharCode(...bytes.slice(0, 5)) !== "%PDF-"
  ) {
    throw new Error("Electron returned invalid PDF data");
  }
  return buffer;
}

/** One note to place in the print render: the file supplies the source path
 * for links/embeds, `content` is its raw markdown. `heading` is the chapter
 * title to render as an H1 above the note when the note doesn't already open
 * with one — set by the caller so DOCX/EPUB/PDF label chapters alike. */
export interface PdfChapter {
  file: TFile;
  content: string;
  heading?: string;
}

/** Chapter boundaries start a new sheet (chapter export = a single wrapper,
 * so the sibling selector never fires there). Outside `@media print` so it
 * survives both screen layout and print-media emulation. */
const CHAPTER_BREAK_CSS =
  ".export-chapter + .export-chapter { break-before: page; page-break-before: always; }";

/** Stand-in typography for the fallback pass: the copied app/theme CSS is
 * dropped there, so the snapshot needs its own readable defaults (serif body,
 * black on white, heading spacing) — combined with buildPrintCss below, which
 * keeps the manuscript indent/flush/drop-cap rules either way. */
const MANUSCRIPT_FALLBACK_CSS = [
  "html, body { margin: 0; padding: 0; }",
  "body { font-family: Georgia, 'Times New Roman', serif; font-size: 12pt; line-height: 1.7; }",
  ".markdown-preview-view { margin: 0; padding: 0; }",
  ".markdown-preview-view p { margin: 0; }",
  ".markdown-preview-view h1, .markdown-preview-view h2, .markdown-preview-view h3, .markdown-preview-view h4, .markdown-preview-view h5, .markdown-preview-view h6 { line-height: 1.3; margin: 1.2em 0 0.5em; font-weight: bold; }",
  ".markdown-preview-view h1:first-child { margin-top: 0; }",
  ".markdown-preview-view ul, .markdown-preview-view ol { margin: 0.5em 0; padding-left: 1.6em; }",
  ".markdown-preview-view li { margin: 0; }",
  ".markdown-preview-view blockquote { margin: 0.8em 0 0.8em 1.6em; }",
  ".markdown-preview-view hr { border: none; border-top: 1px solid #000000 !important; margin: 1.4em auto; width: 30%; }",
  ".markdown-preview-view pre, .markdown-preview-view code { font-family: 'Courier New', monospace; white-space: pre-wrap; }",
  ".markdown-preview-view img { max-width: 100%; }",
].join("\n");

/** Post-print sanity floor. A real render embeds font subsets, so prose this
 * long can't come out as a file this small — under it, Chromium painted
 * nothing (a blank PDF once measured well under 1 kB). */
const MIN_PLAUSIBLE_PDF_BYTES = 2048;
const MIN_PROSE_FOR_SIZE_CHECK = 200;

/** Runs inside the print window once images + fonts settle. Reports whether
 * every chapter's opening text is actually laid out *and painted*: a box on
 * screen (not `display:none`/`visibility:hidden`/transparent, not clipped off
 * to the left of the viewport). `textContent` alone is not evidence — hidden
 * text counts there, which is exactly how a blank PDF slips through. */
const PROBE_SCRIPT = `(() => {
  const imgs = Array.from(document.images);
  const wait = (im) => im.complete ? Promise.resolve() : new Promise((res) => {
    im.addEventListener("load", res, { once: true });
    im.addEventListener("error", res, { once: true });
    setTimeout(res, 5000);
  });
  const fonts = document.fonts ? document.fonts.ready.catch(() => undefined) : Promise.resolve();
  return Promise.all([Promise.all(imgs.map(wait)), fonts])
    .then(() => new Promise((res) => setTimeout(res, 100)))
    .then(() => {
      const doc = document.documentElement;
      const views = Array.from(document.querySelectorAll(".markdown-preview-view"));
      let hidden = 0;
      let sample = null;
      for (const view of views) {
        const el = view.querySelector("p, h1, h2, h3, li, blockquote, pre, hr, img");
        if (!el) continue;
        const rects = el.getClientRects();
        const r = rects.length ? rects[0] : null;
        const cs = getComputedStyle(el);
        const ok = !!r && r.width > 0 && r.height > 0 && r.right > 0 && r.bottom > 0 &&
          r.left < doc.clientWidth &&
          cs.visibility !== "hidden" && cs.display !== "none" && Number(cs.opacity) > 0 &&
          (typeof el.checkVisibility !== "function" || el.checkVisibility({ checkOpacity: true, checkVisibilityCSS: true, contentVisibilityAuto: true }));
        if (!ok) {
          hidden += 1;
          if (!sample) {
            sample = el.tagName.toLowerCase() +
              (r ? " box=" + Math.round(r.left) + "," + Math.round(r.top) + " " + Math.round(r.width) + "x" + Math.round(r.height) : " no-box") +
              " display=" + cs.display + " visibility=" + cs.visibility + " opacity=" + cs.opacity;
          }
        }
      }
      return {
        textLength: (document.body.textContent || "").trim().length,
        views: views.length,
        hidden: hidden,
        sample: sample,
      };
    });
})()`;

interface ProbeResult {
  textLength: number;
  views: number;
  hidden: number;
  sample: string | null;
  visible: boolean;
}

function normalizeProbe(value: unknown): ProbeResult {
  const raw = (value ?? {}) as Partial<ProbeResult>;
  const textLength = Number(raw.textLength ?? 0);
  const views = Number(raw.views ?? 0);
  const hidden = Number(raw.hidden ?? 0);
  return {
    textLength: Number.isFinite(textLength) ? textLength : 0,
    views: Number.isFinite(views) ? views : 0,
    hidden: Number.isFinite(hidden) ? hidden : 0,
    sample: typeof raw.sample === "string" ? raw.sample : null,
    visible: textLength > 0 && views > 0 && hidden === 0,
  };
}

/** Probe the loaded snapshot under print-media emulation, then restore the
 * default media so printToPDF lays the page out exactly as it does today. */
async function probeRender(printWindow: PrintWindow): Promise<ProbeResult> {
  const emulate = (media: string): void => {
    try {
      const fn = printWindow.webContents.setEmulatedMedia;
      if (typeof fn === "function") fn.call(printWindow.webContents, { media });
    } catch {
      // Optional API: fall back to probing on screen media.
    }
  };
  emulate("print");
  try {
    return normalizeProbe(
      await printWindow.webContents.executeJavaScript(PROBE_SCRIPT, true),
    );
  } finally {
    emulate("");
  }
}

/** Render one or more notes fully (non-lazy, no dependency on scroll state
 * or view mode), mark manuscript typography deterministically per chapter,
 * and print them to a single PDF — chapter 2+ start on a new page. Throws
 * with user-facing messages; desktop Electron only. */
export async function chaptersToPdfBuffer(
  app: App,
  chapters: PdfChapter[],
  opts: ManuscriptPdfOptions,
  title: string,
): Promise<ArrayBuffer> {
  const BrowserWindow = getBrowserWindowCtor();
  if (!BrowserWindow) {
    throw new Error(
      "PDF export needs desktop Obsidian (Electron). On mobile, use DOCX or EPUB export.",
    );
  }
  const host = document.createElement("div");
  host.style.cssText =
    "position:absolute;left:-100000px;top:0;pointer-events:none;";
  document.body.appendChild(host);
  const component = new Component();
  component.load();
  let printWindow: PrintWindow | null = null;
  try {
    const rendered: string[] = [];
    let proseLength = 0;
    for (const chapter of chapters) {
      const stripped = chapter.content.replace(/^---\n[\s\S]*?\n---\n?/, "");
      await MarkdownRenderer.render(
        app,
        stripped,
        host,
        chapter.file.path,
        component,
      );
      // Chapter title as an H1 above the note (added DOM-side so file names
      // with markdown punctuation stay literal). Prepended before marking:
      // headings don't consume the first-paragraph flush or the drop cap.
      if (chapter.heading) {
        const heading = document.createElement("h1");
        heading.textContent = chapter.heading;
        host.prepend(heading);
      }
      // Deterministic marking (does not rely on the post-processor hook):
      // every source line becomes an indentable paragraph. Per chapter, so
      // each note's first paragraph is flush and carries the drop cap.
      markManuscriptRoot(host, opts);
      for (const script of Array.from(host.querySelectorAll("script"))) {
        script.remove();
      }
      await inlineLocalImages(host);

      proseLength += (host.textContent ?? "").trim().length;
      rendered.push(host.innerHTML);
      host.replaceChildren();
    }
    if (proseLength === 0) throw new Error("nothing to export.");
    const generatedCss = buildPrintCss(
      opts.indent,
      opts.lineHeight,
      opts.flushAfterHeading,
      opts.enableIndent,
      opts.enableDropCap,
    ) +
      "\n" +
      CHAPTER_BREAK_CSS +
      "\n" +
      VISIBILITY_OVERRIDES;
    // The theme-free stylesheet: fallback attempt, and the reprint guard.
    const fallbackCss = MANUSCRIPT_FALLBACK_CSS + "\n" + generatedCss;
    const escapedTitle = escapeHtml(title);
    // Every chapter is its own preview root directly under <body>: the shape
    // the visibility overrides (and a plain single-note PDF) were built and
    // proven with. The export-chapter class exists only to break the page
    // between siblings — no extra wrapper the page's CSS can key off.
    const body = rendered
      .map(
        (html) =>
          `<div class="markdown-preview-view markdown-rendered author-manuscript author-indent export-chapter">${html}</div>`,
      )
      .join("\n");
    const buildHtml = (css: string) =>
      `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>${escapedTitle}</title>\n<style>\n${css}\n</style>\n</head>\n<body class="export-manuscript-document">\n${body}\n</body>\n</html>`;
    // Data URL (not blob:): no cross-window ownership issues, debuggable.
    const toDataUrl = (html: string) =>
      "data:text/html;charset=utf-8," + encodeURIComponent(html);
    try {
      printWindow = new BrowserWindow({
        title: `${title} - manuscript PDF`,
        width: 900,
        height: 1200,
        show: false,
        frame: false,
        focusable: false,
        skipTaskbar: true,
        backgroundColor: "#ffffff",
        webPreferences: {
          backgroundThrottling: false,
          contextIsolation: true,
          devTools: false,
          nodeIntegration: false,
          sandbox: true,
          spellcheck: false,
        },
      });
      // The copied app/theme CSS was written for Obsidian's workspace, not a
      // standalone page, so it can hide the snapshot.
      const attempts = [
        {
          label: "app + theme CSS",
          css: collectDocumentCss() + "\n" + generatedCss,
        },
        { label: "manuscript CSS only", css: fallbackCss },
      ];
      // Probe each stylesheet (in print media, the way printToPDF lays the
      // page out) and print the first one whose text is actually painted —
      // text in the DOM proves nothing, hidden text counts there too. The
      // probe only ever *chooses* the CSS: a wrong probe degrades to the
      // plain stylesheet below instead of losing the export.
      const results: { label: string; css: string; probe: ProbeResult }[] = [];
      for (const attempt of attempts) {
        await printWindow.loadURL(toDataUrl(buildHtml(attempt.css)));
        const probe = await probeRender(printWindow);
        results.push({ ...attempt, probe });
        if (probe.visible) break;
        console.error(
          `[Obsidian Author] PDF render hidden with ${attempt.label}:`,
          probe,
        );
      }
      const chosen = results.find((result) => result.probe.visible) ??
        results[results.length - 1];
      if (!results.some((result) => result.probe.visible)) {
        console.error(
          "[Obsidian Author] no stylesheet rendered visibly; printing the manuscript stylesheet anyway.",
        );
      }
      const active = printWindow;
      const print = () =>
        active.webContents.printToPDF({
          printBackground: true,
          landscape: false,
        });
      let buffer = toVerifiedArrayBuffer(await print());
      // Objective last line of defence: a real render embeds font subsets,
      // so prose this long can't come out as a file this small — Chromium
      // painted nothing (which the probe can't see when it fails only at
      // print time). Reprint with the other stylesheet, keep whichever
      // output holds the text, and only a still-blank result blocks it.
      if (
        chosen.probe.textLength >= MIN_PROSE_FOR_SIZE_CHECK &&
        buffer.byteLength < MIN_PLAUSIBLE_PDF_BYTES
      ) {
        if (chosen.css !== fallbackCss) {
          console.error(
            `[Obsidian Author] PDF printed ${buffer.byteLength} bytes for ${chosen.probe.textLength} characters of prose; reprinting with manuscript CSS.`,
            chosen.probe,
          );
          await printWindow.loadURL(toDataUrl(buildHtml(fallbackCss)));
          const retry = toVerifiedArrayBuffer(await print());
          if (retry.byteLength > buffer.byteLength) buffer = retry;
        }
        if (buffer.byteLength < MIN_PLAUSIBLE_PDF_BYTES) {
          throw new Error(
            "the print render came out empty (nothing was painted onto the page).",
          );
        }
      }
      return buffer;
    } finally {
      try {
        if (printWindow && !printWindow.isDestroyed()) printWindow.destroy();
      } catch {
        // Ignore teardown errors.
      }
      printWindow = null;
    }
  } finally {
    component.unload();
    host.remove();
  }
}
