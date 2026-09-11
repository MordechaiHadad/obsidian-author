/// <reference lib="dom" />
import { Component, MarkdownRenderer } from "obsidian";
import type { App, TFile } from "obsidian";
import { buildPrintCss, markManuscriptRoot } from "../scope.ts";

export interface ManuscriptPdfOptions {
  indent: string;
  lineHeight: string;
  flushAfterHeading: boolean;
}

/** Minimal Electron surface we need. Acquired at runtime via Obsidian's
 * `window.electron.remote` (desktop only); never imported, so mobile and
 * bundling stay safe. Mirrors the proven export-readview-pdf technique. */
interface PrintWindow {
  loadURL(url: string): Promise<void>;
  webContents: {
    executeJavaScript(code: string, userGesture?: boolean): Promise<unknown>;
    printToPDF(options: Record<string, unknown>): Promise<unknown>;
  };
  destroy(): void;
  isDestroyed(): boolean;
}

function getBrowserWindowCtor(): (new (
  opts: Record<string, unknown>,
) => PrintWindow) | null {
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
      for (const rule of Array.from(sheet.cssRules)) text += rule.cssText + "\n";
      if (text) parts.push(text);
    } catch {
      // Cross-origin / inaccessible sheet: skip.
    }
  };
  for (const sheet of Array.from(document.styleSheets)) pushSheet(sheet);
  const adopted = (document as Document & {
    adoptedStyleSheets?: CSSStyleSheet[];
  }).adoptedStyleSheets;
  if (Array.isArray(adopted)) for (const sheet of adopted) pushSheet(sheet);
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
        reader.onerror = () =>
          reject(new Error("read failed"));
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

/** Render a note fully (non-lazy, no dependency on scroll state or view
 * mode), mark manuscript typography deterministically, and print to PDF.
 * Throws with user-facing messages; desktop Electron only. */
export async function noteToPdfBuffer(
  app: App,
  file: TFile,
  content: string,
  opts: ManuscriptPdfOptions,
): Promise<ArrayBuffer> {
  const BrowserWindow = getBrowserWindowCtor();
  if (!BrowserWindow) {
    throw new Error(
      "PDF export needs desktop Obsidian (Electron). On mobile, use DOCX or EPUB export.",
    );
  }
  const stripped = content.replace(/^---\n[\s\S]*?\n---\n?/, "");
  const host = document.createElement("div");
  host.style.cssText = "position:absolute;left:-100000px;top:0;pointer-events:none;";
  document.body.appendChild(host);
  const component = new Component();
  component.load();
  let printWindow: PrintWindow | null = null;
  try {
    await MarkdownRenderer.render(app, stripped, host, file.path, component);
    // Deterministic marking (does not rely on the post-processor hook):
    // every source line becomes an indentable paragraph.
    markManuscriptRoot(host, opts);
    for (const script of Array.from(host.querySelectorAll("script"))) {
      script.remove();
    }
    await inlineLocalImages(host);

    const proseLength = (host.textContent ?? "").trim().length;
    if (proseLength === 0) throw new Error("nothing to export in this note.");
    const css = collectDocumentCss() +
      "\n" +
      buildPrintCss(opts.indent, opts.lineHeight, opts.flushAfterHeading) +
      "\n" +
      VISIBILITY_OVERRIDES;
    const title = escapeHtml(file.basename);
    const html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>${title}</title>\n<style>\n${css}\n</style>\n</head>\n<body class="export-manuscript-document">\n<div class="markdown-preview-view markdown-rendered author-manuscript author-indent">${host.innerHTML}</div>\n</body>\n</html>`;
    // Data URL (not blob:): no cross-window ownership issues, debuggable.
    const pageUrl = "data:text/html;charset=utf-8," + encodeURIComponent(html);
    try {
      printWindow = new BrowserWindow({
        title: `${file.basename} - manuscript PDF`,
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
      await printWindow.loadURL(pageUrl);
      // Let images + fonts settle, then verify the window actually holds
      // our content — a blank render must fail loudly, never print blank.
      const settled = await printWindow.webContents.executeJavaScript(
        `(() => { const imgs = Array.from(document.images); const wait = (im) => im.complete ? Promise.resolve() : new Promise((res) => { im.addEventListener("load", res, { once: true }); im.addEventListener("error", res, { once: true }); setTimeout(res, 5000); }); const fonts = document.fonts ? document.fonts.ready.catch(() => undefined) : Promise.resolve(); return Promise.all([Promise.all(imgs.map(wait)), fonts]).then(() => new Promise((res) => setTimeout(res, 100))).then(() => ({ textLength: (document.body.textContent || "").trim().length })); })()`,
        true,
      );
      const textLength = Number(
        (settled as { textLength?: unknown } | null)?.textLength ?? 0,
      );
      if (!Number.isFinite(textLength) || textLength === 0) {
        throw new Error(
          "the print render came out empty (styles may be hiding content).",
        );
      }
      const raw = await printWindow.webContents.printToPDF({
        printBackground: true,
        landscape: false,
      });
      return toVerifiedArrayBuffer(raw);
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
