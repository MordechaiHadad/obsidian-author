/// <reference lib="dom" />
import {
  AbstractInputSuggest,
  App,
  MarkdownView,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  TFile,
  TFolder,
  WorkspaceLeaf,
} from "obsidian";
import { chaptersToDocxBuffer } from "./export/docx.ts";
import { chaptersToEpubBuffer } from "./export/epub.ts";
import {
  EXPORT_FORMATS,
  type ExportFormat,
  ExportModal,
  isPdfAvailable,
} from "./export/export-modal.ts";
import { noteToBlocks } from "./export/model.ts";
import { chaptersToPdfBuffer, type PdfChapter } from "./export/print-pdf.ts";
import { chooseSavePath, writeAbsoluteFile } from "./export/save-dialog.ts";
import { novelFiles } from "./export/target.ts";
import {
  type Chapter,
  chapterTitleNeeded,
  titleHeading,
} from "./export/text.ts";
import {
  buildPrintCss,
  isFirstParagraphSection,
  isFrontmatterContainer,
  isManuscriptPath,
  markManuscriptSection,
  normalizeFolder,
  reconcileFirstParagraphMarkers,
} from "./scope.ts";
import { dropcapExtension, setDropcapFolder } from "./dropcap.ts";
import { countWords, formatPrintPages, wordsToPages } from "./stats.ts";

interface AuthorSettings {
  manuscriptFolder: string;
  enableIndent: boolean;
  indentSize: string;
  removeIndentAfterHeading: boolean;
  lineHeight: string;
  enableDropCap: boolean;
  /** Format preselected the next time the Export modal opens. */
  lastExportFormat: ExportFormat;
}

const DEFAULT_SETTINGS: AuthorSettings = {
  manuscriptFolder: "",
  enableIndent: true,
  indentSize: "2em",
  removeIndentAfterHeading: true,
  lineHeight: "1.7",
  enableDropCap: true,
  lastExportFormat: "docx",
};

// Scope class + CSS variables consumed by styles.css. Applied per markdown
// leaf container (not body): only manuscript notes match, even with
// side-by-side panes or pop-out windows. Mirrors the proven pattern of
// per-note scoping (e.g. via cssclasses) used by community typography snippets.
const SCOPE_CLASS = "author-manuscript";
const CLASS_INDENT = "author-indent";
const CLASS_FLUSH_AFTER_HEADING = "author-flush-after-heading";
const CLASS_DROPCAP = "author-dropcap";
const VAR_INDENT = "--author-indent";
const VAR_LINE_HEIGHT = "--author-line-height";
// Id of the dynamic `@media print` style element injected into
// document.head (carries the user's literal indent/line-height).
const PRINT_STYLE_ID = "obsidian-author-print";

export default class AuthorPlugin extends Plugin {
  declare settings: AuthorSettings;
  private statusEl: HTMLElement | null = null;
  private previewObserver: MutationObserver | null = null;
  private previewScopeQueued = false;
  private statusTimer: ReturnType<typeof setTimeout> | null = null;
  private editorOptionsTimer: ReturnType<typeof setTimeout> | null = null;

  override async onload() {
    await this.loadSettings();
    setDropcapFolder(this.normalizedFolder());
    this.addSettingTab(new AuthorSettingTab(this.app, this));
    // Live Preview drop cap as a CodeMirror extension (state decoration —
    // survives editor transactions, no DOM fighting, vim-safe).
    this.registerEditorExtension(dropcapExtension);
    // Status indicator: visible proof of manuscript scope for the active note.
    this.statusEl = this.addStatusBarItem();
    this.addCommand({
      id: "export",
      name: "Export…",
      callback: () => {
        const active = this.app.workspace.getActiveFile();
        new ExportModal(this.app, {
          // Prefill the open note (chapter); fall back to the manuscript
          // folder (novel) when there is no Markdown note to point at.
          initialPath: active instanceof TFile && active.extension === "md"
            ? active.path
            : this.normalizedFolder(),
          initialFormat: this.settings.lastExportFormat,
          onFormatChange: (format) => {
            this.settings.lastExportFormat = format;
            // Format only: no full refresh() needed for a modal preference.
            void this.saveData(this.settings);
          },
          onExport: (target, format) => {
            void this.exportTarget(target, format);
          },
        }).open();
      },
    });
    // Tag rendered sections of manuscript notes (Reading view). Note:
    // native Export to PDF re-renders without post-processors, so it never
    // sees these markers — manuscript PDF export clones a render we mark
    // ourselves (see export/print-pdf.ts).
    this.registerMarkdownPostProcessor((el, ctx) => {
      if (!this.settings.enableIndent) return;
      const src = ctx.sourcePath;
      if (!src || !isManuscriptPath(src, this.normalizedFolder())) return;
      // Never mark frontmatter/properties containers (display:none, but
      // keep the marker strictly for content sections).
      if (isFrontmatterContainer(el)) return;
      // Manuscript convention: the very first paragraph of the note starts
      // flush left. Frontmatter/properties containers before it don't
      // consume "first" (see isFirstParagraphSection).
      const isFirst = isFirstParagraphSection(el);
      markManuscriptSection(el, isFirst, {
        indent: this.sanitizeIndent(this.settings.indentSize),
        lineHeight: this.sanitizeLineHeight(this.settings.lineHeight),
        flushAfterHeading: this.settings.removeIndentAfterHeading,
      });
    });

    // Re-evaluate whenever the open note or vault contents may have changed.
    this.registerEvent(
      this.app.workspace.on("active-leaf-change", () => this.refresh()),
    );
    this.registerEvent(
      this.app.workspace.on("file-open", () => this.refresh()),
    );
    this.registerEvent(this.app.vault.on("rename", () => this.refresh()));
    this.registerEvent(this.app.vault.on("create", () => this.refresh()));
    this.registerEvent(this.app.vault.on("delete", () => this.refresh()));
    // Mode switches (source <-> live <-> reading) rebuild the view without
    // an editor-change: re-scope reading views on any layout change
    // (rAF-throttled, no-op when nothing changed). Live Preview needs no
    // hook: the drop-cap editor extension is state-driven.
    this.registerEvent(
      this.app.workspace.on("layout-change", () => {
        this.queuePreviewScope();
      }),
    );
    // Live print-page count: recompute (debounced) as the active note is
    // edited, including edits made from another device / outside the editor.
    this.registerEvent(
      this.app.workspace.on("editor-change", () => this.scheduleStatusUpdate()),
    );
    this.registerEvent(
      this.app.vault.on("modify", (file) => {
        if (
          file instanceof TFile && file === this.app.workspace.getActiveFile()
        ) this.scheduleStatusUpdate();
      }),
    );

    // Reading-view elements render asynchronously after their leaf opens.
    // Watch for them and scope them when they appear. ChildList only: the
    // Live Preview drop cap is a state-driven editor extension, so nothing
    // here may touch .cm-line classes (that fight shook the buffer).
    this.previewObserver = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of Array.from(mutation.addedNodes)) {
          if (
            node instanceof HTMLElement &&
            (node.matches(".markdown-preview-view") ||
              node.querySelector(".markdown-preview-view"))
          ) {
            this.queuePreviewScope();
            return;
          }
        }
      }
    });
    this.previewObserver.observe(document.body, {
      childList: true,
      subtree: true,
    });

    this.app.workspace.onLayoutReady(() => this.refresh());
    this.refresh();
  }

  override onunload() {
    this.previewObserver?.disconnect();
    this.previewObserver = null;
    if (this.statusTimer !== null) {
      globalThis.clearTimeout(this.statusTimer);
      this.statusTimer = null;
    }
    if (this.editorOptionsTimer !== null) {
      globalThis.clearTimeout(this.editorOptionsTimer);
      this.editorOptionsTimer = null;
    }
    document.getElementById(PRINT_STYLE_ID)?.remove();
    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      this.clearScope(leaf);
    }
    this.statusEl?.setText("");
  }

  /** Throttled pass applying scope classes to reading-view elements. */
  private queuePreviewScope(): void {
    if (this.previewScopeQueued) return;
    this.previewScopeQueued = true;
    requestAnimationFrame(() => {
      this.previewScopeQueued = false;
      this.scopePreviewElements();
    });
  }

  /** Reconfigure editor extensions after the manuscript folder changes.
   * Folder input saves on each keystroke, so debounce to one refresh after
   * typing settles rather than rebuilding every open editor repeatedly. */
  private scheduleEditorOptionsRefresh(): void {
    if (this.editorOptionsTimer !== null) {
      globalThis.clearTimeout(this.editorOptionsTimer);
    }
    this.editorOptionsTimer = globalThis.setTimeout(() => {
      this.editorOptionsTimer = null;
      this.app.workspace.updateOptions();
    }, 250);
  }

  /** Toggle scope classes/values on every rendered reading view, matched to
   * its leaf's file. Note: native Export to PDF does NOT reuse these nodes —
   * it re-renders into a separate print container — so the PDF path is the
   * post-processor marker (`author-pp`) plus `@media print` CSS instead. */
  private scopePreviewElements(): void {
    const s = this.settings;
    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      const view = leaf.view;
      if (!(view instanceof MarkdownView)) continue;
      const preview = view.containerEl.querySelector(
        ".markdown-preview-view",
      );
      if (!(preview instanceof HTMLElement)) continue;
      const inScope = view.file instanceof TFile &&
        this.isManuscriptFile(view.file);
      preview.classList.toggle(SCOPE_CLASS, inScope);
      preview.classList.toggle(CLASS_INDENT, inScope && s.enableIndent);
      preview.classList.toggle(
        CLASS_FLUSH_AFTER_HEADING,
        inScope && s.enableIndent && s.removeIndentAfterHeading,
      );
      preview.classList.toggle(CLASS_DROPCAP, inScope && s.enableDropCap);
      if (inScope) {
        preview.style.setProperty(
          VAR_INDENT,
          this.sanitizeIndent(s.indentSize),
        );
        preview.style.setProperty(
          VAR_LINE_HEIGHT,
          this.sanitizeLineHeight(s.lineHeight),
        );
        // Bulk renders can mark several sections as "first" (post-processor
        // fires while nodes are detached); document order decides.
        reconcileFirstParagraphMarkers(preview);
      } else {
        preview.style.removeProperty(VAR_INDENT);
        preview.style.removeProperty(VAR_LINE_HEIGHT);
      }
    }
  }

  /** Remove all scope classes and variables from one leaf. */
  private clearScope(leaf: WorkspaceLeaf): void {
    if (!(leaf.view instanceof MarkdownView)) return;
    const elements = [leaf.view.containerEl];
    const preview = leaf.view.containerEl.querySelector(
      ".markdown-preview-view",
    );
    if (preview instanceof HTMLElement) elements.push(preview);
    for (const el of elements) {
      el.classList.remove(
        SCOPE_CLASS,
        CLASS_INDENT,
        CLASS_FLUSH_AFTER_HEADING,
        CLASS_DROPCAP,
      );
      el.style.removeProperty(VAR_INDENT);
      el.style.removeProperty(VAR_LINE_HEIGHT);
    }
  }

  async loadSettings() {
    // loadData() is typed as Promise<any> by the Obsidian API; narrow it
    // at the boundary so no `any` leaks into the plugin.
    const data = (await this.loadData()) as
      | Partial<AuthorSettings>
      | null
      | undefined;
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
    // Reject a stored format from an older/foreign version instead of
    // letting an unknown value reach the dropdown.
    if (
      this.settings.lastExportFormat !== "docx" &&
      this.settings.lastExportFormat !== "epub" &&
      this.settings.lastExportFormat !== "pdf"
    ) {
      this.settings.lastExportFormat = DEFAULT_SETTINGS.lastExportFormat;
    }
  }

  async saveSettings() {
    await this.saveData(this.settings);
    this.refresh();
  }

  /** Normalize "Novels/", "/Novels", " novels " -> "Novels". "" means disabled. */
  private normalizedFolder(): string {
    return normalizeFolder(this.settings.manuscriptFolder);
  }

  /** True when file lives inside the manuscript folder (subfolders included). */
  isManuscriptFile(file: TFile | null | undefined): boolean {
    if (!file || file.extension !== "md") return false;
    return isManuscriptPath(file.path, this.normalizedFolder());
  }

  /** Does the vault currently contain that folder? Used for the settings warning. */
  manuscriptFolderExists(): boolean {
    const folder = this.normalizedFolder();
    if (!folder) return false;
    const abstract = this.app.vault.getAbstractFileByPath(folder);
    return abstract instanceof TFolder;
  }

  /** Recompute scope classes and variables for every markdown leaf.
   * All actual styling lives in styles.css; here we only pass values. */
  refresh() {
    const s = this.settings;
    if (setDropcapFolder(this.normalizedFolder())) {
      this.scheduleEditorOptionsRefresh();
    }
    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      const view = leaf.view;
      if (!(view instanceof MarkdownView)) continue;
      const inScope = view.file instanceof TFile &&
        this.isManuscriptFile(view.file);
      const el = view.containerEl;
      el.classList.toggle(SCOPE_CLASS, inScope);
      el.classList.toggle(CLASS_INDENT, inScope && s.enableIndent);
      el.classList.toggle(
        CLASS_FLUSH_AFTER_HEADING,
        inScope && s.enableIndent && s.removeIndentAfterHeading,
      );
      el.classList.toggle(CLASS_DROPCAP, inScope && s.enableDropCap);
      if (inScope) {
        el.style.setProperty(VAR_INDENT, this.sanitizeIndent(s.indentSize));
        el.style.setProperty(
          VAR_LINE_HEIGHT,
          this.sanitizeLineHeight(s.lineHeight),
        );
      } else {
        el.style.removeProperty(VAR_INDENT);
        el.style.removeProperty(VAR_LINE_HEIGHT);
      }
    }

    const active = this.app.workspace.getActiveFile();
    const activeOn = active instanceof TFile && this.isManuscriptFile(active);
    if (!activeOn) this.statusEl?.setText("");
    else void this.updateStatusBar();
    this.scopePreviewElements();
    this.ensurePrintStyle();
  }

  /** Inject (or refresh) a document-level `@media print` style carrying the
   * user's literal indent/line-height. The static styles.css cannot know
   * settings values, and the print render never copies leaf inline
   * variables — so without this, PDF export falls back to theme defaults.
   * No extra plugin or user snippet required. */
  private ensurePrintStyle(): void {
    const existing = document.getElementById(PRINT_STYLE_ID);
    // Always inject (even when the toggle is off): a zero-indent rule must
    // override the static styles.css print fallback, otherwise OFF would
    // still indent via its var() defaults.
    const css = buildPrintCss(
      this.sanitizeIndent(this.settings.indentSize),
      this.sanitizeLineHeight(this.settings.lineHeight),
      this.settings.removeIndentAfterHeading,
      this.settings.enableIndent,
      this.settings.enableDropCap,
    );
    if (existing instanceof HTMLStyleElement) {
      if (existing.textContent !== css) existing.textContent = css;
      return;
    }
    existing?.remove();
    const style = document.createElement("style");
    style.id = PRINT_STYLE_ID;
    style.textContent = css;
    document.head.appendChild(style);
  }

  /** Debounced wrapper so fast typing re-reads at most ~3x/second. */
  private scheduleStatusUpdate(): void {
    if (this.statusTimer !== null) globalThis.clearTimeout(this.statusTimer);
    this.statusTimer = globalThis.setTimeout(() => {
      this.statusTimer = null;
      void this.updateStatusBar();
    }, 300);
  }

  /** Status bar: "✒ Manuscript · ~48 print pages". Core Obsidian already
   * shows words/characters, so we only add the print-page estimate. */
  private async updateStatusBar(): Promise<void> {
    const active = this.app.workspace.getActiveFile();
    if (!(active instanceof TFile) || !this.isManuscriptFile(active)) {
      this.statusEl?.setText("");
      return;
    }
    const live = this.getActiveEditorText(active);
    if (live !== null) {
      this.setStatusFromText(live);
      return;
    }
    try {
      const content = await this.app.vault.cachedRead(active);
      // Guard against a race: user switched notes while reading.
      if (this.app.workspace.getActiveFile() !== active) return;
      this.setStatusFromText(content);
    } catch (error) {
      console.error("[Obsidian Author] page count read failed:", error);
      this.statusEl?.setText("✒ Manuscript");
    }
  }

  /** Fast path: current editor buffer, no vault I/O. Null when the active
   * note isn't open in an editable Markdown view (e.g. Reading view). */
  private getActiveEditorText(active: TFile): string | null {
    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      const view = leaf.view;
      if (!(view instanceof MarkdownView)) continue;
      if (view.file !== active) continue;
      try {
        const value = view.editor?.getValue();
        if (typeof value === "string") return value;
      } catch {
        return null;
      }
    }
    return null;
  }

  private setStatusFromText(text: string): void {
    const pages = wordsToPages(countWords(text));
    this.statusEl?.setText(`✒ Manuscript · ${formatPrintPages(pages)}`);
  }

  private sanitizeIndent(input: string): string {
    const v = (input ?? "").trim() || DEFAULT_SETTINGS.indentSize;
    // Allow "2", "2em", "24px", "2rem", "5%". Bare numbers become em.
    if (/^\d+(\.\d+)?$/.test(v)) return `${v}em`;
    if (/^\d+(\.\d+)?(em|rem|px|%)$/.test(v)) return v;
    return DEFAULT_SETTINGS.indentSize;
  }

  private sanitizeLineHeight(input: string): string {
    const v = (input ?? "").trim() || DEFAULT_SETTINGS.lineHeight;
    if (/^\d+(\.\d+)?$/.test(v)) return v;
    return DEFAULT_SETTINGS.lineHeight;
  }

  /** Save an export buffer: Save dialog on desktop (user picks the path),
   * vault-relative fallback on mobile. Returns the saved path, or null
   * when the user cancelled. */
  private async saveExportBuffer(
    suggestedVaultPath: string,
    filterName: string,
    ext: string,
    buffer: ArrayBuffer,
  ): Promise<string | null> {
    const choice = await chooseSavePath(this.app, suggestedVaultPath, [
      { name: filterName, extensions: [ext] },
    ]);
    if (choice.kind === "cancelled") return null;
    if (choice.kind === "chosen") {
      await writeAbsoluteFile(choice.path, buffer);
      return choice.path;
    }
    const existing = this.app.vault.getAbstractFileByPath(suggestedVaultPath);
    if (existing instanceof TFile) {
      await this.app.vault.modifyBinary(existing, buffer);
    } else await this.app.vault.createBinary(suggestedVaultPath, buffer);
    return suggestedVaultPath;
  }

  /** Export a chapter (one Markdown note) or a novel (every Markdown note
   * under a folder, in reading order) in the chosen format. Builds one
   * buffer, then opens the OS save dialog — vault-relative fallback on
   * mobile. Re-exporting overwrites the previous file. Unlike native
   * Export to PDF (which re-renders without post-processors and drops the
   * indent for single-newline manuscripts), the PDF target prints a fully
   * marked render through a hidden Electron window, from any view mode. */
  async exportTarget(
    target: TFile | TFolder,
    format: ExportFormat,
  ): Promise<void> {
    if (target instanceof TFile && target.extension !== "md") {
      new Notice("Obsidian Author: only Markdown notes can be exported.");
      return;
    }
    // The modal already greys PDF out on mobile; this covers a desktop-persisted
    // "pdf" reaching a build without Electron's print window.
    if (format === "pdf" && !isPdfAvailable()) {
      new Notice(
        "Obsidian Author: PDF export needs PC — use the desktop app, or export DOCX/EPUB here.",
      );
      return;
    }
    const isNovel = target instanceof TFolder;
    // Book title: the note's name, or the folder's (vault root included).
    const title = isNovel
      ? target.path === "/" ? "Manuscript" : target.name
      : target.basename;
    try {
      const chapters: Chapter[] = [];
      const pdfChapters: PdfChapter[] = [];
      let outDir = "";
      if (isNovel) {
        const files = novelFiles(this.app, target);
        for (const file of files) {
          const content = await this.app.vault.read(file);
          const blocks = await noteToBlocks(this.app, file, content);
          // Notes with no renderable prose don't become (empty) chapters.
          if (blocks.length === 0) continue;
          // Label the chapter with its file name so chapter headers survive
          // every format — unless the note already opens with that heading.
          const heading = chapterTitleNeeded(file.basename, blocks)
            ? file.basename
            : undefined;
          chapters.push({
            title: file.basename,
            blocks: heading ? [titleHeading(heading), ...blocks] : blocks,
          });
          pdfChapters.push({ file, content, heading });
        }
        outDir = target.path === "/" ? "" : `${target.path}/`;
      } else {
        const content = await this.app.vault.read(target);
        const blocks = await noteToBlocks(this.app, target, content);
        if (blocks.length > 0) {
          const heading = chapterTitleNeeded(target.basename, blocks)
            ? target.basename
            : undefined;
          chapters.push({
            title: target.basename,
            blocks: heading ? [titleHeading(heading), ...blocks] : blocks,
          });
          pdfChapters.push({ file: target, content, heading });
        }
        outDir = target.parent && target.parent.path !== "/"
          ? `${target.parent.path}/`
          : "";
      }
      if (chapters.length === 0) {
        new Notice("Obsidian Author: nothing to export.");
        return;
      }
      if (format === "pdf") {
        new Notice("Obsidian Author: rendering manuscript PDF…");
      }
      const buffer = await this.buildExportBuffer(
        format,
        title,
        chapters,
        pdfChapters,
      );
      const meta = EXPORT_FORMATS[format];
      const saved = await this.saveExportBuffer(
        `${outDir}${title}.${meta.ext}`,
        meta.filter,
        meta.ext,
        buffer,
      );
      if (saved) new Notice(`Obsidian Author: exported ${saved}.`);
    } catch (error) {
      console.error("[Obsidian Author] export failed:", error);
      new Notice(
        `Obsidian Author: export failed (${
          error instanceof Error ? error.message : String(error)
        }).`,
      );
    }
  }

  /** Run the writer for `format`. DOCX/EPUB consume the parsed block model;
   * PDF renders the raw markdown itself (each note keeps its own source
   * path, so links and embeds resolve per chapter). */
  private async buildExportBuffer(
    format: ExportFormat,
    title: string,
    chapters: Chapter[],
    pdfChapters: PdfChapter[],
  ): Promise<ArrayBuffer> {
    const indent = this.sanitizeIndent(this.settings.indentSize);
    const lineHeight = this.sanitizeLineHeight(this.settings.lineHeight);
    if (format === "pdf") {
      return await chaptersToPdfBuffer(this.app, pdfChapters, {
        indent,
        lineHeight,
        flushAfterHeading: this.settings.removeIndentAfterHeading,
        enableIndent: this.settings.enableIndent,
        enableDropCap: this.settings.enableDropCap,
      }, title);
    }
    if (format === "docx") {
      return await chaptersToDocxBuffer(
        chapters,
        this.settings.indentSize,
        this.settings.lineHeight,
        this.settings.enableDropCap,
      );
    }
    return await chaptersToEpubBuffer(chapters, {
      title,
      indent,
      lineHeight,
      enableIndent: this.settings.enableIndent,
      flushAfterHeading: this.settings.removeIndentAfterHeading,
      enableDropCap: this.settings.enableDropCap,
    });
  }
}

class AuthorSettingTab extends PluginSettingTab {
  plugin: AuthorPlugin;

  constructor(app: App, plugin: AuthorPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  override display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Obsidian Author" });
    containerEl.createEl("p", {
      text:
        "All Markdown notes under the folder below automatically get manuscript typography (first-line indent) in Reading view and Live Preview.",
      cls: "setting-item-description",
    });

    new Setting(containerEl)
      .setName("Manuscript folder")
      .setDesc(
        "Vault-relative path, e.g. Manuscript. Subfolders are included. Empty = disabled.",
      )
      .addText((text) => {
        new FolderSuggest(this.app, text.inputEl);
        text
          .setPlaceholder("Manuscript")
          .setValue(this.plugin.settings.manuscriptFolder)
          // No full re-render here: display() on every keystroke would
          // steal focus. The warning below refreshes next time settings open.
          .onChange(async (value) => {
            this.plugin.settings.manuscriptFolder = value.trim();
            await this.plugin.saveSettings();
          });
      });

    if (
      this.plugin.settings.manuscriptFolder.trim() &&
      !this.plugin.manuscriptFolderExists()
    ) {
      new Setting(containerEl).setName("⚠ Folder not found").setDesc(
        `No folder named "${this.plugin.settings.manuscriptFolder.trim()}" exists in this vault. Styling is inactive until you create it or fix the path.`,
      );
    }

    new Setting(containerEl)
      .setName("First-line indent")
      .setDesc("Indent the first line of each paragraph, like a printed novel.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableIndent)
          .onChange(async (value) => {
            this.plugin.settings.enableIndent = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Indent size")
      .setDesc("CSS length, e.g. 2em, 24px, 5%.")
      .addText((text) =>
        text
          .setPlaceholder("2em")
          .setValue(this.plugin.settings.indentSize)
          .onChange(async (value) => {
            this.plugin.settings.indentSize = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("No indent after heading / break")
      .setDesc(
        "First paragraph after a heading or horizontal rule starts flush left (standard novel style).",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.removeIndentAfterHeading)
          .onChange(async (value) => {
            this.plugin.settings.removeIndentAfterHeading = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Drop cap")
      .setDesc(
        "Enlarge the first letter of the first paragraph (two lines tall) in Reading, Live Preview, and DOCX, EPUB, and manuscript PDF exports. Frontmatter is never affected.",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableDropCap)
          .onChange(async (value) => {
            this.plugin.settings.enableDropCap = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Line height")
      .setDesc(
        "Unitless multiplier, e.g. 1.7. Applies to manuscript notes only.",
      )
      .addText((text) =>
        text
          .setPlaceholder("1.7")
          .setValue(this.plugin.settings.lineHeight)
          .onChange(async (value) => {
            this.plugin.settings.lineHeight = value;
            await this.plugin.saveSettings();
          })
      );
  }
}

/** Autocomplete vault folders inside the folder text input. */
class FolderSuggest extends AbstractInputSuggest<string> {
  constructor(
    app: App,
    private textInput: HTMLInputElement,
  ) {
    super(app, textInput);
  }

  protected override getSuggestions(query: string): string[] {
    const q = query.toLowerCase().trim();
    const folders = this.app.vault
      .getAllLoadedFiles()
      .filter((f): f is TFolder => f instanceof TFolder)
      .map((f) => f.path)
      .sort();
    if (!q) return folders.slice(0, 20);
    return folders.filter((p) => p.toLowerCase().includes(q)).slice(0, 20);
  }

  override renderSuggestion(path: string, el: HTMLElement): void {
    el.createEl("div", { text: path || "/" });
  }

  override selectSuggestion(path: string): void {
    this.textInput.value = path;
    this.textInput.dispatchEvent(new Event("input"));
    this.close();
  }
}
