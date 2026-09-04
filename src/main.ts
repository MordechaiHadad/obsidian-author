/// <reference lib="dom" />
import {
  AbstractInputSuggest,
  App,
  MarkdownView,
  Plugin,
  PluginSettingTab,
  Setting,
  TFile,
  TFolder,
  WorkspaceLeaf,
} from "obsidian";

interface AuthorSettings {
  manuscriptFolder: string;
  enableIndent: boolean;
  indentSize: string;
  removeIndentAfterHeading: boolean;
  lineHeight: string;
}

const DEFAULT_SETTINGS: AuthorSettings = {
  manuscriptFolder: "",
  enableIndent: true,
  indentSize: "2em",
  removeIndentAfterHeading: true,
  lineHeight: "1.7",
};

// Scope class + CSS variables consumed by styles.css. Applied per markdown
// leaf container (not body): only manuscript notes match, even with
// side-by-side panes or pop-out windows. Mirrors the proven pattern of
// per-note scoping (e.g. via cssclasses) used by community typography snippets.
const SCOPE_CLASS = "author-manuscript";
const CLASS_INDENT = "author-indent";
const CLASS_FLUSH_AFTER_HEADING = "author-flush-after-heading";
const VAR_INDENT = "--author-indent";
const VAR_LINE_HEIGHT = "--author-line-height";

export default class AuthorPlugin extends Plugin {
  declare settings: AuthorSettings;
  private statusEl: HTMLElement | null = null;

  override async onload() {
    await this.loadSettings();
    this.addSettingTab(new AuthorSettingTab(this.app, this));
    // Status indicator: visible proof of manuscript scope for the active note.
    this.statusEl = this.addStatusBarItem();

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

    this.app.workspace.onLayoutReady(() => this.refresh());
    this.refresh();
  }

  override onunload() {
    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      this.clearScope(leaf);
    }
    this.statusEl?.setText("");
  }

  /** Remove all scope classes and variables from one leaf. */
  private clearScope(leaf: WorkspaceLeaf): void {
    if (!(leaf.view instanceof MarkdownView)) return;
    leaf.view.containerEl.classList.remove(
      SCOPE_CLASS,
      CLASS_INDENT,
      CLASS_FLUSH_AFTER_HEADING,
    );
    leaf.view.containerEl.style.removeProperty(VAR_INDENT);
    leaf.view.containerEl.style.removeProperty(VAR_LINE_HEIGHT);
  }

  async loadSettings() {
    // loadData() is typed as Promise<any> by the Obsidian API; narrow it
    // at the boundary so no `any` leaks into the plugin.
    const data = (await this.loadData()) as
      | Partial<AuthorSettings>
      | null
      | undefined;
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
  }

  async saveSettings() {
    await this.saveData(this.settings);
    this.refresh();
  }

  /** Normalize "Novels/", "/Novels", " novels " -> "Novels". "" means disabled. */
  private normalizedFolder(): string {
    return (this.settings.manuscriptFolder ?? "")
      .trim()
      .replace(/^\/+|\/+$/g, "")
      .replace(/\/{2,}/g, "/");
  }

  /** True when file lives inside the manuscript folder (subfolders included). */
  isManuscriptFile(file: TFile | null | undefined): boolean {
    if (!file || file.extension !== "md") return false;
    const folder = this.normalizedFolder();
    if (!folder) return false;
    return file.path.startsWith(folder + "/");
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
    this.statusEl?.setText(activeOn ? "✒ Manuscript" : "");
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
