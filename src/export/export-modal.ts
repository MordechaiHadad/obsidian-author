/// <reference lib="dom" />
import {
  AbstractInputSuggest,
  Modal,
  Notice,
  Setting,
  TFile,
  TFolder,
} from "obsidian";
import type { App, TAbstractFile } from "obsidian";
import { novelFiles } from "./target.ts";

/** Every export target the writer stack supports. */
export type ExportFormat = "docx" | "epub" | "pdf";

export const EXPORT_FORMATS: Record<
  ExportFormat,
  { label: string; ext: string; filter: string }
> = {
  docx: { label: "DOCX (Word document)", ext: "docx", filter: "Word Document" },
  epub: { label: "EPUB (e-book)", ext: "epub", filter: "EPUB e-book" },
  pdf: { label: "PDF (manuscript)", ext: "pdf", filter: "PDF" },
};

export interface ExportModalOptions {
  /** Prefill: the active note, else the manuscript folder, else "". */
  initialPath: string;
  initialFormat: ExportFormat;
  /** Fired when the dropdown changes so the choice can be persisted. */
  onFormatChange: (format: ExportFormat) => void;
  /** Validated target (a Markdown note or a folder containing them). */
  onExport: (target: TFile | TFolder, format: ExportFormat) => void;
}

/** One dialog for every export: the path decides the scope (a file is a
 * chapter, a folder is a whole novel) and the dropdown decides the format.
 * The OS save picker still opens afterwards — this modal only collects the
 * choice, it never writes files itself. */
export class ExportModal extends Modal {
  private pathInput: HTMLInputElement | null = null;
  private scopeEl: HTMLElement | null = null;
  private format: ExportFormat;

  constructor(app: App, private options: ExportModalOptions) {
    super(app);
    this.format = options.initialFormat;
  }

  override onOpen(): void {
    this.titleEl.setText("Export");
    const { contentEl } = this;
    contentEl.empty();

    contentEl.createEl("p", {
      text:
        "The path decides the scope: a Markdown note exports one chapter, a folder exports the whole novel (every note underneath, in reading order).",
      cls: "setting-item-description",
    });

    const pathSetting = new Setting(contentEl).setName("Path");
    // The description doubles as a live badge for whatever the path resolves
    // to right now, so it is never set statically.
    this.scopeEl = pathSetting.descEl;
    pathSetting.addText((text) => {
      this.pathInput = text.inputEl;
      text.setPlaceholder("Manuscript/Chapter One.md")
        .setValue(this.options.initialPath);
      new PathSuggest(this.app, text.inputEl);
      text.inputEl.addEventListener("input", () => this.refreshScope());
      text.inputEl.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          this.submit();
        }
      });
    });

    new Setting(contentEl)
      .setName("Format")
      .addDropdown((dropdown) => {
        for (const [value, meta] of Object.entries(EXPORT_FORMATS)) {
          dropdown.addOption(value, meta.label);
        }
        dropdown.setValue(this.format);
        dropdown.onChange((value) => {
          this.format = value as ExportFormat;
          this.options.onFormatChange(this.format);
        });
      });

    new Setting(contentEl)
      .addButton((button) =>
        button.setButtonText("Cancel").onClick(() => this.close())
      )
      .addButton((button) =>
        button.setButtonText("Export").setCta().onClick(() => this.submit())
      );

    this.refreshScope();
    globalThis.setTimeout(() => this.pathInput?.focus(), 0);
  }

  override onClose(): void {
    this.contentEl.empty();
    this.pathInput = null;
    this.scopeEl = null;
  }

  /** Resolve the typed path to an exportable target, or explain why not. */
  private resolve(): { target?: TFile | TFolder; message?: string } {
    const typed = (this.pathInput?.value ?? "").trim();
    // "/" selects the vault root (whole vault = one novel); anything else is
    // trimmed of slashes so "Novels/" and "/Novels" resolve like "Novels".
    const raw = typed.replace(/^\/+/, "").replace(/\/+$/, "");
    if (!raw && typed !== "/") {
      return { message: "Enter a note or folder path." };
    }
    const abstract = typed === "/"
      ? this.app.vault.getRoot()
      : this.app.vault.getAbstractFileByPath(raw);
    if (abstract instanceof TFile) {
      if (abstract.extension !== "md") {
        return { message: "Only Markdown notes can be exported." };
      }
      return { target: abstract };
    }
    if (abstract instanceof TFolder) {
      if (novelFiles(this.app, abstract).length === 0) {
        return { message: "This folder has no Markdown notes." };
      }
      return { target: abstract };
    }
    return { message: `Nothing found at "${raw}".` };
  }

  /** Keep the badge under the path input in sync while typing. */
  private refreshScope(): void {
    if (!this.scopeEl) return;
    const { target, message } = this.resolve();
    if (!target) {
      this.scopeEl.setText(message ?? "");
      return;
    }
    this.scopeEl.setText(
      target instanceof TFile
        ? `Chapter · ${target.basename}.${target.extension}`
        : `Novel · ${novelFiles(this.app, target).length} Markdown note(s)`,
    );
  }

  private submit(): void {
    const { target, message } = this.resolve();
    if (!target) {
      new Notice(`Obsidian Author: ${message}`);
      return;
    }
    this.close();
    this.options.onExport(target, this.format);
  }
}

/** Autocomplete vault folders and Markdown notes inside the path input.
 * Folders come first: the novel scope is the one you can't reach by just
 * having a note open. */
class PathSuggest extends AbstractInputSuggest<TAbstractFile> {
  constructor(
    app: App,
    private input: HTMLInputElement,
  ) {
    super(app, input);
  }

  protected override getSuggestions(query: string): TAbstractFile[] {
    const q = query.trim().replace(/^\/+/, "").toLowerCase();
    const match = (path: string) => !q || path.toLowerCase().includes(q);
    const all = this.app.vault.getAllLoadedFiles();
    const folders = all
      .filter((f): f is TFolder => f instanceof TFolder && match(f.path))
      .sort((a, b) => a.path.localeCompare(b.path));
    const files = all
      .filter(
        (f): f is TFile =>
          f instanceof TFile && f.extension === "md" && match(f.path),
      )
      .sort((a, b) => a.path.localeCompare(b.path));
    return [...folders, ...files].slice(0, 20);
  }

  override renderSuggestion(value: TAbstractFile, el: HTMLElement): void {
    el.createEl("div", { text: value.path || "/" });
  }

  override selectSuggestion(value: TAbstractFile): void {
    this.input.value = value.path;
    this.input.dispatchEvent(new Event("input"));
    this.close();
  }
}
