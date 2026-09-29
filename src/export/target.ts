import type { App, TFile, TFolder } from "obsidian";

/** Reading order for a novel's chapters: path-wise, naturally (numeric) so
 * `Ch2.md` sorts before `Ch10.md`, case-insensitively. Pure — unit tested. */
export function compareChapterPaths(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

/** Every Markdown note under a folder, subfolders included, in reading
 * order. The vault root (`path === "/"`) means the whole vault. */
export function novelFiles(app: App, folder: TFolder): TFile[] {
  const prefix = folder.path === "/" ? "" : `${folder.path}/`;
  return app.vault.getMarkdownFiles()
    .filter((file) => file.path.startsWith(prefix))
    .sort((a, b) => compareChapterPaths(a.path, b.path));
}
