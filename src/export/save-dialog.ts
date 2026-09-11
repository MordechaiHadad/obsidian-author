/// <reference lib="dom" />
import type { App } from "obsidian";

declare function require(id: string): unknown;

export interface SaveFilter {
  name: string;
  extensions: string[];
}

export type SaveChoice =
  | { kind: "chosen"; path: string }
  | { kind: "cancelled" }
  | { kind: "unavailable" };

/** Join a vault-absolute base path with a vault-relative path. Pure and
 * unit-tested. */
export function toAbsolutePath(
  basePath: string,
  vaultPath: string,
): string {
  const base = (basePath ?? "").replace(/\/+$/g, "");
  const rel = (vaultPath ?? "").replace(/^\/+/g, "");
  return rel ? `${base}/${rel}` : base;
}

function nodeRequire(id: string): unknown {
  try {
    const w = window as unknown as { require?: (id: string) => unknown };
    if (typeof w.require === "function") return w.require(id);
  } catch {
    // Fall through to the bundled require below.
  }
  return require(id);
}

interface ElectronDialog {
  showSaveDialog?: (
    options: Record<string, unknown>,
  ) => Promise<{ canceled?: boolean; filePath?: string }>;
}

function getDialog(): ElectronDialog | null {
  try {
    const w = window as unknown as Record<string, unknown>;
    const electron = w["electron"] as
      | { remote?: { dialog?: ElectronDialog } }
      | undefined;
    if (electron?.remote?.dialog) return electron.remote.dialog;
    const mod = nodeRequire("@electron/remote") as {
      dialog?: ElectronDialog;
    };
    if (mod?.dialog) return mod.dialog;
  } catch {
    // Mobile or unavailable: caller falls back to vault-relative save.
  }
  return null;
}

function vaultBasePath(app: App): string | null {
  try {
    const adapter = app.vault.adapter as unknown as
      | { getBasePath?: () => string }
      | undefined;
    const base = adapter?.getBasePath?.();
    return base || null;
  } catch {
    return null;
  }
}

/** Open the OS Save dialog seeded at the suggested vault-relative path.
 * Returns "unavailable" on mobile / without Electron so the caller can
 * fall back to saving next to the note. */
export async function chooseSavePath(
  app: App,
  suggestedVaultPath: string,
  filters: SaveFilter[],
): Promise<SaveChoice> {
  const dialog = getDialog();
  const base = vaultBasePath(app);
  if (!dialog?.showSaveDialog || !base) return { kind: "unavailable" };
  try {
    const result = await dialog.showSaveDialog({
      title: "Export manuscript",
      defaultPath: toAbsolutePath(base, suggestedVaultPath),
      filters: filters.map((f) => ({
        name: f.name,
        extensions: f.extensions,
      })),
      properties: ["createDirectory", "showOverwriteConfirmation"],
    });
    if (result?.canceled || !result?.filePath) return { kind: "cancelled" };
    return { kind: "chosen", path: result.filePath };
  } catch {
    return { kind: "unavailable" };
  }
}

/** Write a buffer to an absolute OS path (Save-dialog target, possibly
 * outside the vault). Desktop only; throws when Node fs is unavailable. */
export async function writeAbsoluteFile(
  fullPath: string,
  data: ArrayBuffer,
): Promise<void> {
  const fs = nodeRequire("fs") as {
    promises?: { writeFile?: (path: string, data: Uint8Array) => Promise<void> };
  };
  const writeFile = fs?.promises?.writeFile;
  if (typeof writeFile !== "function") {
    throw new Error("file system access is unavailable.");
  }
  await writeFile(fullPath, new Uint8Array(data));
}
