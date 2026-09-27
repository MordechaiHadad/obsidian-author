/// <reference lib="dom" />
/** Live Preview drop cap as a CodeMirror 6 line decoration. */

import { RangeSetBuilder } from "@codemirror/state";
import {
  Decoration,
  DecorationSet,
  EditorView,
  ViewPlugin,
  ViewUpdate,
} from "@codemirror/view";
import { editorInfoField, editorLivePreviewField } from "obsidian";
import { isManuscriptPath, normalizeFolder } from "./scope.ts";

/** Line-decoration class consumed by styles.css. */
const DROPCAP_LINE_CLASS = "author-dropcap-line";

const dropcapMark = Decoration.line({ class: DROPCAP_LINE_CLASS });
const noDecorations = Decoration.none;

let manuscriptFolder = "";
let scopeRevision = 0;

/** Update the folder predicate; returns true only when it actually changes. */
export function setDropcapFolder(folder: string): boolean {
  const normalized = normalizeFolder(folder);
  if (normalized === manuscriptFolder) return false;
  manuscriptFolder = normalized;
  scopeRevision++;
  return true;
}

/** True only for a manuscript file in Obsidian Live Preview. */
function isEligible(view: EditorView): boolean {
  if (!view.state.field(editorLivePreviewField, false)) return false;
  const info = view.state.field(editorInfoField, false);
  const file = info?.file;
  return !!file && isManuscriptPath(file.path, manuscriptFolder);
}

/** True for source lines that must not receive the manuscript's cap. */
function isBlockedLine(text: string): boolean {
  const t = text.trim();
  return t === "" || /^#{1,6}(\s|$)/.test(t) ||
    /^(\*\*\*|___|---)(\s|$)/.test(t) ||
    /^\s*([-*+]|\d+[.)])\s+/.test(text) || /^\s*>/.test(text) ||
    /^\s*```/.test(text) || /^\s*\|/.test(text);
}

/** Find the first prose line, ignoring a leading YAML frontmatter block. */
function firstProseLineNumber(doc: {
  lines: number;
  line(n: number): { text: string };
}): number | null {
  let inFrontmatter = false;
  for (let n = 1; n <= doc.lines; n++) {
    const text = doc.line(n).text;
    if (n === 1 && text.trim() === "---") {
      inFrontmatter = true;
      continue;
    }
    if (inFrontmatter) {
      if (text.trim() === "---" || text.trim() === "...") {
        inFrontmatter = false;
      }
      continue;
    }
    if (!isBlockedLine(text)) return n;
  }
  return null;
}

export const dropcapExtension = ViewPlugin.fromClass(
  class {
    decorations: DecorationSet = noDecorations;
    private eligible = false;
    private revision = scopeRevision;
    private markerTo: number | null = null;

    constructor(view: EditorView) {
      this.eligible = isEligible(view);
      if (this.eligible) this.rebuild(view);
    }

    update(update: ViewUpdate) {
      const nowEligible = isEligible(update.view);
      const folderChanged = this.revision !== scopeRevision;

      if (!nowEligible) {
        this.decorations = noDecorations;
        this.eligible = false;
        this.revision = scopeRevision;
        this.markerTo = null;
        return;
      }

      // Entering Live Preview, opening a manuscript, or changing the folder
      // predicate gets one fresh marker calculation.
      if (!this.eligible || folderChanged) {
        this.eligible = true;
        this.revision = scopeRevision;
        this.rebuild(update.view);
        return;
      }

      if (!update.docChanged) return;

      // The marker is always before unrelated edits below the first prose
      // line. Only rescan when a transaction can affect the candidate line
      // or text before it (including frontmatter delimiters).
      if (this.markerTo === null) {
        this.rebuild(update.view);
        return;
      }
      let affectsMarker = false;
      update.changes.iterChanges((fromA) => {
        if (fromA <= this.markerTo!) affectsMarker = true;
      });
      if (affectsMarker) this.rebuild(update.view);
    }

    private rebuild(view: EditorView): void {
      const builder = new RangeSetBuilder<Decoration>();
      const n = firstProseLineNumber(view.state.doc);
      if (n === null) {
        this.decorations = builder.finish();
        this.markerTo = null;
        return;
      }
      const line = view.state.doc.line(n);
      builder.add(line.from, line.from, dropcapMark);
      this.decorations = builder.finish();
      this.markerTo = line.to;
    }
  },
  { decorations: (plugin) => plugin.decorations },
);
