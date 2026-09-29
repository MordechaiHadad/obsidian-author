# Obsidian Author

Writer's toolkit for Obsidian: manuscript typography for all notes under a chosen folder.

Apply classic novel-style formatting: first-line indent, tight paragraph rhythm, configurable line height, automatically to Markdown notes in your manuscript folder, in Reading view, Live Preview, and print/PDF export.

## Features

- **Folder-scoped manuscript styling**: only notes under the configured manuscript folder (subfolders included) get typography; everything else stays untouched
- **First-line indent**: novel-style paragraph indent in Reading view and Live Preview, with configurable size (`2em`, `24px`, `5%`, etc.)
- **Novel conventions**: first paragraph of note starts flush left; optional flush-left after headings and horizontal rules
- **Configurable line height**: unitless multiplier applied to manuscript notes only
- **Status bar indicator**: `✒ Manuscript · ~N print pages` live estimate (250 words/page, debounced) for the active manuscript note
- **One Export command with a modal**: `Export…` opens a dialog where the path decides the scope and a dropdown decides the format:
  - **Path** — a Markdown file exports a **chapter**, a folder exports the whole **novel** (every note underneath, recursively, in reading order: `Ch2` before `Ch10`)
  - **Format** — DOCX (Word document), EPUB (e-book) or PDF (manuscript via hidden render); the choice is remembered
  - The OS save picker still opens so you choose exactly where the file lands (vault-relative fallback on mobile)
  - Novel exports are one section per chapter: an EPUB TOC entry per note, a page break between chapters in DOCX and PDF
  - Each chapter opens with a heading carrying its file name (skipped when the note already starts with that heading), in every format
