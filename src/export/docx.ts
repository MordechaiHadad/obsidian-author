import {
  AlignmentType,
  Document as DocxDocument,
  DropCapType,
  FrameAnchorType,
  HeadingLevel,
  LevelFormat,
  LineRuleType,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import type { IParagraphOptions, ISpacingProperties } from "docx";
import type { Block, Chapter, TextRunModel } from "./text.ts";

/** Convert a CSS length to Word twips (1pt = 20 twips).
 * em/rem assume a 12pt body; px assumes 96dpi; % assumes a 6.5" text width.
 * Falls back to 2em (480 twips) for anything unparseable. */
export function lengthToTwips(input: string): number {
  const v = (input ?? "").trim();
  const num = (re: RegExp): number | null => {
    const m = v.match(re);
    return m ? Number(m[1]) : null;
  };
  const em = num(/^(\d+(?:\.\d+)?)(?:em|rem)?$/);
  if (em !== null) return Math.round(em * 12 * 20);
  const px = num(/^(\d+(?:\.\d+)?)px$/);
  if (px !== null) return Math.round(px * 15);
  const pct = num(/^(\d+(?:\.\d+)?)%$/);
  if (pct !== null) return Math.round((pct / 100) * 9360);
  return 480;
}

/** Manuscript body spacing: no gaps between paragraphs (the indent separates
 * them), line height mapped from the unitless setting (×240 twips). */
function bodySpacing(lineHeight: string): ISpacingProperties {
  const parsed = Number((lineHeight ?? "").trim());
  if (!Number.isFinite(parsed) || parsed <= 0) return { after: 0, before: 0 };
  return {
    after: 0,
    before: 0,
    line: Math.round(parsed * 240),
    lineRule: LineRuleType.AUTO,
  };
}

const BULLET_REF = "author-bullet";
const DECIMAL_REF = "author-decimal";

function toTextRuns(runs: TextRunModel[]): TextRun[] {
  return runs.map(
    (r) =>
      new TextRun({
        text: r.text,
        bold: r.bold || undefined,
        italics: r.italic || undefined,
        strike: r.strike || undefined,
        font: r.code ? "Courier New" : undefined,
      }),
  );
}

/** Build a .docx buffer from a single note's blocks (chapter export). */
export async function blocksToDocxBuffer(
  blocks: Block[],
  indentSize: string,
  lineHeight: string,
  enableDropCap = false,
): Promise<ArrayBuffer> {
  return await chaptersToDocxBuffer(
    [{ title: "", blocks }],
    indentSize,
    lineHeight,
    enableDropCap,
  );
}

/** Build a .docx buffer from one or more chapters (novel export). Each
 * chapter restarts the manuscript conventions — flush first paragraph and
 * drop cap — and every chapter after the first starts on a new page.
 * The first body paragraph overall is flush left; the rest carry the
 * first-line indent. */
export async function chaptersToDocxBuffer(
  chapters: Chapter[],
  indentSize: string,
  lineHeight: string,
  enableDropCap = false,
): Promise<ArrayBuffer> {
  const firstLine = lengthToTwips(indentSize);
  const spacing = bodySpacing(lineHeight);
  const children: Paragraph[] = [];
  let indentedYet = false;
  // Applies to the next paragraph actually emitted, so an empty chapter
  // never swallows or misplaces the break.
  let pageBreak = false;
  let emitted = false;
  const push = (opts: IParagraphOptions): void => {
    children.push(
      pageBreak
        ? new Paragraph({ ...opts, pageBreakBefore: true })
        : new Paragraph(opts),
    );
    pageBreak = false;
    emitted = true;
  };

  for (const chapter of chapters) {
    // A chapter opens like a fresh note: its first paragraph is flush and
    // carries the drop cap.
    indentedYet = false;
    if (emitted) pageBreak = true;
    for (const block of chapter.blocks) {
      switch (block.kind) {
        case "heading": {
          const key = `HEADING_${block.level}` as keyof typeof HeadingLevel;
          push({
            heading: HeadingLevel[key],
            children: toTextRuns(block.runs),
          });
          break;
        }
        case "paragraph": {
          const isFirstParagraph = !indentedYet;
          const indent = indentedYet ? firstLine : 0;
          indentedYet = true;
          push({
            indent: { firstLine: indent },
            spacing,
            frame: isFirstParagraph && enableDropCap
              ? {
                type: "absolute",
                position: { x: 0, y: 0 },
                width: 720,
                height: 720,
                anchor: {
                  horizontal: FrameAnchorType.TEXT,
                  vertical: FrameAnchorType.TEXT,
                },
                dropCap: DropCapType.DROP,
                lines: 2,
              }
              : undefined,
            children: toTextRuns(block.runs),
          });
          break;
        }
        case "list": {
          for (const item of block.items) {
            push({
              numbering: {
                reference: block.ordered ? DECIMAL_REF : BULLET_REF,
                level: 0,
              },
              spacing,
              children: toTextRuns(item),
            });
          }
          break;
        }
        case "break": {
          push({
            alignment: AlignmentType.CENTER,
            children: [new TextRun("* * *")],
          });
          break;
        }
      }
    }
  }

  const doc = new DocxDocument({
    numbering: {
      config: [
        {
          reference: BULLET_REF,
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
            },
          ],
        },
        {
          reference: DECIMAL_REF,
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });
  return await Packer.toArrayBuffer(doc);
}
