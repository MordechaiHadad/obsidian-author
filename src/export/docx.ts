import {
  AlignmentType,
  Document as DocxDocument,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import type { Block, TextRunModel } from "./model.ts";

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

/** Build a .docx buffer from blocks. The first body paragraph is flush left
 * (manuscript convention); the rest carry the first-line indent. */
export async function blocksToDocxBuffer(
  blocks: Block[],
  indentSize: string,
): Promise<ArrayBuffer> {
  const firstLine = lengthToTwips(indentSize);
  const children: Paragraph[] = [];
  let indentedYet = false;

  for (const block of blocks) {
    switch (block.kind) {
      case "heading": {
        const key = `HEADING_${block.level}` as keyof typeof HeadingLevel;
        children.push(
          new Paragraph({
            heading: HeadingLevel[key],
            children: toTextRuns(block.runs),
          }),
        );
        break;
      }
      case "paragraph": {
        const indent = indentedYet ? firstLine : 0;
        indentedYet = true;
        children.push(
          new Paragraph({
            indent: { firstLine: indent },
            children: toTextRuns(block.runs),
          }),
        );
        break;
      }
      case "list": {
        for (const item of block.items) {
          children.push(
            new Paragraph({
              numbering: {
                reference: block.ordered ? DECIMAL_REF : BULLET_REF,
                level: 0,
              },
              children: toTextRuns(item),
            }),
          );
        }
        break;
      }
      case "break": {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun("* * *")],
          }),
        );
        break;
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
