import type { Mark, Node as PMNode } from "prosemirror-model";
import type { ParagraphFormat } from "../format";
import type { LayoutRun } from "./types";

const DEFAULT_FAMILY = "Times New Roman";
const DEFAULT_SIZE = 12;

function markAttrs(mark: Mark): Record<string, unknown> {
  return mark.attrs as Record<string, unknown>;
}

/** Effective character formatting at a document offset (paragraph mark + marks). */
export function effectiveRunAt(
  block: PMNode,
  inlineOffset: number,
  marks: readonly Mark[],
): LayoutRun {
  const format = (block.attrs.format as ParagraphFormat | null) ?? {};
  let fontFamily = format.fontFamily ?? DEFAULT_FAMILY;
  let fontSizePt = format.fontSize ?? DEFAULT_SIZE;
  let bold = false;
  let italic = false;
  let fontKern = false;

  for (const m of marks) {
    if (m.type.name === "strong") bold = true;
    if (m.type.name === "em") italic = true;
    if (m.type.name === "caps") {
      const mode = (m.attrs as { mode?: string }).mode;
      if (mode === "small") fontSizePt *= 0.8;
    }
    if (m.type.name === "textStyle") {
      const a = markAttrs(m);
      if (a.fontFamily) fontFamily = String(a.fontFamily);
      if (a.fontSize) fontSizePt = Number(a.fontSize);
      if (a.fontKern) fontKern = true;
    }
  }

  return { text: "", fontFamily, fontSizePt, bold, italic, fontKern };
}

export function collectInlineRuns(block: PMNode): { text: string; run: LayoutRun }[] {
  const out: { text: string; run: LayoutRun }[] = [];
  let offset = 0;
  block.forEach((child) => {
    if (child.isText) {
      const t = child.text ?? "";
      out.push({ text: t, run: effectiveRunAt(block, offset, child.marks) });
      offset += t.length;
    } else if (child.type.name === "tab") {
      out.push({ text: "\t", run: effectiveRunAt(block, offset, child.marks) });
      offset += 1;
    } else if (child.type.name === "hard_break") {
      out.push({ text: "\n", run: effectiveRunAt(block, offset, child.marks) });
      offset += 1;
    } else if (child.type.name === "image") {
      const w = Number(child.attrs.width) || 48;
      out.push({
        text: "\ufffc",
        run: { ...effectiveRunAt(block, offset, child.marks), text: "" },
      });
      (out[out.length - 1] as { imageWidth?: number }).imageWidth = w;
      offset += 1;
    }
  });
  return out;
}

export type InlinePiece =
  | { kind: "text"; text: string; run: LayoutRun }
  | { kind: "tab"; pos: number; run: LayoutRun }
  | { kind: "break" }
  | { kind: "image"; width: number; run: LayoutRun; anchored: boolean };

export function inlineWrapMarginRight(block: PMNode): number {
  let margin = 0;
  block.forEach((child) => {
    if (child.type.name !== "image") return;
    if (child.attrs.anchorX != null || child.attrs.anchorY != null || child.attrs.vmlAnchor) return;
    margin = Math.max(margin, Number(child.attrs.width) || 48);
  });
  return margin;
}

export type InlinePiecesOptions = {
  includeDeletionText?: boolean;
};

export function inlinePieces(
  block: PMNode,
  blockPos = 0,
  options: InlinePiecesOptions = {},
): InlinePiece[] {
  const includeDeletionText = options.includeDeletionText ?? true;
  const pieces: InlinePiece[] = [];
  let offset = 0;
  block.forEach((child, childOffset) => {
    const childPos = blockPos + 1 + childOffset;
    if (child.isText) {
      if (!includeDeletionText && child.marks.some((m) => m.type.name === "deletion")) {
        offset += (child.text ?? "").length;
        return;
      }
      const t = child.text ?? "";
      pieces.push({ kind: "text", text: t, run: effectiveRunAt(block, offset, child.marks) });
      offset += t.length;
    } else if (child.type.name === "tab") {
      pieces.push({ kind: "tab", pos: childPos, run: effectiveRunAt(block, offset, child.marks) });
      offset += 1;
    } else if (child.type.name === "hard_break") {
      pieces.push({ kind: "break" });
      offset += 1;
    } else if (child.type.name === "image") {
      pieces.push({
        kind: "image",
        width: Number(child.attrs.width) || 48,
        run: effectiveRunAt(block, offset, child.marks),
        anchored: child.attrs.anchorX != null || child.attrs.anchorY != null || !!child.attrs.vmlAnchor,
      });
      offset += 1;
    }
  });
  return pieces;
}
