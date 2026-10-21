import type { Node as PMNode } from "prosemirror-model";
import {
  docGridPitchPxFromTwips,
  lineBoxPx,
  lineBoxSlack,
  lineMetricsContextFromFormat,
  PT_TO_PX,
  type ParagraphFormat,
} from "../format";
import { gridStackTop } from "../pagination/gridStack";
import { inlinePieces, inlineWrapMarginRight } from "./effectiveRun";
import type { LayoutTabWidth } from "./types";
import { activeMeasurer, type TextMeasurer } from "./measure";
import type { LayoutRun } from "./types";
import { resolveTabWidth } from "./tabs";
import type { EngineLayoutLine } from "./types";

function verticalPadding(format: ParagraphFormat, onDocGrid: boolean): { top: number; bottom: number } {
  let top = onDocGrid ? 0 : (format.spaceBefore ?? 0);
  let bottom = onDocGrid ? 0 : (format.spaceAfter ?? 0);
  const box = format.boxBorder;
  if (box?.top) top += box.top.paddingPx ?? 0;
  if (box?.bottom) bottom += box.bottom.paddingPx ?? 0;
  return { top, bottom };
}

/** Alternating runs of whitespace and non-whitespace. */
function textTokens(text: string): { text: string; space: boolean }[] {
  const out: { text: string; space: boolean }[] = [];
  const re = /(\s+|[^\s]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    out.push({ text: m[1], space: /^\s+$/.test(m[1]) });
  }
  return out;
}

type LineState = {
  width: number;
  pieces: { text: string; width: number }[];
};

export function layoutParagraph(
  block: PMNode,
  pos: number,
  contentWidth: number,
  naturalTop: number,
  blockId: number,
  section: number,
  options: {
    defaultTabStop: number;
    justifyShrink?: number;
    pitchTwips?: number;
    measurer?: TextMeasurer;
    includeDeletionText?: boolean;
  },
  keep: Partial<EngineLayoutLine>,
): { lines: EngineLayoutLine[]; height: number; tabWidths: LayoutTabWidth[] } {
  const format = (block.attrs.format as ParagraphFormat | null) ?? {};
  const measurer = options.measurer ?? activeMeasurer();
  const padL = format.indentLeft ?? 0;
  const padR = format.indentRight ?? 0;
  const firstIndent = format.indentFirstLine ?? 0;
  const wrapRight = inlineWrapMarginRight(block);
  const justified = format.align === "justify";
  const shrink =
    justified && options.justifyShrink != null && options.justifyShrink > 0
      ? options.justifyShrink
      : 0;
  const maxW = Math.max(8, contentWidth - padL - padR - wrapRight);
  const tabWidths: LayoutTabWidth[] = [];

  const metricsCtx = options.pitchTwips
    ? { docGridLinePitchTwips: options.pitchTwips }
    : lineMetricsContextFromFormat(format);
  const onDocGrid = !!options.pitchTwips || !!format.docGridSectionPitchTwips;
  const vPad = verticalPadding(format, onDocGrid);
  const rowH = lineBoxPx(format, metricsCtx);
  const pitchPx = options.pitchTwips ? docGridPitchPxFromTwips(options.pitchTwips) : undefined;
  const lineAdvance = pitchPx ?? rowH ?? 16;
  /** Painted line box (CSS); pagination still advances one pitch per engine wrap line. */
  const gridRowHeightPx =
    pitchPx && rowH && rowH > pitchPx + 0.51 ? rowH : pitchPx;

  const pieces = inlinePieces(block, pos, {
    includeDeletionText: options.includeDeletionText,
  });
  const lines: EngineLayoutLine[] = [];
  let y = naturalTop + vPad.top;
  let textOffset = 0;
  let isFirstLine = true;
  let gridNext: number | undefined;

  const glyphHeightPx = (format.fontSize ?? 12) * PT_TO_PX;

  const pushLine = (lineWidth: number, anchorOffset: number) => {
    const glyphs = pitchPx ? glyphHeightPx : lineAdvance * 0.72;
    const top = pitchPx ? gridStackTop(y, gridNext) : y;
    if (pitchPx) gridNext = top + pitchPx;
    const height = pitchPx ?? lineAdvance;
    const slack = lineBoxSlack(height, glyphs, format, metricsCtx);
    lines.push({
      top,
      bottom: top + height,
      block: blockId,
      section,
      slack,
      rowHeight: gridRowHeightPx,
      anchor: { pos, textOffset: anchorOffset },
      ...keep,
    });
    y = top + height;
    isFirstLine = false;
  };

  let line: LineState = { width: 0, pieces: [] };
  let lineStartOffset = 0;
  let x = isFirstLine ? firstIndent : 0;
  const lineLimit = () => Math.max(8, maxW - x);

  const flushLine = () => {
    if (line.width > 0 || line.pieces.length) pushLine(line.width, lineStartOffset);
    line = { width: 0, pieces: [] };
    lineStartOffset = textOffset;
    x = 0;
  };

  const layoutText = (text: string, run: LayoutRun) => {
    const tokens = textTokens(text);
    for (let ti = 0; ti < tokens.length; ti++) {
      const { text: tok, space } = tokens[ti];
      const trailingHang = space && ti === tokens.length - 1;
      let tokW = measurer.measureRun(tok, run);
      if (space && shrink > 0) {
        const spaces = tok.length;
        if (spaces > 0) tokW *= 1 - shrink * Math.min(1, spaces / Math.max(1, tok.length));
      }

      if (!trailingHang && x + tokW > lineLimit() && x > 0) flushLine();

      if (!space && tokW > lineLimit() && x === 0) {
        let k = 0;
        while (k < tok.length) {
          let chunk = "";
          let chunkW = 0;
          while (k < tok.length) {
            const ch = tok[k];
            const cw = measurer.measureChar(ch, run);
            if (x + chunkW + cw > lineLimit() && (chunk.length > 0 || x > 0)) break;
            chunk += ch;
            chunkW += cw;
            k += 1;
          }
          if (!chunk.length && k < tok.length) {
            chunk = tok[k];
            chunkW = measurer.measureChar(chunk, run);
            k += 1;
          }
          x += chunkW;
          line.width += chunkW;
          textOffset += chunk.length;
          if (k < tok.length && x >= lineLimit() - 0.5) flushLine();
        }
        continue;
      }

      x += tokW;
      line.width += tokW;
      textOffset += tok.length;
      if (!trailingHang && x >= lineLimit() - 0.5) flushLine();
    }
  };

  for (const piece of pieces) {
    if (piece.kind === "break") {
      flushLine();
      continue;
    }
    if (piece.kind === "tab") {
      const w = resolveTabWidth(
        padL + x,
        format,
        options.defaultTabStop,
        measurer,
        piece.run,
        "",
      );
      tabWidths.push({ pos: piece.pos, width: Math.round(w * 100) / 100 });
      x += w;
      line.width += w;
      continue;
    }
    if (piece.kind === "image") {
      const w = piece.width;
      if (x + w > lineLimit() && x > 0) flushLine();
      x += w;
      line.width += w;
      textOffset += 1;
      continue;
    }
    layoutText(piece.text, piece.run);
  }
  flushLine();

  if (!lines.length) {
    const top = pitchPx ? gridStackTop(y, gridNext) : y;
    const height = pitchPx ?? lineAdvance;
    lines.push({
      top,
      bottom: top + height,
      block: blockId,
      section,
      anchor: { pos, textOffset: 0 },
      rowHeight: gridRowHeightPx,
      ...keep,
    });
    y = top + height;
  }

  y += vPad.bottom;
  const total = y - naturalTop;
  return { lines, height: total, tabWidths };
}
