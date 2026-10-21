import type { Node as PMNode } from "prosemirror-model";
import {
  lineBoxPx,
  lineMetricsContextFromFormat,
  type ParagraphFormat,
  type SectionLayout,
} from "../format";
import { gridStackTop } from "./gridStack";
import type { PageSetup } from "./pageSetup";
import type { LayoutLine } from "./computePageBreaks";

/** Section-indexed char wrap widths used for PAPASOTIRIOU-style grid simulation. */
export const PAPASOTIRIOU_CHARS_PER_LINE = [94, 88, 88, 88] as const;

/** Fixed-pitch wrap width from the laid-out paragraph box (Courier ~7.09px/char at 12pt). */
export function gridWrapChars(el: HTMLElement, scale: number): number {
  const w = el.getBoundingClientRect().width / scale;
  if (w < 8) return 88;
  return Math.max(32, Math.round(w / 7.09));
}

export function estLineCount(text: string, chars: number): number {
  let n = 0;
  for (const part of text.split("\n")) {
    const t = part.trimEnd();
    n += t ? Math.max(1, Math.ceil(t.length / chars)) : 1;
  }
  return Math.max(1, n);
}

/** Rough visual line texts matching `estLineCount` wrapping. */
export function estLineTexts(text: string, chars: number): string[] {
  const out: string[] = [];
  for (const part of text.split("\n")) {
    const t = part.trimEnd();
    if (!t) {
      out.push("");
      continue;
    }
    for (let i = 0; i < t.length; i += chars) {
      out.push(t.slice(i, i + chars));
    }
  }
  return out.length ? out : [""];
}

export type SimulatedGridOptions = {
  /** Y offset (px) before the first simulated line on page 1. */
  startY?: number;
  charsPerSection?: readonly number[];
};

/**
 * Node-side line model: walks the doc like pagination `measure()` for text blocks and
 * tables, but uses char-count wrapping instead of DOM `getClientRects()`.
 */
export function buildSimulatedGridLayoutLines(
  root: PMNode,
  setup: PageSetup,
  sections: SectionLayout[],
  options: SimulatedGridOptions = {},
): { lines: LayoutLine[]; lineTexts: string[] } {
  const charsPerSection = options.charsPerSection ?? PAPASOTIRIOU_CHARS_PER_LINE;
  const lines: LayoutLine[] = [];
  const lineTexts: string[] = [];
  let si = 0;
  const gridNext = new Map<number, number>();
  let y = options.startY ?? 115;

  const pushTextLines = (text: string, f: ParagraphFormat | null, section: number) => {
    const ctx = lineMetricsContextFromFormat(f);
    const pitch = sections[section]?.docGridLinePitch ? (lineBoxPx(f, ctx) ?? 24) : 0;
    const box = lineBoxPx(f, ctx) ?? 20;
    const chars = charsPerSection[Math.min(section, charsPerSection.length - 1)] ?? 88;
    const chunks = estLineTexts(text, chars);
    for (const chunk of chunks) {
      if (pitch) {
        const prev = gridNext.get(section);
        const top = gridStackTop(y, prev);
        gridNext.set(section, top + pitch);
        lines.push({
          top,
          bottom: top + pitch,
          section,
          slack: 0,
          rowHeight: pitch,
        });
        y = top + pitch;
      } else {
        lines.push({ top: y, bottom: y + box, section, slack: 0 });
        y += box;
      }
      lineTexts.push(chunk);
    }
  };

  root.descendants((node) => {
    if (node.type.name === "section_break") {
      si += 1;
      return;
    }
    if (node.type.name === "table") {
      node.forEach((row) => {
        let text = "";
        row.descendants((cell) => {
          if (cell.isTextblock) text += cell.textContent + " ";
        });
        pushTextLines(text, null, si);
      });
      return false;
    }
    if (!node.isTextblock) return;
    pushTextLines(node.textContent, node.attrs.format as ParagraphFormat | null, si);
  });

  return { lines, lineTexts };
}
