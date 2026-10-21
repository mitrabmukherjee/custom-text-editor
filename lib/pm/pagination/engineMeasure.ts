import type { Node as PMNode } from "prosemirror-model";
import type { DocumentLayout } from "../format";
import { layoutDocument } from "../word-layout/layoutDocument";
import type { EngineLayoutLine, LayoutResult } from "../word-layout/types";
import type { PageSetup } from "./pageSetup";
import type { ComputedLayout, LayoutLine } from "./computePageBreaks";
export type EngineTabWidth = { pos: number; width: number };

export type EngineMeasureInput = {
  doc: PMNode;
  setup: PageSetup;
  defaultTabStop: number;
  getDocGridLinePitchTwips?: (section: number) => number | undefined;
  contentWidthPx?: number;
  includeDeletionText?: boolean;
};

export type EngineMeasured = {
  lines: (LayoutLine & { ref: { kind: "engine"; pos: number; textOffset: number }; numbered: boolean })[];
  layout: ComputedLayout;
  lineWrapPositions: number[];
  tabWidths: EngineTabWidth[];
};

function mapEngineLine(line: EngineLayoutLine): EngineMeasured["lines"][number] {
  const anchor = line.anchor ?? { pos: 0, textOffset: 0 };
  return {
    top: line.top,
    bottom: line.bottom,
    breakAfter: line.breakAfter,
    breakBefore: line.breakBefore,
    block: line.block,
    section: line.section,
    widowControl: line.widowControl,
    keepNext: line.keepNext,
    keepLines: line.keepLines,
    slack: line.slack,
    rowHeight: line.rowHeight,
    numbered: line.numbered !== false,
    ref: { kind: "engine", pos: anchor.pos, textOffset: anchor.textOffset },
  };
}

export function measureWithEngine(input: EngineMeasureInput): EngineMeasured {
  const layout = input.doc.attrs.layout as DocumentLayout | null;
  const result: LayoutResult = layoutDocument(input.doc, {
    defaultSetup: input.setup,
    defaultTabStop: input.defaultTabStop,
    justifyShrink: layout?.justifyShrink,
    getDocGridLinePitchTwips: input.getDocGridLinePitchTwips,
    contentWidthPx: input.contentWidthPx,
    includeDeletionText: input.includeDeletionText,
  });
  return {
    lines: result.lines.map(mapEngineLine),
    layout: result.pagination,
    lineWrapPositions: result.lineWrapPositions,
    tabWidths: result.tabWidths ?? [],
  };
}
