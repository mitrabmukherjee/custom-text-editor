import type { ComputedLayout } from "./computePageBreaks";

export type LayoutCompareDiff = {
  domPageCount: number;
  enginePageCount: number;
  lineCountDelta: number;
  firstBreakIndex: number | null;
};

export function diffLayouts(
  dom: { pageCount: number; breaks: ComputedLayout["breaks"]; lineCount: number },
  engine: { pageCount: number; breaks: ComputedLayout["breaks"]; lineCount: number },
): LayoutCompareDiff {
  let firstBreakIndex: number | null = null;
  const max = Math.max(dom.breaks.length, engine.breaks.length);
  for (let i = 0; i < max; i++) {
    const a = dom.breaks[i]?.lineIndex;
    const b = engine.breaks[i]?.lineIndex;
    if (a !== b) {
      firstBreakIndex = i;
      break;
    }
  }
  return {
    domPageCount: dom.pageCount,
    enginePageCount: engine.pageCount,
    lineCountDelta: engine.lineCount - dom.lineCount,
    firstBreakIndex,
  };
}

export function diffBreakSpecs(
  a: { pos: number; page: number }[],
  b: { pos: number; page: number }[],
): number | null {
  const max = Math.max(a.length, b.length);
  for (let i = 0; i < max; i++) {
    if (a[i]?.pos !== b[i]?.pos || a[i]?.page !== b[i]?.page) return i;
  }
  return null;
}
