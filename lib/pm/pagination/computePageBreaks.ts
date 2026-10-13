import { pageStride, type PageSetup } from "./pageSetup";
import type { PageFrame, PageInfo } from "../format";

/** Hard break after a line: `true` for the next page, or the next odd/even page. */
export type BreakKind = boolean | "odd" | "even";

/**
 * One rendered line (or a forced break) in "natural" coordinates: px from the
 * top of the first page, as if no page-break spacers existed.
 */
export type LayoutLine = {
  top: number;
  bottom: number;
  /** Content after this line must start on a new page (hard page or section break). */
  breakAfter?: BreakKind;
  /** The line must start a page (`pageBreakBefore` on its paragraph). */
  breakBefore?: boolean;
  /** Lines of the same paragraph share a block id. */
  block?: number;
  /** 0-based section index (sections are separated by `section_break` nodes). */
  section?: number;
  widowControl?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  /**
   * Leading below the text that may hang past the page bottom. Word puts the extra space
   * of multiple line spacing (e.g. 1.5 lines) under the text and only needs the text to fit.
   */
  slack?: number;
};

export type ComputedBreak = {
  /** Index of the first line on the new page. */
  lineIndex: number;
  /** Spacer height that pushes that line to the next page's content top. */
  height: number;
  forced: boolean;
};

/** Where a line ends up: page index and px from that page's top edge. */
export type LinePlacement = { page: number; top: number; bottom: number; section?: number };

export type ComputedLayout = {
  breaks: ComputedBreak[];
  pageCount: number;
  placements: LinePlacement[];
  /** Section and page number of every page. */
  pages: PageInfo[];
};

export type PaginateOptions = {
  /** `w:pgNumType w:start` of a section, if it restarts numbering. */
  numberStart?: (section: number) => number | undefined;
};

function blockBounds(lines: LayoutLine[]) {
  const first = new Map<number, number>();
  const last = new Map<number, number>();
  lines.forEach((line, i) => {
    if (line.block === undefined) return;
    if (!first.has(line.block)) first.set(line.block, i);
    last.set(line.block, i);
  });
  const firstOf = (i: number) =>
    lines[i].block === undefined ? i : first.get(lines[i].block!)!;
  const lastOf = (i: number) =>
    lines[i].block === undefined ? i : last.get(lines[i].block!)!;
  return { firstOf, lastOf };
}

/**
 * Word-style pagination: a line that would cross the bottom of the page's
 * text area moves to the next page, so paragraphs split between lines.
 * Widow/orphan control, keep-lines-together and keep-with-next move the
 * break earlier when the page still keeps at least one line. A line taller
 * than a page is left to overflow rather than looping. Odd/even section
 * breaks leave a blank page when needed.
 */
export function computePageBreaks(
  lines: LayoutLine[],
  setup: PageSetup,
  frame: (page: PageInfo) => PageFrame = () => ({
    top: setup.marginTop,
    bottom: setup.marginBottom,
  }),
  options: PaginateOptions = {},
): ComputedLayout {
  const stride = pageStride(setup);
  const { firstOf, lastOf } = blockBounds(lines);
  const breaks: ComputedBreak[] = [];
  const placements: LinePlacement[] = new Array(lines.length);
  const pages: PageInfo[] = [];
  let page = 0;
  let shift = 0;
  let pageStart = 0;
  let forceNext: BreakKind = false;

  const startPage = (index: number, section: number) => {
    const prev = pages[pages.length - 1];
    const first = !prev || prev.section !== section;
    const restart = first ? options.numberStart?.(section) : undefined;
    // Blank pages left by odd/even breaks belong to the previous section.
    while (pages.length < index) {
      const last = pages[pages.length - 1];
      pages.push({
        section: last?.section ?? section,
        first: false,
        number: (last?.number ?? 0) + 1,
      });
    }
    const before = pages[pages.length - 1];
    pages[index] = { section, first, number: restart ?? (before ? before.number + 1 : 1) };
  };
  startPage(0, lines[0]?.section ?? 0);

  const keepTarget = (i: number): number => {
    let t = i;
    const line = lines[t];
    const first = firstOf(t);
    const last = lastOf(t);
    if (line.keepLines) {
      t = first;
    } else if (line.widowControl && last > first) {
      if (last === t && last - first >= 2) t -= 1;
      if (t - first === 1) t = first;
    }
    while (t === firstOf(t) && t > 0 && lines[t - 1].keepNext && !lines[t - 1].breakAfter) {
      t = firstOf(t - 1);
    }
    return t > pageStart ? t : i;
  };

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const pageBottom = page * stride + setup.height - frame(pages[page]).bottom;
    const overflows = line.bottom - (line.slack ?? 0) + shift > pageBottom + 0.5;
    const breakBefore = !!line.breakBefore && i > pageStart && firstOf(i) === i;

    if (forceNext || breakBefore || (overflows && i > pageStart)) {
      const forced = !!forceNext || breakBefore;
      const target = forced ? i : keepTarget(i);
      page += 1;
      if (forceNext === "odd" && (page + 1) % 2 === 0) page += 1;
      if (forceNext === "even" && (page + 1) % 2 === 1) page += 1;
      startPage(page, lines[target].section ?? 0);
      const height = page * stride + frame(pages[page]).top - (lines[target].top + shift);
      breaks.push({ lineIndex: target, height, forced });
      shift += height;
      pageStart = target;
      forceNext = false;
      i = target;
    }

    const placed = lines[i];
    placements[i] = {
      page,
      top: placed.top + shift - page * stride,
      bottom: placed.bottom + shift - page * stride,
      section: placed.section ?? 0,
    };
    if (placed.breakAfter) forceNext = placed.breakAfter;
    i += 1;
  }

  // A hard break at the very end still starts an (empty) new page.
  if (forceNext) {
    page += 1;
    startPage(page, pages[page - 1].section);
  }

  return { breaks, pageCount: page + 1, placements, pages };
}
