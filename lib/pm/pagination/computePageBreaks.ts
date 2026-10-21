import { pageStride, type PageSetup } from "./pageSetup";
import type { PageFrame, PageInfo } from "../format";
import { computePageBreaksWithContext } from "./geometryPaginate";

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
  /** Word doc-grid row height (px); margin numbers use this instead of glyph-inflated bottom-top. */
  rowHeight?: number;
};

export type ComputedBreak = {
  /** Index of the first line on the new page. */
  lineIndex: number;
  /** Spacer height that pushes that line to the next page's content top. */
  height: number;
  forced: boolean;
};

/** Where a line ends up: page index and px from that page's top edge. */
export type LinePlacement = {
  page: number;
  top: number;
  bottom: number;
  section?: number;
  rowHeight?: number;
};

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

export { computePageBreaksWithContext } from "./geometryPaginate";

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
  const bound = {
    getPageGeometry: (_pageIndex: number) => ({
      width: setup.width,
      height: setup.height,
      gap: setup.gap,
      sectionIndex: 0,
      marginTop: setup.marginTop,
      marginRight: setup.marginRight,
      marginBottom: setup.marginBottom,
      marginLeft: setup.marginLeft,
      frames: [
        {
          left: setup.marginLeft,
          width: setup.width - setup.marginLeft - setup.marginRight,
        },
      ],
    }),
    getFrame: frame,
  };
  return computePageBreaksWithContext(lines, bound, options, setup);
}
