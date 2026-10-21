import type { LayoutLine } from "../pagination/computePageBreaks";
import type { PageSetup } from "../pagination/pageSetup";
import type { PageFrame, PageInfo } from "../format";

/** One text column on a page (`w:cols`). */
export type ContentFrame = {
  left: number;
  width: number;
};

/** Screen + print geometry for a single page index. */
export type PageGeometry = {
  width: number;
  height: number;
  /** Gap after this page on screen (not printed). */
  gap: number;
  sectionIndex: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  frames: ContentFrame[];
};

export type PaginationContext = {
  getPageGeometry: (pageIndex: number, pageInfo: PageInfo | undefined) => PageGeometry;
  getFrame: (page: PageInfo) => PageFrame;
};

export type GlyphBox = {
  char: string;
  advance: number;
};

export type LayoutRun = {
  text: string;
  fontFamily: string;
  fontSizePt: number;
  bold?: boolean;
  italic?: boolean;
  fontKern?: boolean;
};

export type LineAnchor = {
  /** ProseMirror position at line start. */
  pos: number;
  /** Text offset within the block for this line. */
  textOffset: number;
};

export type EngineLayoutLine = LayoutLine & {
  column?: number;
  anchor?: LineAnchor;
  /** When false, excluded from margin line numbering (tables, breaks, suppressed). */
  numbered?: boolean;
};

export type BlockPlacement = {
  pos: number;
  top: number;
  bottom: number;
  kind: "paragraph" | "heading" | "table_row" | "break";
};

export type LayoutTabWidth = { pos: number; width: number };

export type LayoutBodyResult = {
  lines: EngineLayoutLine[];
  blocks: BlockPlacement[];
  tabWidths: LayoutTabWidth[];
};

export type LayoutBodyOptions = {
  defaultSetup: PageSetup;
  getDocGridLinePitchTwips?: (section: number) => number | undefined;
  justifyShrink?: number;
  defaultTabStop?: number;
  /** When set, overrides frame width for line breaking (match the live ProseMirror content box). */
  contentWidthPx?: number;
  /** When false, deleted revision text is omitted (matches hide-deletions markup). */
  includeDeletionText?: boolean;
};

export type DocumentCompat = {
  compatibilityMode?: number;
  htmlAutoSpacing?: boolean;
};

/** Stable golden record for regression (no pixel positions). */
export type LayoutGoldenPages12 = {
  pageCount: number;
  perPageLines: [number, number];
  breakLineIndices: number[];
  lineAnchorHashPages12: string;
};

export type LayoutGolden = {
  fixture: string;
  pageCount: number;
  lineCount: number;
  /** FNV-style hash of line-start anchors. */
  lineAnchorHash: string;
  perPageLines: number[];
  /** First-line indices where pagination inserts a page break (stable regression signal). */
  breakLineIndices?: number[];
  /** Title + TOC pages only (Papasotiriou regression). */
  pages12?: LayoutGoldenPages12;
};

export type LayoutResult = {
  lines: EngineLayoutLine[];
  pagination: import("../pagination/computePageBreaks").ComputedLayout;
  /** Document positions where the engine forces a soft line break in the DOM. */
  lineWrapPositions: number[];
  tabWidths: LayoutTabWidth[];
};
