import type { PageSetup } from "./pagination/pageSetup";
import { TWIPS_PER_PX } from "./pagination/pageSetup";

/** Word's single line height is about 1.15x the font size for Times New Roman and Arial. */
export const SINGLE_LINE_FACTOR = 1.15;

export const PT_TO_PX = 96 / 72;

export type TextAlign = "left" | "center" | "right" | "justify";

export type TabStop = { pos: number; align: "left" | "center" | "right" | "decimal" };

/** One edge of a Word paragraph box border (`w:pBdr`). */
export type ParagraphBorderSide = {
  style: "solid" | "dashed" | "dotted" | "double";
  /** Border width in pt (`w:sz` is eighths of a point). */
  widthPt: number;
  color: string;
  /** Gap between border and text, px (`w:space` is eighths of a point, like `w:sz`). */
  paddingPx?: number;
};

/** Resolved Word paragraph formatting. Lengths are CSS px, font sizes pt. */
export type ParagraphFormat = {
  align?: TextAlign;
  indentLeft?: number;
  indentRight?: number;
  /** Positive: first-line indent. Negative: hanging indent. */
  indentFirstLine?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  /** "auto": `lineValue` is a multiple of single spacing; otherwise px. */
  lineRule?: "auto" | "exact" | "atLeast";
  lineValue?: number;
  /** Font of the paragraph mark; sets the height of empty lines. */
  fontFamily?: string;
  fontSize?: number;
  tabs?: TabStop[];
  widowControl?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  suppressLineNumbers?: boolean;
  /** Start the paragraph on a new page. */
  pageBreakBefore?: boolean;
  /** Ignore spacing between this and adjacent paragraphs of the same style. */
  contextualSpacing?: boolean;
  /** Word paragraph style id (e.g. "Heading1"); see `DocumentLayout.styles`. */
  styleId?: string;
  /** Paragraph box border (`w:pBdr`), e.g. certification form outlines. */
  boxBorder?: Partial<
    Record<"top" | "right" | "bottom" | "left", ParagraphBorderSide>
  >;
  /**
   * Section `w:docGrid` pitch (twips) when this paragraph lives in a grid section.
   * Set on import so CSS line metrics match pagination without a live section index.
   */
  docGridSectionPitchTwips?: number;
};

export type TableAlign = "left" | "center" | "right";

/** Per-edge table cell border (`w:tcBorders`); `"none"` is explicit `w:val="nil"`. */
export type CellBorderEdges = Partial<
  Record<"top" | "right" | "bottom" | "left", ParagraphBorderSide | "none">
>;

export function cellBorderCss(edges: CellBorderEdges): string {
  const css: string[] = [];
  const keys = (["top", "right", "bottom", "left"] as const).filter((s) => edges[s] !== undefined);
  if (!keys.length) return "";
  css.push("border: none !important");
  for (const side of keys) {
    const edge = edges[side];
    if (edge === "none") css.push(`border-${side}: none !important`);
    else if (edge) css.push(`border-${side}: ${borderCss(edge)} !important`);
  }
  return css.join("; ");
}

export type CellMarginPx = { top?: number; right?: number; bottom?: number; left?: number };

export function cellMarginCss(margin: CellMarginPx | null | undefined): string {
  if (!margin) return "";
  const parts: string[] = [];
  if (margin.top) parts.push(`padding-top: ${round(margin.top)}px`);
  if (margin.right) parts.push(`padding-right: ${round(margin.right)}px`);
  if (margin.bottom) parts.push(`padding-bottom: ${round(margin.bottom)}px`);
  if (margin.left) parts.push(`padding-left: ${round(margin.left)}px`);
  return parts.join("; ");
}

/** A tracked paragraph formatting revision (`w:pPrChange`), stored on the paragraph. */
export type ParagraphChange = {
  id: string;
  authorId: string;
  authorName: string;
  date: string;
  /** Format before the revision; rejecting restores it. */
  old: ParagraphFormat | null;
  /** Block type before the revision, when a style change turned e.g. a paragraph into a heading. */
  block?: { type: string; level?: number };
  description: string;
};

type HFRunStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** "#rrggbb". */
  color?: string;
  fontFamily?: string;
  fontSize?: number;
};

export type HFRun =
  | ({ kind: "text"; text: string } & HFRunStyle)
  | ({ kind: "field"; field: "PAGE" | "NUMPAGES" | "SECTIONPAGES" } & HFRunStyle)
  | { kind: "tab" };

/** Position of a Word frame (`w:framePr`). Lengths are px. */
export type HFFrame = {
  hAnchor: "page" | "margin" | "text";
  vAnchor: "page" | "margin" | "text";
  /** Offset from the anchor, used when there is no alignment. */
  x?: number;
  y?: number;
  xAlign?: "left" | "center" | "right" | "inside" | "outside";
  yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
  width?: number;
  height?: number;
};

export type HFParagraph = {
  format: ParagraphFormat;
  runs: HFRun[];
  /** Positioned frame (w:framePr): drawn but takes no space. */
  floating?: boolean;
  frame?: HFFrame;
};

/** Header/footer content block (paragraph or single-cell table with inside borders). */
export type HFBlock =
  | ({ kind?: "paragraph" } & HFParagraph)
  | {
      kind: "table";
      widthPx?: number;
      insideBorder?: ParagraphBorderSide;
      paragraphs: HFParagraph[];
    };

export type HeaderFooterSet = {
  default?: HFBlock[];
  first?: HFBlock[];
  /** Even pages when `DocumentLayout.evenAndOddHeaders` is on. */
  even?: HFBlock[];
};

export function isHFTable(block: HFBlock): block is Extract<HFBlock, { kind: "table" }> {
  return block.kind === "table";
}

export function hfBlockParagraphs(block: HFBlock): HFParagraph[] {
  if (isHFTable(block)) return block.paragraphs;
  return [block];
}

export function hfSetParagraphs(set: HFBlock[] | undefined): HFParagraph[] {
  return (set ?? []).flatMap(hfBlockParagraphs);
}

export type LineNumbering = {
  countBy: number;
  /** Number shown on the first counted line. */
  start: number;
  restart: "newPage" | "newSection" | "continuous";
  /** Gap between the number and the text, px. */
  distance: number;
  fontFamily?: string;
  fontSize?: number;
};

/** How a section starts (`w:type` of `w:sectPr`). */
export type SectionStart = "nextPage" | "continuous" | "evenPage" | "oddPage";

/** Page layout of one Word section (`w:sectPr`). */
export type SectionLayout = {
  page?: PageSetup;
  headerDistance?: number;
  footerDistance?: number;
  headers?: HeaderFooterSet;
  footers?: HeaderFooterSet;
  titlePage?: boolean;
  lineNumbers?: LineNumbering | null;
  /** Restart page numbering at this value. */
  pageNumberStart?: number;
  /** Text columns (`w:cols`); 1 when absent. */
  columns?: { count: number; space: number };
  /** `w:docGrid w:linePitch` (twips), when present. */
  docGridLinePitch?: number;
};

/** A paragraph or character style from `styles.xml`, flattened along `basedOn`. */
export type StyleDefinition = {
  id: string;
  name: string;
  type: "paragraph" | "character";
  format?: ParagraphFormat;
  /** Character formatting the style adds (RunFormat fields). */
  run?: Record<string, unknown>;
  /** Style of the paragraph Enter creates after this one. */
  next?: string;
  /** Shown in the styles gallery (Word's `w:qFormat`). */
  gallery?: boolean;
};

/**
 * Document-level Word layout, stored on the doc node's `layout` attribute. Its section
 * fields describe the last section; earlier sections live on `section_break` nodes.
 */
export type DocumentCompat = {
  compatibilityMode?: number;
  htmlAutoSpacing?: boolean;
};

export type DocumentLayout = SectionLayout & {
  defaultTabStop?: number;
  /** From `word/settings.xml` compatibility settings. */
  compat?: DocumentCompat;
  /** Fraction of a space that justified lines may give up to fit another word. */
  justifyShrink?: number;
  /** Separate headers/footers on even pages (`w:evenAndOddHeaders`). */
  evenAndOddHeaders?: boolean;
  /** Style definitions, for the styles gallery. */
  styles?: StyleDefinition[];
  /** When true, word-layout engine may drive pagination (DOCX import). Studio layout sets false. */
  layoutEngine?: boolean;
};

/** Which page of a section a page is, for picking its header/footer. */
export type PageInfo = {
  /** 0-based index of the section the page starts in. */
  section: number;
  /** The page is the first page of its section. */
  first: boolean;
  /** 1-based displayed page number. */
  number: number;
};

/** CSS variable carrying `DocumentLayout.justifyShrink` to justified paragraphs. */
export const JUSTIFY_SHRINK_VAR = "--pm-justify-shrink";

/** Effective top/bottom margins of one page (Word grows them to fit the header/footer). */
export type PageFrame = { top: number; bottom: number };

const SERIF = /times|georgia|garamond|cambria|book|serif|palatino|century/i;
const MONO = /courier|consolas|mono|lucida console/i;

export function fontStack(family: string): string {
  const generic = MONO.test(family) ? "monospace" : SERIF.test(family) ? "serif" : "sans-serif";
  return `"${family.replace(/"/g, "")}", ${generic}`;
}

function round(n: number) {
  return Math.round(n * 1000) / 1000;
}

/** Advance width of U+0020 in em (hmtx / unitsPerEm). */
const SPACE_EM: [RegExp, number][] = [
  [/arial|helvetica|arimo|liberation sans/i, 0.2778],
  [/calibri|carlito/i, 0.2261],
  [/cambria|caladea/i, 0.2197],
  [/georgia/i, 0.2412],
  [/verdana/i, 0.3516],
  [/tahoma/i, 0.3125],
  [/segoe ui/i, 0.2744],
  [/courier|consolas|mono/i, 0.6],
];

function spaceEm(family: string | undefined): number {
  if (family) for (const [re, em] of SPACE_EM) if (re.test(family)) return em;
  return 0.25;
}

/** Section-level inputs for Word line box metrics (`w:spacing` + `w:docGrid`). */
export type LineMetricsContext = {
  docGridLinePitchTwips?: number;
};

export function lineMetricsContextFromFormat(
  format: ParagraphFormat | null | undefined,
): LineMetricsContext {
  const twips = format?.docGridSectionPitchTwips;
  return twips ? { docGridLinePitchTwips: twips } : {};
}

export function docGridPitchPxFromTwips(twips: number): number {
  return twips / TWIPS_PER_PX;
}

/** Single-line height in px (grid sections: font em box; otherwise Word's ~1.15×). */
export function singleLinePx(fontPt: number, context: LineMetricsContext = {}): number {
  const base = fontPt * PT_TO_PX;
  return context.docGridLinePitchTwips ? base : base * SINGLE_LINE_FACTOR;
}

/** One line's layout box height in px (CSS `line-height` and pagination). */
export function lineBoxPx(
  format: ParagraphFormat | null | undefined,
  context: LineMetricsContext = {},
): number | undefined {
  if (!format) return undefined;
  const fontPt = format.fontSize ?? 10;
  const single = singleLinePx(fontPt, context);
  const { lineRule, lineValue } = format;
  if (!lineRule || lineValue === undefined) return format.fontSize ? single : undefined;

  let box: number;
  if (lineRule === "auto") box = single * lineValue;
  else if (lineRule === "exact") box = lineValue;
  else if (lineRule === "atLeast") box = Math.max(single, lineValue);
  else return undefined;

  const pitchTwips = context.docGridLinePitchTwips;
  if (pitchTwips && lineRule === "auto") {
    const pitchPx = docGridPitchPxFromTwips(pitchTwips);
    const product = single * lineValue;
    if (Math.abs(product - pitchPx) < 0.51) box = pitchPx;
    // Word doc grid: single spacing still advances one linePitch row (e.g. 360 twips).
    else if (lineValue === 1 && box < pitchPx) box = pitchPx;
    // 1.5× (etc.) between one and two pitches: Word advances whole grid rows (e.g. 272 twips @ 11.5pt).
    else if (lineValue < 2 && product > pitchPx + 0.51) {
      const rows = Math.ceil(product / pitchPx);
      box = rows * pitchPx;
    }
  }
  return box;
}

export type ResolvedLineMetrics = {
  lineBoxPx: number | undefined;
  singleLinePx: number;
  /** Use pixel `line-height` in CSS (grid snap, exact, atLeast). */
  lineHeightPx: boolean;
  /** Top-align glyphs in the line box (grid or Word multiple > 1). */
  topAlign: boolean;
};

export function resolveLineMetrics(
  format: ParagraphFormat | null | undefined,
  context: LineMetricsContext = {},
): ResolvedLineMetrics {
  const fontPt = format?.fontSize ?? 10;
  const single = singleLinePx(fontPt, context);
  const box = lineBoxPx(format, context);
  const mult = format?.lineRule === "auto" ? (format.lineValue ?? 1) : 1;
  const grid = !!context.docGridLinePitchTwips;
  return {
    lineBoxPx: box,
    singleLinePx: single,
    lineHeightPx:
      format?.lineRule === "exact" ||
      format?.lineRule === "atLeast" ||
      (grid && format?.lineRule === "auto"),
    topAlign: grid || (format?.lineRule === "auto" && mult > 1),
  };
}

/** Small top inset so margin digits sit nearer the text baseline in a doc-grid row. */
export function gridLineNumberPaddingTop(lineBoxPx: number, fontSizePt = 12): number {
  const content = fontSizePt * PT_TO_PX;
  return Math.max(0, Math.min(lineBoxPx - 1, (lineBoxPx - content) * 0.35));
}

/** Slack below glyphs that may hang past the page bottom (pagination). */
export function lineBoxSlack(
  lineBox: number,
  glyphHeight: number,
  format: ParagraphFormat | null | undefined,
  context: LineMetricsContext = {},
): number {
  const mult = format?.lineRule === "auto" ? (format.lineValue ?? 1) : 1;
  if (context.docGridLinePitchTwips) {
    const pitchPx = docGridPitchPxFromTwips(context.docGridLinePitchTwips);
    if (Math.abs(lineBox - pitchPx) < 0.51) return 0;
    return Math.max(0, lineBox - glyphHeight);
  }
  if (format?.lineRule === "auto" && mult > 1) return lineBox - lineBox / mult;
  if (format?.lineRule === "exact" || format?.lineRule === "atLeast") {
    return Math.max(0, lineBox - glyphHeight);
  }
  return lineBox - lineBox / mult;
}

/** One line's box height in px from Word paragraph format (matches `hfParagraphHeight` / CSS intent). */
export function paragraphLineHeightPx(
  format: ParagraphFormat | null | undefined,
  context?: LineMetricsContext,
): number | undefined {
  const ctx = context ?? lineMetricsContextFromFormat(format);
  return lineBoxPx(format, ctx);
}

/** CSS line-height for a paragraph format, or undefined for the stylesheet default. */
export function lineHeightCss(
  format: ParagraphFormat,
  context?: LineMetricsContext,
): string | undefined {
  const ctx = context ?? lineMetricsContextFromFormat(format);
  const metrics = resolveLineMetrics(format, ctx);
  const { lineRule, lineValue } = format;
  if (!lineRule || lineValue === undefined) {
    if (!format.fontSize) return undefined;
    return metrics.lineHeightPx ? `${round(metrics.singleLinePx)}px` : String(SINGLE_LINE_FACTOR);
  }
  if (metrics.lineHeightPx && metrics.lineBoxPx !== undefined) {
    return `${round(metrics.lineBoxPx)}px`;
  }
  if (lineRule === "auto") return String(round(SINGLE_LINE_FACTOR * lineValue));
  if (lineRule === "exact") return `${round(lineValue)}px`;
  return `max(${SINGLE_LINE_FACTOR}em, ${round(lineValue)}px)`;
}

type CssDeclarations = Partial<
  Record<
    | "paddingLeft"
    | "paddingRight"
    | "textIndent"
    | "paddingTop"
    | "paddingBottom"
    | "fontFamily"
    | "fontSize"
    | "lineHeight"
    | "wordSpacing"
    | "borderTop"
    | "borderRight"
    | "borderBottom"
    | "borderLeft",
    string
  >
> & { textAlign?: TextAlign };

function borderCss(side: ParagraphBorderSide): string {
  return `${round(side.widthPt)}pt ${side.style} ${side.color}`;
}

/**
 * CSS for a formatted paragraph. Word spacing is additive, so it maps to padding, not margins.
 * Justified paragraphs get negative word spacing scaled by `JUSTIFY_SHRINK_VAR`: lines then
 * break as if spaces were narrower, and justification stretches them back, like Word's
 * space shrinking.
 */
export function paragraphStyle(format: ParagraphFormat | null | undefined): CssDeclarations {
  if (!format) return {};
  const css: CssDeclarations = {};
  if (format.align) css.textAlign = format.align;
  if (format.align === "justify") {
    css.wordSpacing = `calc(var(${JUSTIFY_SHRINK_VAR}, 0) * -${spaceEm(format.fontFamily)}em)`;
  }
  let padLeft = format.indentLeft ?? 0;
  let padRight = format.indentRight ?? 0;
  // Doc-grid body: vertical rhythm is the line pitch, not w:before/w:after padding.
  const onDocGrid = !!format.docGridSectionPitchTwips;
  let padTop = onDocGrid ? 0 : (format.spaceBefore ?? 0);
  let padBottom = onDocGrid ? 0 : (format.spaceAfter ?? 0);
  if (format.indentFirstLine) css.textIndent = `${round(format.indentFirstLine)}px`;
  if (format.fontFamily) css.fontFamily = fontStack(format.fontFamily);
  if (format.fontSize) css.fontSize = `${round(format.fontSize)}pt`;
  const lh = lineHeightCss(format);
  if (lh) css.lineHeight = lh;
  const box = format.boxBorder;
  if (box?.top) {
    css.borderTop = borderCss(box.top);
    padTop += box.top.paddingPx ?? 0;
  }
  if (box?.right) {
    css.borderRight = borderCss(box.right);
    padRight += box.right.paddingPx ?? 0;
  }
  if (box?.bottom) {
    css.borderBottom = borderCss(box.bottom);
    padBottom += box.bottom.paddingPx ?? 0;
  }
  if (box?.left) {
    css.borderLeft = borderCss(box.left);
    padLeft += box.left.paddingPx ?? 0;
  }
  if (padLeft) css.paddingLeft = `${round(padLeft)}px`;
  if (padRight) css.paddingRight = `${round(padRight)}px`;
  if (padTop) css.paddingTop = `${round(padTop)}px`;
  if (padBottom) css.paddingBottom = `${round(padBottom)}px`;
  return css;
}

export function runStyle(run: {
  fontFamily?: string | null;
  fontSize?: number | null;
  fontKern?: boolean | null;
}): CssDeclarations {
  const css: CssDeclarations = {};
  if (run.fontFamily) css.fontFamily = fontStack(run.fontFamily);
  if (run.fontSize) css.fontSize = `${round(run.fontSize)}pt`;
  return css;
}

function toCssText(css: CssDeclarations): string {
  return Object.entries(css)
    .map(([k, v]) => `${k.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)}: ${v}`)
    .join("; ");
}

export function paragraphCss(format: ParagraphFormat | null | undefined): string {
  return toCssText(paragraphStyle(format));
}

export function runCss(run: {
  fontFamily?: string | null;
  fontSize?: number | null;
  fontKern?: boolean | null;
}): string {
  const base = toCssText(runStyle(run));
  if (!run.fontKern) return base;
  return base ? `${base}; font-kerning: auto` : "font-kerning: auto";
}

function hfParagraphHasContent(p: HFParagraph): boolean {
  if (p.floating) return false;
  if (p.format.boxBorder) return true;
  return p.runs.some(
    (r) => r.kind === "field" || (r.kind === "text" && r.text.replace(/\s/g, "").length > 0),
  );
}

/** Height in px of one line of a header/footer paragraph (assumes it does not wrap). */
function hfParagraphHeight(p: HFParagraph): number {
  if (p.floating) return 0;
  const sizes = p.runs
    .map((r) => (r.kind === "tab" ? undefined : r.fontSize))
    .filter((s): s is number => !!s);
  const fontPt = Math.max(p.format.fontSize ?? 10, ...sizes);
  const ctx = lineMetricsContextFromFormat(p.format);
  const line = lineBoxPx(p.format, ctx) ?? singleLinePx(fontPt, ctx);
  return line + (p.format.spaceBefore ?? 0) + (p.format.spaceAfter ?? 0);
}

export function headerFooterHeight(
  blocks: HFBlock[] | undefined,
  _section?: SectionLayout | null,
): number {
  const paragraphs = hfSetParagraphs(blocks);
  let end = paragraphs.length;
  while (end > 0 && !hfParagraphHasContent(paragraphs[end - 1])) end -= 1;
  let sum = 0;
  for (let i = 0; i < end; i++) sum += hfParagraphHeight(paragraphs[i]);
  return sum;
}

/** Which header/footer of a section applies to a page. */
export function headerFooterFor(
  set: HeaderFooterSet | undefined,
  section: SectionLayout | null | undefined,
  page: Pick<PageInfo, "first" | "number">,
  evenAndOdd = false,
): HFBlock[] | undefined {
  if (!set) return undefined;
  if (page.first && section?.titlePage && set.first?.length) return set.first;
  if (evenAndOdd && page.number % 2 === 0 && set.even?.length) return set.even;
  return set.default;
}

/**
 * Layout of every section in order: the `section` attributes of the document's
 * `section_break` nodes, then the document layout (the last section).
 */
export function sectionLayouts(
  layout: DocumentLayout | null | undefined,
  breaks: (SectionLayout | null)[] = [],
): SectionLayout[] {
  return [...breaks.map((s) => s ?? {}), layout ?? {}];
}

/** Per-page effective margins: the body never overlaps the header or footer. */
export function pageFrames(
  setup: PageSetup,
  layout: DocumentLayout | null | undefined,
  sections: SectionLayout[] = sectionLayouts(layout),
): (page: PageInfo) => PageFrame {
  if (!layout) return () => ({ top: setup.marginTop, bottom: setup.marginBottom });
  const cache = new Map<string, PageFrame>();
  return (page) => {
    const section = sections[Math.min(page.section, sections.length - 1)] ?? layout;
    const parity = layout.evenAndOddHeaders ? page.number % 2 : 1;
    const key = `${page.section}:${page.first}:${parity}`;
    const cached = cache.get(key);
    if (cached) return cached;
    const marginTop = section.page?.marginTop ?? setup.marginTop;
    const marginBottom = section.page?.marginBottom ?? setup.marginBottom;
    const header = headerFooterFor(section.headers, section, page, layout.evenAndOddHeaders);
    const footer = headerFooterFor(section.footers, section, page, layout.evenAndOddHeaders);
    const headerDist = section.headerDistance ?? 48;
    const footerDist = section.footerDistance ?? 48;
    const frame = {
      top: header?.length
        ? Math.max(marginTop, headerDist + headerFooterHeight(header, section))
        : marginTop,
      bottom: footer?.length
        ? Math.max(marginBottom, footerDist + headerFooterHeight(footer, section))
        : marginBottom,
    };
    cache.set(key, frame);
    return frame;
  };
}
