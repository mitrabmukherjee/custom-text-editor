import {
  AlignmentType,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  ExternalHyperlink,
  Footer,
  FrameAnchorType,
  Header,
  HeadingLevel,
  HighlightColor,
  HorizontalPositionAlign,
  ImageRun,
  InsertedTextRun,
  InternalHyperlink,
  LevelFormat,
  LineNumberRestartFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  SectionType,
  ShadingType,
  Tab,
  TabStopType,
  Table,
  TableBorders,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlignTable,
  VerticalPositionAlign,
  WidthType,
  HeightRule,
  type ICommentOptions,
  type IFrameOptions,
  type IParagraphOptions,
  type IRunOptions,
  type ParagraphChild,
} from "docx";
import type {
  DocumentLayout,
  HFFrame,
  HFParagraph,
  ParagraphChange,
  ParagraphFormat,
  SectionLayout,
  SectionStart,
  StyleDefinition,
} from "@/lib/pm/format";
import type { CommentReply, CommentThread } from "@/lib/pm/comments";
import type { DocJSON, MarkJSON, NodeJSON } from "@/lib/pm/json";
import { exportCellBorders, exportParagraphBorders } from "./borderExport";
import type { CellBorderEdges, CellMarginPx } from "@/lib/pm/format";
import { TWIPS_PER_PX, type PageSetup } from "@/lib/pm/pagination/pageSetup";
import { BULLET_CHARS, type BulletStyle, type NumFormat } from "@/lib/pm/schema";
import { stylesOf } from "@/lib/pm/styles";
import { highlightName } from "@/lib/pm/textColors";
import { PLAIN_RUN, fullRunFormat, type RunFormat } from "@/lib/pm/track-changes/formatting";

export type ExportDocxOptions = {
  /** Turn on "Track Changes" in Word when the file is opened. */
  trackRevisions?: boolean;
};

const LIST_INDENT = 720;

const ROW_HEIGHT_RULE = {
  exact: HeightRule.EXACT,
  atLeast: HeightRule.ATLEAST,
  auto: HeightRule.AUTO,
} as const;

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

const ALIGNMENT = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.BOTH,
} as const;

const TAB_TYPE = {
  left: TabStopType.LEFT,
  center: TabStopType.CENTER,
  right: TabStopType.RIGHT,
  decimal: TabStopType.DECIMAL,
} as const;

type ChangeAttrs = { authorName?: string; date?: string };

/** Run font (from the paragraph mark or a textStyle mark). */
type RunFont = { fontFamily?: string | null; fontSize?: number | null };

function twips(px: number) {
  return Math.round(px * TWIPS_PER_PX);
}

/** Word dates are xsd:dateTime; drop milliseconds for maximum compatibility. */
function wordDate(date: string | undefined) {
  const d = date ? new Date(date) : new Date();
  const valid = Number.isNaN(d.getTime()) ? new Date() : d;
  return valid.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function findMark(marks: MarkJSON[] | undefined, type: string) {
  return marks?.find((m) => m.type === type);
}

function fontOptions(font: RunFont): Pick<IRunOptions, "font" | "size"> {
  return {
    font: font.fontFamily ?? undefined,
    size: font.fontSize ? Math.round(font.fontSize * 2) : undefined,
  };
}

function runFormatOfJSON(marks: MarkJSON[] | undefined): RunFormat {
  const attrs = (type: string) => findMark(marks, type)?.attrs;
  const style = attrs("textStyle") as RunFont | undefined;
  return {
    bold: !!findMark(marks, "strong"),
    italic: !!findMark(marks, "em"),
    fontFamily: style?.fontFamily ?? null,
    fontSize: style?.fontSize ?? null,
    underline: (attrs("underline")?.style as RunFormat["underline"]) ?? null,
    strike: (attrs("strike")?.style as RunFormat["strike"]) ?? null,
    color: (attrs("textColor")?.color as string | undefined) ?? null,
    highlight: (attrs("highlight")?.color as string | undefined) ?? null,
    caps: (attrs("caps")?.style as RunFormat["caps"]) ?? null,
    vertAlign: (attrs("vertAlign")?.align as RunFormat["vertAlign"]) ?? null,
  };
}

/** Word run properties for a run format; `font` is the paragraph mark's font. */
function runFormatOptions(format: RunFormat, font: RunFont) {
  const f = fullRunFormat(format);
  const highlight = f.highlight ? highlightName(f.highlight) : null;
  return {
    ...fontOptions({
      fontFamily: f.fontFamily ?? font.fontFamily,
      fontSize: f.fontSize ?? font.fontSize,
    }),
    bold: f.bold || undefined,
    italics: f.italic || undefined,
    underline: f.underline
      ? { type: f.underline as (typeof UnderlineType)[keyof typeof UnderlineType] }
      : undefined,
    strike: f.strike === "single" || undefined,
    doubleStrike: f.strike === "double" || undefined,
    color: f.color ? f.color.slice(1).toUpperCase() : undefined,
    highlight: highlight
      ? (highlight as (typeof HighlightColor)[keyof typeof HighlightColor])
      : undefined,
    shading:
      f.highlight && !highlight
        ? { type: ShadingType.CLEAR, color: "auto", fill: f.highlight.slice(1).toUpperCase() }
        : undefined,
    allCaps: f.caps === "all" || undefined,
    smallCaps: f.caps === "small" || undefined,
    superScript: f.vertAlign === "superscript" || undefined,
    subScript: f.vertAlign === "subscript" || undefined,
  } satisfies IRunOptions;
}

/**
 * Paragraph properties for a Word paragraph format. `styleIds`: styles written to
 * styles.xml, so the paragraph can reference its style (`w:pStyle`).
 */
function paragraphOptions(
  format: ParagraphFormat | null | undefined,
  styleIds?: ReadonlySet<string>,
  frame?: HFFrame | null,
): IParagraphOptions {
  if (!format) return frame ? { frame: frameOptions(frame) } : {};
  const first = format.indentFirstLine ?? 0;
  const lineRule = format.lineRule ?? "auto";
  return {
    style: format.styleId && styleIds?.has(format.styleId) ? format.styleId : undefined,
    alignment: format.align ? ALIGNMENT[format.align] : undefined,
    indent:
      format.indentLeft || format.indentRight || first
        ? {
            left: format.indentLeft ? twips(format.indentLeft) : undefined,
            right: format.indentRight ? twips(format.indentRight) : undefined,
            firstLine: first > 0 ? twips(first) : undefined,
            hanging: first < 0 ? twips(-first) : undefined,
          }
        : undefined,
    spacing: {
      before: twips(format.spaceBefore ?? 0),
      after: twips(format.spaceAfter ?? 0),
      line:
        lineRule === "auto"
          ? Math.round((format.lineValue ?? 1) * 240)
          : twips(format.lineValue ?? 0),
      lineRule:
        lineRule === "exact"
          ? LineRuleType.EXACT
          : lineRule === "atLeast"
            ? LineRuleType.AT_LEAST
            : LineRuleType.AUTO,
    },
    tabStops: format.tabs?.map((t) => ({ type: TAB_TYPE[t.align], position: twips(t.pos) })),
    widowControl: format.widowControl,
    keepNext: format.keepNext,
    keepLines: format.keepLines,
    suppressLineNumbers: format.suppressLineNumbers,
    pageBreakBefore: format.pageBreakBefore,
    contextualSpacing: format.contextualSpacing,
    run: fontOptions(format),
    border: exportParagraphBorders(format.boxBorder),
    ...(frame ? { frame: frameOptions(frame) } : {}),
  };
}

/** Paragraph properties plus a `w:pPrChange` for a tracked paragraph revision. */
function trackedParagraphOptions(
  node: NodeJSON,
  ids: { next: number },
  styleIds?: ReadonlySet<string>,
): IParagraphOptions {
  const options = paragraphOptions(
    node.attrs?.format as ParagraphFormat | null | undefined,
    styleIds,
    node.attrs?.frame as HFFrame | null | undefined,
  );
  const change = node.attrs?.paraChange as ParagraphChange | null | undefined;
  if (!change) return options;
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const { run, ...old } = paragraphOptions(change.old, styleIds);
  return {
    ...options,
    revision: {
      ...old,
      id: ids.next++,
      author: change.authorName || "Unknown",
      date: wordDate(change.date),
    },
  };
}

/** Word's Hyperlink style colour, for links without a colour of their own. */
const LINK_COLOR = "#0563c1";

function hyperlink(href: string, children: ParagraphChild[]): ParagraphChild {
  return href.startsWith("#")
    ? new InternalHyperlink({ anchor: href.slice(1), children })
    : new ExternalHyperlink({ link: href, children });
}

const IMAGE_EXPORT_TYPES: Record<string, "png" | "jpg" | "gif" | "bmp"> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
};

/** Bytes of a base64 data URL in a format Word can embed; null otherwise (e.g. SVG, remote URLs). */
function decodeDataUrl(src: string): { type: "png" | "jpg" | "gif" | "bmp"; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(src);
  const type = m ? IMAGE_EXPORT_TYPES[m[1].toLowerCase()] : undefined;
  if (!m || !type) return null;
  const binary = atob(m[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return { type, bytes };
}

/** Pixel size from a PNG, GIF or JPEG header. */
function intrinsicSize(b: Uint8Array): { width: number; height: number } | null {
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50) {
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49) {
    return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length; ) {
      if (b[i] !== 0xff) return null;
      const marker = b[i + 1];
      const length = view.getUint16(i + 2);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { width: view.getUint16(i + 7), height: view.getUint16(i + 5) };
      }
      i += 2 + length;
    }
  }
  return null;
}

function commentIdsOf(node: NodeJSON): string[] {
  return (node.marks ?? []).filter((m) => m.type === "comment").map((m) => String(m.attrs?.id));
}

/** The last inline node carrying each comment, where its range ends. */
function commentEnds(json: DocJSON): Map<NodeJSON, string[]> {
  const last = new Map<string, NodeJSON>();
  const visit = (node: NodeJSON) => {
    for (const id of commentIdsOf(node)) last.set(id, node);
    node.content?.forEach(visit);
  };
  json.content?.forEach(visit);
  const ends = new Map<NodeJSON, string[]>();
  for (const [id, node] of last) ends.set(node, [...(ends.get(node) ?? []), id]);
  return ends;
}

/** Word comment ids for each thread and its replies, in order of first appearance. */
class CommentNumbers {
  readonly order: string[] = [];
  private readonly ids = new Map<string, number[]>();
  private next = 0;

  constructor(private readonly threads: Record<string, CommentThread>) {}

  has(id: string) {
    return !!this.threads[id];
  }

  /** The thread's Word id followed by one per reply (replies share the anchor). */
  numbers(id: string): number[] {
    let ids = this.ids.get(id);
    if (!ids) {
      const count = 1 + (this.threads[id]?.replies?.length ?? 0);
      ids = Array.from({ length: count }, () => this.next++);
      this.ids.set(id, ids);
      this.order.push(id);
    }
    return ids;
  }

  options(): ICommentOptions[] {
    const paragraphs = (text: string) =>
      text.split("\n").map((line) => new Paragraph({ children: [new TextRun(line)] }));
    const initials = (name: string) =>
      name
        .split(/\s+/)
        .filter(Boolean)
        .map((w) => w[0]!.toUpperCase())
        .join("")
        .slice(0, 3);
    const entry = (c: CommentReply, id: number) => ({
      id,
      author: c.authorName || "Unknown",
      initials: initials(c.authorName || "U"),
      date: c.date ? new Date(c.date) : undefined,
      children: paragraphs(c.text),
    });
    return this.order.flatMap((threadId) => {
      const thread = this.threads[threadId];
      const [id, ...replyIds] = this.ids.get(threadId)!;
      return [
        { ...entry(thread, id), resolved: thread.resolved || undefined },
        ...(thread.replies ?? []).map((reply, i) => ({ ...entry(reply, replyIds[i]), parentId: id })),
      ];
    });
  }
}

class RunBuilder {
  /** Revision ids, shared with paragraph revisions. */
  readonly ids = { next: 1 };
  /** Comment ranges started but not yet ended. */
  private readonly openComments = new Set<string>();

  constructor(
    private readonly comments: CommentNumbers,
    private readonly ends: Map<NodeJSON, string[]>,
  ) {}

  inline(nodes: NodeJSON[] | undefined, base: RunFont = {}): ParagraphChild[] {
    const out: ParagraphChild[] = [];
    /** Runs of the hyperlink being collected; consecutive nodes with one href share it. */
    let link: { href: string; children: ParagraphChild[] } | null = null;
    for (const node of nodes ?? []) {
      const href = (findMark(node.marks, "link")?.attrs?.href as string | undefined) ?? null;
      const run = this.inlineRun(node, base, !!href);
      if (!run) continue;
      if (link && link.href !== href) {
        out.push(hyperlink(link.href, link.children));
        link = null;
      }
      if (href && !link) link = { href, children: [] };
      const target = link ? link.children : out;
      this.startComments(node, target);
      target.push(run);
      this.endComments(node, target);
    }
    if (link) out.push(hyperlink(link.href, link.children));
    return out;
  }

  private inlineRun(node: NodeJSON, base: RunFont, linked: boolean): ParagraphChild | null {
    if (node.type === "text" && node.text) return this.run({ text: node.text }, node.marks, base, linked);
    if (node.type === "hard_break") return this.run({ break: 1 }, node.marks, base);
    if (node.type === "tab") return this.run({ children: [new Tab()] }, node.marks, base);
    if (node.type === "image") return this.image(node);
    if (node.type === "timestamp") {
      const time = String((node.attrs as { time?: string })?.time ?? "");
      return time ? this.run({ text: time }, node.marks, base, linked) : null;
    }
    return null;
  }

  private image(node: NodeJSON): ParagraphChild | null {
    const { src, alt, width, height } = (node.attrs ?? {}) as Record<string, unknown>;
    const data = typeof src === "string" ? decodeDataUrl(src) : null;
    if (!data) return null;
    const size =
      typeof width === "number" && typeof height === "number"
        ? { width, height }
        : (intrinsicSize(data.bytes) ?? { width: 200, height: 150 });
    const insertion = findMark(node.marks, "insertion");
    const deletion = findMark(node.marks, "deletion");
    return new ImageRun({
      type: data.type,
      data: data.bytes,
      transformation: size,
      altText: alt ? { name: String(alt), description: String(alt) } : undefined,
      insertion: insertion ? this.revision(insertion) : undefined,
      deletion: deletion ? this.revision(deletion) : undefined,
    });
  }

  private startComments(node: NodeJSON, out: ParagraphChild[]) {
    for (const id of commentIdsOf(node)) {
      if (this.openComments.has(id) || !this.comments.has(id)) continue;
      this.openComments.add(id);
      for (const n of this.comments.numbers(id)) out.push(new CommentRangeStart(n));
    }
  }

  private endComments(node: NodeJSON, out: ParagraphChild[]) {
    for (const id of this.ends.get(node) ?? []) {
      if (!this.openComments.delete(id)) continue;
      for (const n of this.comments.numbers(id)) {
        out.push(new CommentRangeEnd(n), new TextRun({ children: [new CommentReference(n)] }));
      }
    }
  }

  private run(
    base: IRunOptions,
    marks: MarkJSON[] | undefined,
    font: RunFont,
    linked = false,
  ): ParagraphChild {
    const formatChange = findMark(marks, "formatChange");
    const old = formatChange?.attrs?.old as RunFormat | null | undefined;
    const format = runFormatOfJSON(marks);
    if (linked) {
      format.color ??= LINK_COLOR;
      format.underline ??= "single";
    }
    const options: IRunOptions = {
      ...base,
      ...runFormatOptions(format, font),
      ...(formatChange && old
        ? { revision: { ...this.revision(formatChange), ...runFormatOptions(old, font) } }
        : {}),
    };
    // Text inserted by one reviewer and deleted by another is exported as deleted.
    const deletion = findMark(marks, "deletion");
    if (deletion) {
      return new DeletedTextRun({ ...options, ...this.revision(deletion) });
    }
    const insertion = findMark(marks, "insertion");
    if (insertion) {
      return new InsertedTextRun({ ...options, ...this.revision(insertion) });
    }
    return new TextRun(options);
  }

  private revision(mark: MarkJSON) {
    const attrs = (mark.attrs ?? {}) as ChangeAttrs;
    return {
      id: this.ids.next++,
      author: attrs.authorName || "Unknown",
      date: wordDate(attrs.date),
    };
  }
}

/** Body content of one Word section, and how the section starts. */
type SectionChunk = {
  start: SectionStart;
  layout: SectionLayout | null;
  paragraphs: (Paragraph | Table)[];
};

class BlockBuilder {
  /** Closed sections; the open one collects into `paragraphs`. */
  readonly sections: SectionChunk[] = [];
  paragraphs: (Paragraph | Table)[] = [];
  /** How the open section starts. */
  start: SectionStart = "nextPage";
  readonly numbering = new ListNumbering();

  constructor(
    private readonly runs: RunBuilder,
    private readonly styleIds: ReadonlySet<string>,
  ) {}

  blocks(nodes: NodeJSON[] | undefined) {
    for (const node of nodes ?? []) this.block(node);
  }

  block(node: NodeJSON) {
    const format = node.attrs?.format as ParagraphFormat | null | undefined;
    switch (node.type) {
      case "paragraph":
      case "qa_line":
      case "speaker_line": {
        const prefix =
          node.type === "qa_line"
            ? node.attrs?.role === "a"
              ? "A. "
              : "Q. "
            : node.type === "speaker_line"
              ? `${(node.attrs?.label as string) ?? ""} `
              : "";
        const children = prefix
          ? [new TextRun(prefix), ...this.runs.inline(node.content, format ?? {})]
          : this.runs.inline(node.content, format ?? {});
        this.paragraphs.push(
          new Paragraph({
            ...trackedParagraphOptions(node, this.runs.ids, this.styleIds),
            children,
          }),
        );
        break;
      }
      case "heading": {
        const level = Math.min(Math.max(Number(node.attrs?.level ?? 1), 1), 6);
        this.paragraphs.push(
          new Paragraph({
            ...trackedParagraphOptions(node, this.runs.ids),
            heading: HEADINGS[level - 1],
            children: this.runs.inline(node.content, format ?? {}),
          }),
        );
        break;
      }
      case "page_break":
        this.paragraphs.push(new Paragraph({ children: [new PageBreak()] }));
        break;
      case "section_break":
        if (!this.paragraphs.length) this.paragraphs.push(new Paragraph({}));
        this.sections.push({
          start: this.start,
          layout: (node.attrs?.section as SectionLayout | null) ?? null,
          paragraphs: this.paragraphs,
        });
        this.paragraphs = [];
        this.start = (node.attrs?.start as SectionStart) ?? "nextPage";
        break;
      case "bullet_list":
      case "ordered_list":
        this.list(node, 0, this.numbering.forList(node));
        break;
      case "table":
        this.paragraphs.push(this.table(node));
        break;
      default:
        this.blocks(node.content);
    }
  }

  private table(node: NodeJSON): Table {
    const rows = (node.content ?? []).map((row) => {
      const rAttrs = (row.attrs ?? {}) as Record<string, unknown>;
      const rowHeight = rAttrs.rowHeight as number | undefined;
      const rowHeightRule = rAttrs.rowHeightRule as keyof typeof ROW_HEIGHT_RULE | undefined;
      return new TableRow({
        children: (row.content ?? []).map((cell) => this.cell(cell)),
        cantSplit: rAttrs.cantSplit ? true : undefined,
        height:
          rowHeight && rowHeight > 0
            ? {
                value: twips(rowHeight),
                rule: (rowHeightRule && ROW_HEIGHT_RULE[rowHeightRule]) ?? HeightRule.ATLEAST,
              }
            : undefined,
      });
    });
    const firstRow = node.content?.[0]?.content ?? [];
    const widths = firstRow.flatMap((cell) => {
      const colwidth = cell.attrs?.colwidth as number[] | null | undefined;
      const span = Number(cell.attrs?.colspan ?? 1);
      return colwidth?.length === span ? colwidth : Array<number>(span).fill(0);
    });
    const known = widths.length > 0 && widths.every((w) => w > 0);
    const tAttrs = (node.attrs ?? {}) as Record<string, unknown>;
    const tableWidthPx = tAttrs.width as number | undefined;
    const tableWidthPct = tAttrs.widthPct as number | undefined;
    const sumPx = widths.reduce((a, b) => a + b, 0);
    const width =
      tableWidthPx !== undefined && tableWidthPx > 0
        ? { size: Math.round(tableWidthPx * TWIPS_PER_PX), type: WidthType.DXA }
        : tableWidthPct !== undefined && tableWidthPct > 0
          ? { size: Math.round(tableWidthPct * 50), type: WidthType.PERCENTAGE }
          : known
            ? { size: Math.round(sumPx * TWIPS_PER_PX), type: WidthType.DXA }
            : undefined;
    const align = tAttrs.align as string | undefined;
    const indentPx = tAttrs.indent as number | undefined;
    return new Table({
      rows,
      columnWidths: known ? widths.map((w) => Math.round(w * TWIPS_PER_PX)) : undefined,
      width,
      borders: node.attrs?.borders === false ? TableBorders.NONE : undefined,
      alignment:
        align === "center"
          ? AlignmentType.CENTER
          : align === "right"
            ? AlignmentType.RIGHT
            : undefined,
      indent: indentPx ? { size: twips(indentPx), type: WidthType.DXA } : undefined,
    });
  }

  private cell(cell: NodeJSON): TableCell {
    const outer = this.paragraphs;
    this.paragraphs = [];
    this.blocks(cell.content);
    const children = this.paragraphs.length ? this.paragraphs : [new Paragraph({})];
    this.paragraphs = outer;
    const attrs = (cell.attrs ?? {}) as Record<string, unknown>;
    const colwidth = attrs.colwidth as number[] | null | undefined;
    const background = typeof attrs.background === "string" ? attrs.background : null;
    const cellMargin = attrs.cellMargin as CellMarginPx | null | undefined;
    const edges = attrs.edges as CellBorderEdges | null | undefined;
    return new TableCell({
      children,
      columnSpan: Number(attrs.colspan ?? 1) > 1 ? Number(attrs.colspan) : undefined,
      rowSpan: Number(attrs.rowspan ?? 1) > 1 ? Number(attrs.rowspan) : undefined,
      width: colwidth?.every((w) => w > 0)
        ? { size: Math.round(colwidth.reduce((a, b) => a + b, 0) * TWIPS_PER_PX), type: WidthType.DXA }
        : undefined,
      shading: background
        ? { type: ShadingType.CLEAR, color: "auto", fill: background.slice(1).toUpperCase() }
        : undefined,
      verticalAlign:
        attrs.valign === "middle"
          ? VerticalAlignTable.CENTER
          : attrs.valign === "bottom"
            ? VerticalAlignTable.BOTTOM
            : undefined,
      margins: cellMargin
        ? {
            top: cellMargin.top ? twips(cellMargin.top) : undefined,
            right: cellMargin.right ? twips(cellMargin.right) : undefined,
            bottom: cellMargin.bottom ? twips(cellMargin.bottom) : undefined,
            left: cellMargin.left ? twips(cellMargin.left) : undefined,
          }
        : undefined,
      borders: exportCellBorders(edges),
    });
  }

  private list(node: NodeJSON, level: number, numbering: { reference: string; instance: number }) {
    for (const item of node.content ?? []) {
      (item.content ?? []).forEach((child, i) => {
        if (child.type === "bullet_list" || child.type === "ordered_list") {
          this.list(child, Math.min(level + 1, 8), numbering);
          return;
        }
        if (child.type !== "paragraph" && child.type !== "heading") {
          this.block(child);
          return;
        }
        const format = child.attrs?.format as ParagraphFormat | null | undefined;
        const props = trackedParagraphOptions(child, this.runs.ids, this.styleIds);
        const children = this.runs.inline(child.content, format ?? {});
        if (i > 0) {
          this.paragraphs.push(
            new Paragraph({ ...props, children, indent: { left: LIST_INDENT * (level + 1) } }),
          );
        } else {
          this.paragraphs.push(
            new Paragraph({ ...props, children, numbering: { ...numbering, level } }),
          );
        }
      });
    }
  }
}

type ListLevelDef = { kind: "bullet" | "ordered"; format: NumFormat | null; bullet: BulletStyle | null; start: number };

const DEFAULT_ORDERED: NumFormat[] = ["decimal", "lowerLetter", "lowerRoman"];
const DEFAULT_BULLETS: BulletStyle[] = ["disc", "circle", "square"];

const LEVEL_FORMAT: Record<Exclude<NumFormat, "legal">, (typeof LevelFormat)[keyof typeof LevelFormat]> = {
  decimal: LevelFormat.DECIMAL,
  lowerLetter: LevelFormat.LOWER_LETTER,
  upperLetter: LevelFormat.UPPER_LETTER,
  lowerRoman: LevelFormat.LOWER_ROMAN,
  upperRoman: LevelFormat.UPPER_ROMAN,
};

/** Word abstract numbering definitions, one per distinct set of list levels. */
class ListNumbering {
  private readonly refs = new Map<string, { reference: string; levels: ListLevelDef[] }>();
  private instance = 0;

  /** Numbering reference for a top-level list (each list restarts at its own start). */
  forList(node: NodeJSON) {
    const levels = listLevels(node);
    const key = JSON.stringify(levels);
    let ref = this.refs.get(key);
    if (!ref) {
      ref = { reference: `list-${this.refs.size + 1}`, levels };
      this.refs.set(key, ref);
    }
    return { reference: ref.reference, instance: this.instance++ };
  }

  config() {
    return [...this.refs.values()].map(({ reference, levels }) => ({
      reference,
      levels: levels.map((l, level) => {
        const indent = { left: LIST_INDENT * (level + 1), hanging: 360 };
        if (l.kind === "bullet") {
          return {
            level,
            format: LevelFormat.BULLET,
            text: BULLET_CHARS[l.bullet ?? "disc"],
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent } },
          };
        }
        const legal = l.format === "legal";
        return {
          level,
          format: l.format && l.format !== "legal" ? LEVEL_FORMAT[l.format] : LevelFormat.DECIMAL,
          text: legal
            ? `${Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join(".")}.`
            : `%${level + 1}.`,
          start: l.start,
          isLegalNumberingStyle: legal || undefined,
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent } },
        };
      }),
    }));
  }
}

/** The nine Word levels of a list: each nesting depth's first list, else Word's defaults. */
function listLevels(root: NodeJSON): ListLevelDef[] {
  const found: (ListLevelDef | undefined)[] = [];
  const visit = (node: NodeJSON, depth: number) => {
    if (depth > 8) return;
    if (!found[depth]) {
      const ordered = node.type === "ordered_list";
      found[depth] = {
        kind: ordered ? "ordered" : "bullet",
        format: ordered ? ((node.attrs?.numFormat as NumFormat | null) ?? null) : null,
        bullet: ordered ? null : ((node.attrs?.bullet as BulletStyle | null) ?? null),
        start: ordered ? Number(node.attrs?.order ?? 1) : 1,
      };
    }
    for (const item of node.content ?? []) {
      for (const child of item.content ?? []) {
        if (child.type === "bullet_list" || child.type === "ordered_list") visit(child, depth + 1);
      }
    }
  };
  visit(root, 0);
  const legal = found[0]?.format === "legal";
  let kind = found[0]!.kind;
  return Array.from({ length: 9 }, (_, depth): ListLevelDef => {
    const own = found[depth];
    kind = own?.kind ?? kind;
    if (kind === "bullet") {
      return { kind, format: null, bullet: own?.bullet ?? DEFAULT_BULLETS[depth % 3], start: 1 };
    }
    const format = own?.format ?? (legal ? "legal" : DEFAULT_ORDERED[depth % 3]);
    return { kind, format, bullet: null, start: own?.start ?? 1 };
  });
}

const FRAME_ANCHOR = {
  page: FrameAnchorType.PAGE,
  margin: FrameAnchorType.MARGIN,
  text: FrameAnchorType.TEXT,
} as const;

/** `w:framePr` for a positioned header/footer paragraph. */
function frameOptions(frame: HFFrame | undefined): IFrameOptions | undefined {
  if (!frame) return undefined;
  const base = {
    anchor: { horizontal: FRAME_ANCHOR[frame.hAnchor], vertical: FRAME_ANCHOR[frame.vAnchor] },
    width: frame.width !== undefined ? twips(frame.width) : undefined,
    height: frame.height !== undefined ? twips(frame.height) : undefined,
  };
  if (frame.xAlign && frame.yAlign) {
    return {
      ...base,
      type: "alignment",
      alignment: {
        x: frame.xAlign as (typeof HorizontalPositionAlign)[keyof typeof HorizontalPositionAlign],
        y: (frame.yAlign === "inline" ? "top" : frame.yAlign) as (typeof VerticalPositionAlign)[keyof typeof VerticalPositionAlign],
      },
    };
  }
  return {
    ...base,
    type: "absolute",
    position: { x: twips(frame.x ?? 0), y: twips(frame.y ?? 0) },
  };
}

function hfParagraphs(paragraphs: HFParagraph[]): Paragraph[] {
  return paragraphs
    .filter((p) => !p.floating || p.frame)
    .map(
      (p) =>
        new Paragraph({
          ...paragraphOptions(p.format),
          frame: frameOptions(p.frame),
          children: p.runs.map((run) => {
            if (run.kind === "tab") return new TextRun({ children: [new Tab()] });
            const options: IRunOptions = {
              ...fontOptions({
                fontFamily: run.fontFamily ?? p.format.fontFamily,
                fontSize: run.fontSize ?? p.format.fontSize,
              }),
              bold: run.bold,
              italics: run.italic,
              underline: run.underline ? {} : undefined,
              color: run.color ? run.color.slice(1).toUpperCase() : undefined,
            };
            if (run.kind === "field") {
              return new TextRun({
                ...options,
                children: [run.field === "PAGE" ? PageNumber.CURRENT : PageNumber.TOTAL_PAGES],
              });
            }
            return new TextRun({ ...options, text: run.text });
          }),
        }),
    );
}

function headerFooterOptions(section: SectionLayout | null) {
  if (!section) {
    return {
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { line: 240 },
              children: [new TextRun({ children: ["Page ", PageNumber.CURRENT] })],
            }),
          ],
        }),
      },
    };
  }
  const headers: { default?: Header; first?: Header; even?: Header } = {};
  const footers: { default?: Footer; first?: Footer; even?: Footer } = {};
  for (const key of ["default", "first", "even"] as const) {
    const h = section.headers?.[key];
    if (h) headers[key] = new Header({ children: hfParagraphs(h) });
    const f = section.footers?.[key];
    if (f) footers[key] = new Footer({ children: hfParagraphs(f) });
  }
  return { headers, footers };
}

const LINE_RESTART = {
  newPage: LineNumberRestartFormat.NEW_PAGE,
  newSection: LineNumberRestartFormat.NEW_SECTION,
  continuous: LineNumberRestartFormat.CONTINUOUS,
} as const;

const SECTION_TYPE = {
  nextPage: SectionType.NEXT_PAGE,
  continuous: SectionType.CONTINUOUS,
  evenPage: SectionType.EVEN_PAGE,
  oddPage: SectionType.ODD_PAGE,
} as const;

/** `w:sectPr` for one section; `setup` fills in a missing page size. */
function sectionProperties(section: SectionLayout | null, setup: PageSetup, start: SectionStart) {
  const page = section?.page ?? setup;
  const ln = section?.lineNumbers;
  return {
    type: SECTION_TYPE[start],
    page: {
      size: { width: twips(page.width), height: twips(page.height) },
      margin: {
        top: twips(page.marginTop),
        right: twips(page.marginRight),
        bottom: twips(page.marginBottom),
        left: twips(page.marginLeft),
        header: twips(section?.headerDistance ?? page.marginTop / 2),
        footer: twips(section?.footerDistance ?? page.marginBottom / 2),
      },
      pageNumbers:
        section?.pageNumberStart !== undefined ? { start: section.pageNumberStart } : undefined,
    },
    titlePage: section?.titlePage,
    column: section?.columns
      ? { count: section.columns.count, space: twips(section.columns.space) }
      : undefined,
    lineNumbers: ln
      ? {
          countBy: ln.countBy,
          start: Math.max(0, ln.start - 1),
          restart: LINE_RESTART[ln.restart],
          distance: twips(ln.distance),
        }
      : undefined,
  };
}

const BUILT_IN_DEFAULTS: Record<string, "title" | `heading${1 | 2 | 3 | 4 | 5 | 6}`> = {
  Title: "title",
  Heading1: "heading1",
  Heading2: "heading2",
  Heading3: "heading3",
  Heading4: "heading4",
  Heading5: "heading5",
  Heading6: "heading6",
};

function styleProperties(style: StyleDefinition) {
  const format = style.format ?? {};
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const { run: _font, style: _style, ...paragraph } = paragraphOptions(format);
  const run = runFormatOptions({ ...PLAIN_RUN, ...(style.run as Partial<RunFormat> | undefined) }, format);
  return { paragraph, run };
}

/**
 * styles.xml: Word's built-in heading and title styles redefined from the document
 * (so they don't come out in docx's default blue), plus every other paragraph style a
 * paragraph uses. Paragraphs keep their full formatting as direct formatting too.
 */
function exportStyles(json: DocJSON) {
  const all = stylesOf(json.attrs);
  const used = new Set<string>();
  const visit = (nodes: NodeJSON[] | undefined) => {
    for (const node of nodes ?? []) {
      const id = (node.attrs?.format as ParagraphFormat | null | undefined)?.styleId;
      if (id) used.add(id);
      visit(node.content);
    }
  };
  visit(json.content);
  const defaults: Partial<Record<(typeof BUILT_IN_DEFAULTS)[string], ReturnType<typeof styleProperties>>> = {};
  const paragraphStyles = [];
  const ids = new Set<string>(["Normal", ...Object.keys(BUILT_IN_DEFAULTS)]);
  for (const style of all) {
    if (style.type !== "paragraph") continue;
    const key = BUILT_IN_DEFAULTS[style.id];
    if (key) {
      defaults[key] = styleProperties(style);
    } else if (style.id !== "Normal" && used.has(style.id)) {
      ids.add(style.id);
      paragraphStyles.push({
        id: style.id,
        name: style.name,
        basedOn: "Normal",
        next: style.next,
        quickFormat: style.gallery,
        ...styleProperties(style),
      });
    }
  }
  return { defaults, paragraphStyles, ids };
}

/** Serialize a ProseMirror document to DOCX, with tracked changes as Word revisions. */
export async function exportDocx(
  json: DocJSON,
  setup: PageSetup,
  options: ExportDocxOptions = {},
): Promise<ArrayBuffer> {
  const layout = (json.attrs?.layout as DocumentLayout | null | undefined) ?? null;
  const styles = exportStyles(json);
  const threads = (json.attrs?.comments as Record<string, CommentThread> | null | undefined) ?? {};
  const comments = new CommentNumbers(threads);
  const builder = new BlockBuilder(new RunBuilder(comments, commentEnds(json)), styles.ids);
  builder.blocks(json.content);
  if (!builder.paragraphs.length) builder.paragraphs.push(new Paragraph({}));
  const chunks: SectionChunk[] = [
    ...builder.sections,
    { start: builder.start, layout: layout ? { ...layout, page: setup } : null, paragraphs: builder.paragraphs },
  ];

  const doc = new Document({
    creator: "Transcript editor",
    features: { trackRevisions: options.trackRevisions ?? false },
    evenAndOddHeaderAndFooters: layout?.evenAndOddHeaders,
    styles: {
      default: {
        document: layout
          ? { run: { font: "Times New Roman", size: 20 }, paragraph: { spacing: { line: 240 } } }
          : { run: { font: "Times New Roman", size: 24 }, paragraph: { spacing: { line: 480 } } },
        ...styles.defaults,
      },
      paragraphStyles: styles.paragraphStyles,
    },
    numbering: { config: builder.numbering.config() },
    comments: { children: comments.options() },
    sections: chunks.map((chunk, i) => ({
      properties: sectionProperties(
        chunk.layout,
        setup,
        i + 1 < chunks.length ? chunks[i + 1].start : "nextPage",
      ),
      ...headerFooterOptions(layout ? chunk.layout : null),
      children: chunk.paragraphs,
    })),
  });

  return Packer.toArrayBuffer(doc);
}
