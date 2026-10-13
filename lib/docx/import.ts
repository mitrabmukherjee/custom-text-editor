import { DOMParser } from "@xmldom/xmldom";
import JSZip from "jszip";
import type {
  DocumentLayout,
  HFFrame,
  HFParagraph,
  HFRun,
  HeaderFooterSet,
  LineNumbering,
  ParagraphChange,
  ParagraphFormat,
  SectionLayout,
  SectionStart,
  StyleDefinition,
  TabStop,
  TextAlign,
} from "@/lib/pm/format";
import type { CommentReply, CommentThread } from "@/lib/pm/comments";
import type { DocJSON, MarkJSON, NodeJSON } from "@/lib/pm/json";
import {
  LETTER_PAGE_SETUP,
  TWIPS_PER_PX,
  type PageSetup,
} from "@/lib/pm/pagination/pageSetup";
import { schema, type BulletStyle, type NumFormat } from "@/lib/pm/schema";
import { HIGHLIGHT_COLORS, normalizeHex } from "@/lib/pm/textColors";
import {
  describeFormatChange,
  runFormatMarks,
  type RunFormat,
  type UnderlineStyle,
} from "@/lib/pm/track-changes/formatting";
import { describeParagraphChange } from "@/lib/pm/track-changes/paragraphFormat";

export type DocxImportResult = {
  doc: DocJSON;
  /** Page size and margins from the last section, when present. */
  pageSetup?: PageSetup;
  /** Full Word layout (also stored on the doc node's `layout` attribute). */
  layout?: DocumentLayout;
};

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Word's size when no style sets one: 10pt. */
const DEFAULT_FONT_SIZE_PT = 10;
const DEFAULT_FONT = "Times New Roman";
/** "Auto" paragraph spacing in HTML mode. */
const AUTO_SPACING_PX = (14 * 96) / 72;

type XElement = {
  nodeType: number;
  localName: string | null;
  textContent: string | null;
  childNodes: ArrayLike<XElement>;
  getAttributeNS(ns: string, name: string): string | null;
  getAttribute(name: string): string | null;
  getElementsByTagNameNS(ns: string, name: string): ArrayLike<XElement>;
};

type ListKind = "bullet" | "ordered";
type ListLevel = {
  kind: ListKind;
  format: NumFormat | null;
  bullet: BulletStyle | null;
  start: number;
};
type ListInfo = ListLevel & {
  level: number;
  numId: string;
  /** Numbering continues across lists sharing an abstract definition. */
  abstractId: string;
  startOverride?: number;
};
type Block = { node: NodeJSON; list?: ListInfo };

function elements(el: XElement): XElement[] {
  const out: XElement[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i];
    if (n.nodeType === 1) out.push(n);
  }
  return out;
}

function child(el: XElement | undefined, name: string): XElement | undefined {
  return el ? elements(el).find((c) => c.localName === name) : undefined;
}

function attr(el: XElement | undefined, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(W_NS, name) || el.getAttribute(`w:${name}`);
}

function num(el: XElement | undefined, name: string): number | undefined {
  const raw = attr(el, name);
  if (raw === null || raw === "") return undefined;
  const v = Number(raw);
  return Number.isFinite(v) ? v : undefined;
}

function onOff(val: string | null): boolean | undefined {
  if (val === null) return undefined;
  return !["0", "false", "off"].includes(val.toLowerCase());
}

function toggle(el: XElement | undefined): boolean | undefined {
  if (!el) return undefined;
  return onOff(attr(el, "val")) ?? true;
}

function parseXml(xml: string): XElement {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  return doc.documentElement as unknown as XElement;
}

// —— Styles ——

/** Paragraph properties as written in OOXML (twips), before conversion. */
type RawPara = {
  jc?: string;
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
  before?: number;
  after?: number;
  beforeAuto?: boolean;
  afterAuto?: boolean;
  line?: number;
  lineRule?: string;
  tabs?: { pos: number; val: string }[];
  widowControl?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  suppressLineNumbers?: boolean;
  contextualSpacing?: boolean;
};

type RawRun = {
  font?: string;
  /** Half-points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  vanish?: boolean;
  /** `w:u` value; "none" cancels an inherited underline. */
  underline?: string;
  strike?: boolean;
  dstrike?: boolean;
  /** "#rrggbb" or "auto". */
  color?: string;
  highlight?: string;
  shading?: string;
  caps?: boolean;
  smallCaps?: boolean;
  vertAlign?: string;
};

type StyleDef = {
  type: string;
  basedOn?: string;
  name: string;
  next?: string;
  /** Shown in Word's styles gallery (`w:qFormat`). */
  gallery: boolean;
  pPr: RawPara;
  rPr: RawRun;
};

type ThemeFonts = { major?: string; minor?: string };

function defined<T extends object>(obj: T): Partial<T> {
  return Object.fromEntries(
    Object.entries(obj).filter(([, v]) => v !== undefined),
  ) as Partial<T>;
}

function parsePPr(pPr: XElement | undefined): RawPara {
  if (!pPr) return {};
  const ind = child(pPr, "ind");
  const spacing = child(pPr, "spacing");
  const tabsEl = child(pPr, "tabs");
  return defined({
    jc: attr(child(pPr, "jc"), "val") ?? undefined,
    left: num(ind, "left") ?? num(ind, "start"),
    right: num(ind, "right") ?? num(ind, "end"),
    firstLine: num(ind, "firstLine"),
    hanging: num(ind, "hanging"),
    before: num(spacing, "before"),
    after: num(spacing, "after"),
    beforeAuto: onOff(attr(spacing, "beforeAutospacing")),
    afterAuto: onOff(attr(spacing, "afterAutospacing")),
    line: num(spacing, "line"),
    lineRule: attr(spacing, "lineRule") ?? undefined,
    tabs: tabsEl
      ? elements(tabsEl)
          .filter((t) => t.localName === "tab")
          .map((t) => ({ pos: num(t, "pos") ?? 0, val: attr(t, "val") ?? "left" }))
      : undefined,
    widowControl: toggle(child(pPr, "widowControl")),
    keepNext: toggle(child(pPr, "keepNext")),
    keepLines: toggle(child(pPr, "keepLines")),
    pageBreakBefore: toggle(child(pPr, "pageBreakBefore")),
    suppressLineNumbers: toggle(child(pPr, "suppressLineNumbers")),
    contextualSpacing: toggle(child(pPr, "contextualSpacing")),
  });
}

function parseRPr(rPr: XElement | undefined, theme: ThemeFonts): RawRun {
  if (!rPr) return {};
  const fonts = child(rPr, "rFonts");
  const themeName = attr(fonts, "asciiTheme") ?? attr(fonts, "hAnsiTheme");
  const themeFont = themeName
    ? themeName.startsWith("major")
      ? theme.major
      : theme.minor
    : undefined;
  const sz = num(child(rPr, "sz"), "val");
  const color = attr(child(rPr, "color"), "val");
  const highlight = attr(child(rPr, "highlight"), "val");
  const fill = attr(child(rPr, "shd"), "fill");
  return defined({
    font: attr(fonts, "ascii") ?? attr(fonts, "hAnsi") ?? themeFont ?? undefined,
    size: sz,
    bold: toggle(child(rPr, "b")),
    italic: toggle(child(rPr, "i")),
    vanish: toggle(child(rPr, "vanish")),
    underline: child(rPr, "u") ? (attr(child(rPr, "u"), "val") ?? "single") : undefined,
    strike: toggle(child(rPr, "strike")),
    dstrike: toggle(child(rPr, "dstrike")),
    color: color ? (color === "auto" ? "auto" : (normalizeHex(color) ?? undefined)) : undefined,
    highlight: highlight
      ? highlight === "none"
        ? "none"
        : (HIGHLIGHT_COLORS[highlight] ?? undefined)
      : undefined,
    shading: fill ? (fill === "auto" ? "none" : (normalizeHex(fill) ?? undefined)) : undefined,
    caps: toggle(child(rPr, "caps")),
    smallCaps: toggle(child(rPr, "smallCaps")),
    vertAlign: attr(child(rPr, "vertAlign"), "val") ?? undefined,
  });
}

const UNDERLINES: Record<string, UnderlineStyle> = {
  single: "single",
  words: "words",
  double: "double",
  thick: "thick",
  dotted: "dotted",
  dottedHeavy: "dotted",
  dash: "dash",
  dashedHeavy: "dash",
  dashLong: "dash",
  dashLongHeavy: "dash",
  dotDash: "dotDash",
  dashDotHeavy: "dotDash",
  dotDotDash: "dotDash",
  dashDotDotHeavy: "dotDash",
  wave: "wave",
  wavyHeavy: "wave",
  wavyDouble: "wave",
};

function mergePara(base: RawPara, over: RawPara): RawPara {
  const merged: RawPara = { ...base, ...over };
  if (over.hanging !== undefined && over.firstLine === undefined) delete merged.firstLine;
  if (over.firstLine !== undefined && over.hanging === undefined) delete merged.hanging;
  if (over.tabs) {
    const stops = new Map((base.tabs ?? []).map((t) => [t.pos, t]));
    for (const t of over.tabs) {
      if (t.val === "clear") stops.delete(t.pos);
      else stops.set(t.pos, t);
    }
    merged.tabs = [...stops.values()];
  }
  return merged;
}

class StyleSheet {
  private readonly styles = new Map<string, StyleDef>();
  private defaultPPr: RawPara = {};
  private defaultRPr: RawRun = {};
  private defaultParaStyle?: string;

  constructor(xml: string | undefined, readonly theme: ThemeFonts) {
    if (!xml) return;
    const root = parseXml(xml);
    const docDefaults = child(root, "docDefaults");
    this.defaultPPr = parsePPr(child(child(docDefaults, "pPrDefault"), "pPr"));
    this.defaultRPr = parseRPr(child(child(docDefaults, "rPrDefault"), "rPr"), theme);
    for (const el of elements(root)) {
      if (el.localName !== "style") continue;
      const id = attr(el, "styleId") ?? "";
      const type = attr(el, "type") ?? "paragraph";
      this.styles.set(id, {
        type,
        basedOn: attr(child(el, "basedOn"), "val") ?? undefined,
        name: attr(child(el, "name"), "val") ?? id,
        next: attr(child(el, "next"), "val") ?? undefined,
        gallery: !!toggle(child(el, "qFormat")),
        pPr: parsePPr(child(el, "pPr")),
        rPr: parseRPr(child(el, "rPr"), theme),
      });
      if (type === "paragraph" && onOff(attr(el, "default"))) {
        this.defaultParaStyle = id;
      }
    }
  }

  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const style = this.styles.get(cur);
      if (!style) break;
      out.unshift(style);
      cur = style.basedOn;
    }
    return out;
  }

  paraStyleId(id: string | null | undefined) {
    return id && this.styles.has(id) ? id : this.defaultParaStyle;
  }

  styleName(id: string | undefined) {
    return id ? (this.styles.get(id)?.name ?? id) : "";
  }

  paragraph(styleId: string | undefined, direct: RawPara): RawPara {
    let p = this.defaultPPr;
    for (const s of this.chain(styleId)) p = mergePara(p, s.pPr);
    return mergePara(p, direct);
  }

  run(paraStyleId: string | undefined, runStyleId: string | undefined, direct: RawRun): RawRun {
    let r: RawRun = { ...this.defaultRPr };
    for (const s of this.chain(paraStyleId)) r = { ...r, ...s.rPr };
    for (const s of this.chain(runStyleId)) r = { ...r, ...s.rPr };
    return { ...r, ...direct };
  }

  charStyle(id: string): RawRun {
    return this.run(undefined, id, {});
  }

  /** Paragraph and character styles, flattened, for the styles gallery. */
  definitions(settings: Settings): StyleDefinition[] {
    const out: StyleDefinition[] = [];
    for (const [id, s] of this.styles) {
      if (s.type === "paragraph") {
        const mark = this.run(id, undefined, {});
        const format = toFormat(this.paragraph(id, {}), mark, settings, { styleId: id });
        // Font and size are in `format`; `run` keeps the style's other character formatting.
        const run = styleRunFormat(mark);
        delete run.fontFamily;
        delete run.fontSize;
        const hasRun = Object.values(run).some((v) => v);
        out.push(defined({
          id,
          name: s.name,
          type: "paragraph",
          format,
          run: hasRun ? run : undefined,
          next: s.next,
          gallery: s.gallery || undefined,
        }) as StyleDefinition);
      } else if (s.type === "character") {
        let raw: RawRun = {};
        for (const c of this.chain(id)) raw = { ...raw, ...c.rPr };
        out.push(defined({ id, name: s.name, type: "character", run: styleRunFormat(raw), gallery: s.gallery || undefined }) as StyleDefinition);
      }
    }
    return out;
  }
}

/** The run properties a character style sets, as `RunFormat` fields. */
function styleRunFormat(run: RawRun): Partial<RunFormat> {
  return defined({
    bold: run.bold,
    italic: run.italic,
    fontFamily: run.font,
    fontSize: run.size === undefined ? undefined : run.size / 2,
    underline:
      run.underline === undefined
        ? undefined
        : run.underline === "none"
          ? null
          : (UNDERLINES[run.underline] ?? "single"),
    strike: run.dstrike ? "double" : run.strike ? "single" : undefined,
    color: run.color && run.color !== "auto" ? run.color : undefined,
    highlight: run.highlight && run.highlight !== "none" ? run.highlight : undefined,
    caps: run.caps ? "all" : run.smallCaps ? "small" : undefined,
    vertAlign:
      run.vertAlign === "superscript" || run.vertAlign === "subscript" ? run.vertAlign : undefined,
  }) as Partial<RunFormat>;
}

function parseTheme(xml: string | undefined): ThemeFonts {
  if (!xml) return {};
  const root = parseXml(xml);
  const find = (el: XElement, name: string): XElement | undefined => {
    for (const c of elements(el)) {
      if (c.localName === name) return c;
      const deep = find(c, name);
      if (deep) return deep;
    }
    return undefined;
  };
  const latin = (name: string) => {
    const font = find(root, name);
    const l = font ? child(font, "latin") : undefined;
    return l ? (l.getAttributeNS(A_NS, "typeface") || l.getAttribute("typeface") || undefined) : undefined;
  };
  return { major: latin("majorFont"), minor: latin("minorFont") };
}

type Settings = {
  defaultTabStop?: number;
  htmlAutoSpacing: boolean;
  compatibilityMode?: number;
  evenAndOddHeaders: boolean;
};

function parseSettings(xml: string | undefined): Settings {
  if (!xml) return { htmlAutoSpacing: true, evenAndOddHeaders: false };
  const root = parseXml(xml);
  const compat = child(root, "compat");
  const mode = compat
    ? elements(compat).find(
        (c) => c.localName === "compatSetting" && attr(c, "name") === "compatibilityMode",
      )
    : undefined;
  return {
    defaultTabStop: num(child(root, "defaultTabStop"), "val"),
    htmlAutoSpacing: !toggle(child(compat, "doNotUseHTMLParagraphAutoSpacing")),
    compatibilityMode: num(mode, "val"),
    evenAndOddHeaders: !!toggle(child(root, "evenAndOddHeaders")),
  };
}

/** Word 2013+ (compatibility mode 15) shrinks spaces in justified lines by up to 25%. */
const JUSTIFY_SHRINK = 0.25;

function fontSizePt(run: RawRun) {
  return (run.size ?? DEFAULT_FONT_SIZE_PT * 2) / 2;
}

const ALIGN: Record<string, TextAlign> = {
  both: "justify",
  distribute: "justify",
  center: "center",
  right: "right",
  end: "right",
  left: "left",
  start: "left",
};

function toFormat(
  p: RawPara,
  mark: RawRun,
  settings: Settings,
  options: { dropIndent?: boolean; styleId?: string } = {},
): ParagraphFormat {
  const px = (twips: number | undefined) =>
    twips === undefined ? undefined : twips / TWIPS_PER_PX;
  const autoPx = (auto: boolean | undefined, explicit: number | undefined) =>
    auto && settings.htmlAutoSpacing ? AUTO_SPACING_PX : px(explicit);
  const lineRule = p.lineRule === "exact" ? "exact" : p.lineRule === "atLeast" ? "atLeast" : "auto";
  const tabs: TabStop[] = (p.tabs ?? [])
    .filter((t) => t.val !== "clear" && t.val !== "bar")
    .map((t): TabStop => ({
      pos: t.pos / TWIPS_PER_PX,
      align:
        t.val === "center" || t.val === "right" || t.val === "decimal"
          ? t.val
          : t.val === "end"
            ? "right"
            : "left",
    }))
    .sort((a, b) => a.pos - b.pos);
  const format: ParagraphFormat = {
    align: p.jc ? ALIGN[p.jc] : undefined,
    indentLeft: options.dropIndent ? undefined : px(p.left),
    indentRight: px(p.right),
    indentFirstLine: options.dropIndent
      ? undefined
      : p.hanging
        ? -p.hanging / TWIPS_PER_PX
        : px(p.firstLine),
    spaceBefore: autoPx(p.beforeAuto, p.before),
    spaceAfter: autoPx(p.afterAuto, p.after),
    lineRule,
    lineValue:
      p.line === undefined
        ? 1
        : lineRule === "auto"
          ? p.line / 240
          : p.line / TWIPS_PER_PX,
    fontFamily: mark.font ?? DEFAULT_FONT,
    fontSize: fontSizePt(mark),
    tabs: tabs.length ? tabs : undefined,
    widowControl: p.widowControl || undefined,
    keepNext: p.keepNext || undefined,
    keepLines: p.keepLines || undefined,
    suppressLineNumbers: p.suppressLineNumbers || undefined,
    pageBreakBefore: p.pageBreakBefore || undefined,
    contextualSpacing: p.contextualSpacing || undefined,
    styleId: options.styleId,
  };
  return defined(format) as ParagraphFormat;
}

// —— Numbering ——

const NUM_FORMATS: Record<string, NumFormat> = {
  decimal: "decimal",
  decimalZero: "decimal",
  lowerLetter: "lowerLetter",
  upperLetter: "upperLetter",
  lowerRoman: "lowerRoman",
  upperRoman: "upperRoman",
};

/** Bullet characters, including Symbol/Wingdings private-use code points. */
const BULLETS: Record<string, BulletStyle> = {
  "\u2022": "disc",
  "\uf0b7": "disc",
  "\u00b7": "disc",
  o: "circle",
  "\u25cb": "circle",
  "\u25aa": "square",
  "\u25a0": "square",
  "\uf0a7": "square",
  "\uf06e": "square",
  "\u2013": "dash",
  "-": "dash",
  "\u27a2": "arrow",
  "\uf0d8": "arrow",
  "\u2713": "check",
  "\uf0fc": "check",
  "\u2756": "diamond",
  "\uf076": "diamond",
};

function parseLevel(lvl: XElement): ListLevel {
  const fmt = attr(child(lvl, "numFmt"), "val") ?? "";
  const text = attr(child(lvl, "lvlText"), "val") ?? "";
  const start = Number(attr(child(lvl, "start"), "val") ?? 1);
  if (!fmt || fmt === "bullet" || fmt === "none") {
    return { kind: "bullet", format: null, bullet: BULLETS[text] ?? "disc", start: 1 };
  }
  const legal = !!child(lvl, "isLgl") || (text.match(/%\d/g)?.length ?? 0) > 1;
  return { kind: "ordered", format: legal ? "legal" : (NUM_FORMATS[fmt] ?? "decimal"), bullet: null, start };
}

/** List levels of every numId, from word/numbering.xml. */
function parseNumbering(xml: string | undefined) {
  const abstracts = new Map<string, Map<number, ListLevel>>();
  const nums = new Map<string, { abstractId: string; overrides: Map<number, number> }>();
  if (xml) {
    const root = parseXml(xml);
    for (const el of elements(root)) {
      if (el.localName === "abstractNum") {
        const levels = new Map<number, ListLevel>();
        for (const lvl of elements(el)) {
          if (lvl.localName === "lvl") levels.set(Number(attr(lvl, "ilvl") ?? 0), parseLevel(lvl));
        }
        // Outline numbering whose second level is "%1.%2." numbers its first level that way too.
        const first = levels.get(0);
        if (first?.kind === "ordered" && levels.get(1)?.format === "legal") first.format = "legal";
        abstracts.set(attr(el, "abstractNumId") ?? "", levels);
      } else if (el.localName === "num") {
        const overrides = new Map<number, number>();
        for (const o of elements(el)) {
          if (o.localName !== "lvlOverride") continue;
          const start = attr(child(o, "startOverride"), "val");
          if (start !== null && start !== undefined) overrides.set(Number(attr(o, "ilvl") ?? 0), Number(start));
        }
        nums.set(attr(el, "numId") ?? "", {
          abstractId: attr(child(el, "abstractNumId"), "val") ?? "",
          overrides,
        });
      }
    }
  }
  return (numId: string, level: number): Omit<ListInfo, "level"> => {
    const num = nums.get(numId);
    const abstractId = num?.abstractId ?? "";
    const info = abstracts.get(abstractId)?.get(level) ?? {
      kind: "bullet" as const,
      format: null,
      bullet: "disc" as const,
      start: 1,
    };
    return { ...info, numId, abstractId, startOverride: num?.overrides.get(level) };
  };
}

// —— Body ——

type ChangeType = "insertion" | "deletion";

type ParsedSection = { start: SectionStart; layout: SectionLayout };

type Context = {
  styles: StyleSheet;
  settings: Settings;
  listLevel: (numId: string, level: number) => Omit<ListInfo, "level">;
  ids: { next: number };
  /** Every `w:sectPr` in document order; the last is the body's. */
  sections: ParsedSection[];
  /** Section breaks emitted so far. */
  sectionIndex: number;
  comments: CommentImport;
  /** document.xml.rels targets by id. */
  rels: Map<string, string>;
  /** Image data URLs by relationship id. */
  images: Map<string, string>;
  /** The `w:hyperlink` / HYPERLINK `fldSimple` being read. */
  link: LinkTarget | null;
  /** Open complex fields, innermost last. */
  fields: { instr: string; result: boolean; link: LinkTarget | null }[];
};

// —— Comments ——

type CommentImport = {
  /** Top-level threads by Word comment id. */
  threads: Map<string, CommentThread>;
  /** Comment ranges open at the current position (Word ids), outermost first. */
  open: string[];
  /** Threads that got at least one marked run. */
  anchored: Set<string>;
};

const W14_PARA_ID = "w14:paraId";

function commentText(comment: XElement): { text: string; lastParaId: string | null } {
  const paragraphs = elements(comment).filter((p) => p.localName === "p");
  const text = paragraphs
    .map((p) =>
      Array.from(p.getElementsByTagNameNS(W_NS, "t"), (t) => t.textContent ?? "").join(""),
    )
    .join("\n");
  return { text, lastParaId: paragraphs[paragraphs.length - 1]?.getAttribute(W14_PARA_ID) ?? null };
}

/** `comments.xml` threads; `commentsExtended.xml` (Word 2013+) adds replies and Resolved. */
function parseComments(xml: string | undefined, extendedXml: string | undefined): CommentImport {
  const threads = new Map<string, CommentThread>();
  const result: CommentImport = { threads, open: [], anchored: new Set() };
  if (!xml) return result;
  const extended = new Map<string, { parent: string | null; done: boolean }>();
  if (extendedXml) {
    for (const ex of elements(parseXml(extendedXml))) {
      const paraId = ex.getAttribute("w15:paraId");
      if (!paraId) continue;
      extended.set(paraId, {
        parent: ex.getAttribute("w15:paraIdParent"),
        done: ex.getAttribute("w15:done") === "1",
      });
    }
  }
  const byParaId = new Map<string, CommentThread>();
  const replies: { reply: CommentReply; parentParaId: string }[] = [];
  for (const el of elements(parseXml(xml))) {
    if (el.localName !== "comment") continue;
    const wid = attr(el, "id");
    if (wid === null) continue;
    const author = attr(el, "author") ?? "";
    const { text, lastParaId } = commentText(el);
    const entry: CommentReply = {
      id: `c${wid}`,
      authorId: author,
      authorName: author,
      date: attr(el, "date") ?? "",
      text,
    };
    const ex = lastParaId ? extended.get(lastParaId) : undefined;
    if (ex?.parent) {
      replies.push({ reply: { ...entry, id: `r${wid}` }, parentParaId: ex.parent });
      continue;
    }
    const thread: CommentThread = ex?.done ? { ...entry, resolved: true } : entry;
    threads.set(wid, thread);
    if (lastParaId) byParaId.set(lastParaId, thread);
  }
  for (const { reply, parentParaId } of replies) {
    const parent = byParaId.get(parentParaId);
    if (parent) parent.replies = [...(parent.replies ?? []), reply];
  }
  return result;
}

function openCommentRange(el: XElement, ctx: Context) {
  const wid = attr(el, "id");
  if (wid !== null && ctx.comments.threads.has(wid)) ctx.comments.open.push(wid);
}

function closeCommentRange(el: XElement, ctx: Context) {
  const wid = attr(el, "id");
  ctx.comments.open = ctx.comments.open.filter((id) => id !== wid);
}

function commentMarks(ctx: Context): MarkJSON[] {
  return ctx.comments.open.map((wid) => {
    ctx.comments.anchored.add(wid);
    const thread = ctx.comments.threads.get(wid)!;
    return { type: "comment", attrs: { id: thread.id, authorId: thread.authorId } };
  });
}

/** Marks kept on tabs, breaks and images (formatting marks only apply to text). */
const INLINE_NODE_MARKS = new Set(["insertion", "deletion", "comment", "link"]);

/** Collects one Word paragraph, splitting it around hard page breaks. */
class ParagraphBuilder {
  private inline: NodeJSON[] = [];
  private emitted = false;
  private lastChange: { type: ChangeType; author: string; date: string; id: string } | null =
    null;
  private lastFormat: { key: string; id: string } | null = null;

  constructor(
    private readonly blocks: Block[],
    private readonly make: (content: NodeJSON[]) => Block,
    private readonly ids: { next: number },
    /** Run properties of the paragraph style; runs differing from the mark get a textStyle. */
    readonly paraStyleId: string | undefined,
    readonly markFont: { family: string; size: number },
  ) {}

  changeMark(type: ChangeType, el: XElement): MarkJSON {
    const author = attr(el, "author") ?? "";
    const date = attr(el, "date") ?? "";
    const last = this.lastChange;
    const id =
      last && last.type === type && last.author === author && last.date === date
        ? last.id
        : `w${attr(el, "id") ?? ""}-${this.ids.next++}`;
    this.lastChange = { type, author, date, id };
    return {
      type,
      attrs: { id, authorId: author, authorName: author, date },
    };
  }

  /** Mark for a `w:rPrChange`; consecutive runs with the same revision share one id. */
  formatChangeMark(el: XElement | undefined, old: RunFormat, now: RunFormat): MarkJSON | null {
    const description = el ? describeFormatChange(old, now) : "";
    if (!el || !description) {
      this.lastFormat = null;
      return null;
    }
    const author = attr(el, "author") ?? "";
    const date = attr(el, "date") ?? "";
    const key = `${author}|${date}|${description}|${JSON.stringify(old)}`;
    const id =
      this.lastFormat?.key === key ? this.lastFormat.id : `f${attr(el, "id") ?? ""}-${this.ids.next++}`;
    this.lastFormat = { key, id };
    return {
      type: "formatChange",
      attrs: { id, authorId: author, authorName: author, date, description, old },
    };
  }

  plainRun() {
    this.lastChange = null;
  }

  text(text: string, marks: MarkJSON[]) {
    if (!text) return;
    this.inline.push(marks.length ? { type: "text", text, marks } : { type: "text", text });
  }

  inlineNode(type: "tab" | "hard_break" | "image", marks: MarkJSON[], attrs?: Record<string, unknown>) {
    const kept = marks.filter((m) => INLINE_NODE_MARKS.has(m.type));
    this.inline.push({ type, ...(attrs ? { attrs } : {}), ...(kept.length ? { marks: kept } : {}) });
  }

  pageBreak() {
    if (this.inline.length) this.flush();
    this.blocks.push({ node: { type: "page_break" } });
    this.emitted = true;
  }

  end() {
    if (this.inline.length || !this.emitted) this.flush();
  }

  private flush() {
    this.blocks.push(this.make(this.inline));
    this.inline = [];
    this.emitted = true;
  }
}

/** A run's character formatting relative to the paragraph mark's font. */
function runFormat(run: RawRun, b: ParagraphBuilder): RunFormat {
  const family = run.font ?? DEFAULT_FONT;
  const size = fontSizePt(run);
  const highlight = run.highlight && run.highlight !== "none" ? run.highlight : null;
  const shading = run.shading && run.shading !== "none" ? run.shading : null;
  const vert = run.vertAlign;
  return {
    bold: !!run.bold,
    italic: !!run.italic,
    fontFamily: family !== b.markFont.family ? family : null,
    fontSize: size !== b.markFont.size ? size : null,
    underline: run.underline && run.underline !== "none" ? (UNDERLINES[run.underline] ?? "single") : null,
    strike: run.dstrike ? "double" : run.strike ? "single" : null,
    color: run.color && run.color !== "auto" ? run.color : null,
    highlight: highlight ?? shading,
    caps: run.caps ? "all" : run.smallCaps ? "small" : null,
    vertAlign: vert === "superscript" || vert === "subscript" ? vert : null,
  };
}

function runMarks(format: RunFormat, change: MarkJSON | null): MarkJSON[] {
  return change ? [change, ...runFormatMarks(format)] : runFormatMarks(format);
}

function resolvedRPr(rPr: XElement | undefined, b: ParagraphBuilder, ctx: Context): RawRun {
  return ctx.styles.run(
    b.paraStyleId,
    attr(child(rPr, "rStyle"), "val") ?? undefined,
    parseRPr(rPr, ctx.styles.theme),
  );
}

function walkRun(run: XElement, b: ParagraphBuilder, change: MarkJSON | null, ctx: Context) {
  if (!change) b.plainRun();
  const rPr = child(run, "rPr");
  const props = resolvedRPr(rPr, b, ctx);
  if (props.vanish) return;
  const format = runFormat(props, b);
  const marks = runMarks(format, change);
  const rPrChange = child(rPr, "rPrChange");
  const old = rPrChange ? runFormat(resolvedRPr(child(rPrChange, "rPr"), b, ctx), b) : format;
  const formatChange = b.formatChangeMark(rPrChange, old, format);
  if (formatChange) marks.push(formatChange);
  marks.push(...commentMarks(ctx));
  const withLink = () => [...marks, ...linkMarks(ctx)];

  for (const el of elements(run)) {
    switch (el.localName) {
      case "t":
      case "delText":
        b.text(el.textContent ?? "", withLink());
        break;
      case "tab":
      case "ptab":
        b.inlineNode("tab", withLink());
        break;
      case "noBreakHyphen":
        b.text("-", withLink());
        break;
      case "br":
        if (attr(el, "type") === "page") b.pageBreak();
        else b.inlineNode("hard_break", withLink());
        break;
      case "cr":
        b.inlineNode("hard_break", withLink());
        break;
      case "fldChar":
        fieldChar(attr(el, "fldCharType"), ctx);
        break;
      case "instrText": {
        const field = ctx.fields[ctx.fields.length - 1];
        if (field && !field.result) field.instr += el.textContent ?? "";
        break;
      }
      case "drawing":
      case "pict": {
        const image = el.localName === "drawing" ? drawingImage(el, ctx) : vmlImage(el, ctx);
        if (image) b.inlineNode("image", withLink(), image);
        break;
      }
    }
  }
}

// —— Hyperlinks and images ——

type LinkTarget = { href: string; title: string | null };

const HYPERLINK_INSTR = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"(?:.*\\o\s+"([^"]*)")?/;

function hyperlinkField(instr: string): LinkTarget | null {
  const m = HYPERLINK_INSTR.exec(instr);
  if (!m) return null;
  return { href: /\\l\s/.test(instr) ? `#${m[1]}` : m[1], title: m[2] ?? null };
}

/** Complex fields (`w:fldChar` begin / separate / end); only HYPERLINK results are kept as links. */
function fieldChar(type: string | null, ctx: Context) {
  if (type === "begin") {
    ctx.fields.push({ instr: "", result: false, link: null });
  } else if (type === "separate") {
    const field = ctx.fields[ctx.fields.length - 1];
    if (field) {
      field.result = true;
      field.link = hyperlinkField(field.instr);
    }
  } else if (type === "end") {
    ctx.fields.pop();
  }
}

function linkMarks(ctx: Context): MarkJSON[] {
  const target =
    ctx.link ?? [...ctx.fields].reverse().find((f) => f.result && f.link)?.link ?? null;
  return target ? [{ type: "link", attrs: { href: target.href, title: target.title } }] : [];
}

function hyperlinkTarget(el: XElement, ctx: Context): LinkTarget | null {
  const id = el.getAttributeNS(R_NS, "id") || el.getAttribute("r:id");
  const anchor = attr(el, "anchor");
  const base = id ? ctx.rels.get(id) : undefined;
  const href = base ? (anchor ? `${base}#${anchor}` : base) : anchor ? `#${anchor}` : null;
  return href ? { href, title: attr(el, "tooltip") } : null;
}

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const EMU_PER_PX = 9525;

const IMAGE_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
};

/** Data URLs of the document's images by relationship id. */
async function loadImages(zip: JSZip, rels: Map<string, string>) {
  const images = new Map<string, string>();
  for (const [id, target] of rels) {
    const mime = IMAGE_TYPES[target.split(".").pop()?.toLowerCase() ?? ""];
    if (!mime) continue;
    const data = await zip.file(`word/${target}`)?.async("base64");
    if (data) images.set(id, `data:${mime};base64,${data}`);
  }
  return images;
}

function firstNS(el: XElement, ns: string, name: string): XElement | undefined {
  return el.getElementsByTagNameNS(ns, name)[0];
}

type ImageJSON = { src: string; alt: string; width: number | null; height: number | null };

/** `w:drawing` (inline or anchored picture), shown inline. */
function drawingImage(drawing: XElement, ctx: Context): ImageJSON | null {
  const blip = firstNS(drawing, A_NS, "blip");
  const id = blip && (blip.getAttributeNS(R_NS, "embed") || blip.getAttribute("r:embed"));
  const src = id ? ctx.images.get(id) : undefined;
  if (!src) return null;
  const extent = firstNS(drawing, WP_NS, "extent");
  const cx = Number(extent?.getAttribute("cx"));
  const cy = Number(extent?.getAttribute("cy"));
  const docPr = firstNS(drawing, WP_NS, "docPr");
  return {
    src,
    alt: docPr?.getAttribute("descr") || docPr?.getAttribute("title") || "",
    width: cx ? Math.round(cx / EMU_PER_PX) : null,
    height: cy ? Math.round(cy / EMU_PER_PX) : null,
  };
}

const VML_NS = "urn:schemas-microsoft-com:vml";

/** Legacy VML picture (`w:pict` / `v:imagedata`). */
function vmlImage(pict: XElement, ctx: Context): ImageJSON | null {
  const data = firstNS(pict, VML_NS, "imagedata");
  const id = data && (data.getAttributeNS(R_NS, "id") || data.getAttribute("r:id"));
  const src = id ? ctx.images.get(id) : undefined;
  if (!src) return null;
  const style = firstNS(pict, VML_NS, "shape")?.getAttribute("style") ?? "";
  const size = (prop: string) => {
    const m = new RegExp(`${prop}:\\s*([\\d.]+)pt`).exec(style);
    return m ? Math.round(Number(m[1]) / 0.75) : null;
  };
  return { src, alt: data?.getAttribute("o:title") ?? "", width: size("width"), height: size("height") };
}

function walkInline(el: XElement, b: ParagraphBuilder, change: MarkJSON | null, ctx: Context) {
  for (const c of elements(el)) {
    switch (c.localName) {
      case "r":
        walkRun(c, b, change, ctx);
        break;
      case "ins":
      case "moveTo":
        walkInline(c, b, b.changeMark("insertion", c), ctx);
        break;
      case "del":
      case "moveFrom":
        walkInline(c, b, b.changeMark("deletion", c), ctx);
        break;
      case "commentRangeStart":
        openCommentRange(c, ctx);
        break;
      case "commentRangeEnd":
        closeCommentRange(c, ctx);
        break;
      case "hyperlink":
      case "fldSimple": {
        const target =
          c.localName === "hyperlink" ? hyperlinkTarget(c, ctx) : hyperlinkField(attr(c, "instr") ?? "");
        const outer = ctx.link;
        if (target) ctx.link = target;
        walkInline(c, b, change, ctx);
        ctx.link = outer;
        break;
      }
      case "smartTag":
      case "customXml":
        walkInline(c, b, change, ctx);
        break;
      case "sdt": {
        const content = child(c, "sdtContent");
        if (content) walkInline(content, b, change, ctx);
        break;
      }
    }
  }
}

function headingLevel(styleName: string): number | null {
  const m = styleName.match(/^heading\s*([1-9])$/i);
  return m ? Math.min(Number(m[1]), 6) : null;
}

/** Resolved paragraph format plus the style id used for run inheritance. */
function resolveParagraph(pPr: XElement | undefined, ctx: Context, dropIndent = false) {
  const styleId = ctx.styles.paraStyleId(attr(child(pPr, "pStyle"), "val"));
  const raw = ctx.styles.paragraph(styleId, parsePPr(pPr));
  const mark = ctx.styles.run(styleId, undefined, parseRPr(child(pPr, "rPr"), ctx.styles.theme));
  const format = toFormat(raw, mark, ctx.settings, { dropIndent, styleId });
  return { styleId, raw, format };
}

/** A `w:pPrChange` as a paragraph revision, with the format before it. */
function paragraphChange(
  pPr: XElement | undefined,
  format: ParagraphFormat,
  dropIndent: boolean,
  ctx: Context,
): ParagraphChange | null {
  const el = child(pPr, "pPrChange");
  if (!el) return null;
  const oldPPr = child(el, "pPr");
  const old = resolveParagraph(oldPPr, ctx, dropIndent).format;
  // The old pPr has no run properties, so keep the paragraph mark's font.
  const oldFormat = defined({ ...old, fontFamily: format.fontFamily, fontSize: format.fontSize }) as ParagraphFormat;
  const author = attr(el, "author") ?? "";
  return {
    id: `p${attr(el, "id") ?? ""}-${ctx.ids.next++}`,
    authorId: author,
    authorName: author,
    date: attr(el, "date") ?? "",
    old: oldFormat,
    description: describeParagraphChange(oldFormat, format) || "Paragraph properties",
  };
}

function walkParagraph(p: XElement, blocks: Block[], ctx: Context) {
  const pPr = child(p, "pPr");
  const numPr = child(pPr, "numPr");
  const numId = attr(child(numPr, "numId"), "val");
  const ilvl = Number(attr(child(numPr, "ilvl"), "val") ?? 0);
  const list: ListInfo | undefined =
    numId && numId !== "0" ? { ...ctx.listLevel(numId, ilvl), level: ilvl } : undefined;
  const { styleId, format } = resolveParagraph(pPr, ctx, !!list);
  const level = headingLevel(ctx.styles.styleName(styleId));
  const paraChange = paragraphChange(pPr, format, !!list, ctx);
  const attrs = paraChange ? { format, paraChange } : { format };

  const make = (content: NodeJSON[]): Block => {
    const body = content.length ? { content } : {};
    if (list) return { node: { type: "paragraph", attrs, ...body }, list };
    if (level) return { node: { type: "heading", attrs: { level, ...attrs }, ...body } };
    return { node: { type: "paragraph", attrs, ...body } };
  };
  const builder = new ParagraphBuilder(blocks, make, ctx.ids, styleId, {
    family: format.fontFamily ?? DEFAULT_FONT,
    size: format.fontSize ?? DEFAULT_FONT_SIZE_PT,
  });
  walkInline(p, builder, null, ctx);
  builder.end();

  if (child(pPr, "sectPr")) {
    const index = ctx.sectionIndex++;
    blocks.push({
      node: {
        type: "section_break",
        attrs: {
          start: ctx.sections[index + 1]?.start ?? "nextPage",
          section: ctx.sections[index]?.layout ?? null,
        },
      },
    });
  }
}

function walkBody(el: XElement, blocks: Block[], ctx: Context) {
  for (const c of elements(el)) {
    switch (c.localName) {
      case "p":
        walkParagraph(c, blocks, ctx);
        break;
      case "commentRangeStart":
        openCommentRange(c, ctx);
        break;
      case "commentRangeEnd":
        closeCommentRange(c, ctx);
        break;
      case "tbl": {
        const table = walkTable(c, ctx);
        if (table) blocks.push({ node: table });
        break;
      }
      case "sdt": {
        const content = child(c, "sdtContent");
        if (content) walkBody(content, blocks, ctx);
        break;
      }
    }
  }
}

const NO_BORDER = new Set(["nil", "none"]);

/** Whether a table shows gridlines: its own `tblBorders`, else a "…Grid…" table style. */
function tableHasBorders(tblPr: XElement | undefined): boolean {
  const borders = child(tblPr, "tblBorders");
  if (borders) return elements(borders).some((b) => !NO_BORDER.has(attr(b, "val") ?? "nil"));
  return /grid/i.test(attr(child(tblPr, "tblStyle"), "val") ?? "");
}

const VALIGN: Record<string, string> = { center: "middle", bottom: "bottom" };

/** A Word table; vertically merged cells (`w:vMerge`) become `rowspan`. */
function walkTable(tbl: XElement, ctx: Context): NodeJSON | null {
  const grid = elements(child(tbl, "tblGrid") ?? tbl)
    .filter((g) => g.localName === "gridCol")
    .map((g) => (num(g, "w") ?? 0) / TWIPS_PER_PX);
  /** The cell each grid column's vertical merge continues from. */
  const merging = new Map<number, NodeJSON & { attrs: Record<string, unknown> }>();
  const rows: NodeJSON[] = [];
  for (const tr of elements(tbl).filter((r) => r.localName === "tr")) {
    const cells: NodeJSON[] = [];
    let col = num(child(child(tr, "trPr"), "gridBefore"), "val") ?? 0;
    for (const tc of elements(tr).filter((x) => x.localName === "tc")) {
      const tcPr = child(tc, "tcPr");
      const span = Math.max(1, num(child(tcPr, "gridSpan"), "val") ?? 1);
      const vMergeEl = child(tcPr, "vMerge");
      const vMerge = vMergeEl ? (attr(vMergeEl, "val") ?? "continue") : null;
      if (vMerge === "continue" && merging.has(col)) {
        const above = merging.get(col)!;
        above.attrs.rowspan = Number(above.attrs.rowspan) + 1;
        col += span;
        continue;
      }
      const blocks: Block[] = [];
      walkBody(tc, blocks, ctx);
      const content = groupLists(blocks);
      const widths = grid.slice(col, col + span);
      const fill = attr(child(tcPr, "shd"), "fill");
      const cell = {
        type: "table_cell",
        attrs: {
          colspan: span,
          rowspan: 1,
          colwidth: widths.length === span && widths.every((w) => w > 0) ? widths.map(Math.round) : null,
          background: fill && fill !== "auto" ? normalizeHex(`#${fill}`) : null,
          valign: VALIGN[attr(child(tcPr, "vAlign"), "val") ?? ""] ?? null,
        } as Record<string, unknown>,
        content: content.length ? content : [{ type: "paragraph" }],
      };
      for (let c = col; c < col + span; c++) {
        if (vMerge === "restart") merging.set(c, cell);
        else merging.delete(c);
      }
      cells.push(cell);
      col += span;
    }
    if (cells.length) rows.push({ type: "table_row", content: cells });
  }
  if (!rows.length) return null;
  return { type: "table", attrs: { borders: tableHasBorders(child(tbl, "tblPr")) }, content: rows };
}

/**
 * Group consecutive list paragraphs into (nested) bullet/ordered lists. Numbering
 * continues across lists that share an abstract numbering, as in Word: each level counts
 * on until a `startOverride` (Restart at 1) or a higher-level item resets it.
 */
function groupLists(blocks: Block[]): NodeJSON[] {
  const out: NodeJSON[] = [];
  const counters = new Map<string, number>();
  const restarted = new Set<string>();
  const number = (list: ListInfo) => {
    const key = `${list.abstractId}:${list.level}`;
    if (list.startOverride !== undefined && !restarted.has(`${list.numId}:${list.level}`)) {
      restarted.add(`${list.numId}:${list.level}`);
      counters.set(key, list.startOverride - 1);
    }
    const n = (counters.get(key) ?? list.start - 1) + 1;
    counters.set(key, n);
    for (let l = list.level + 1; l < 9; l++) counters.delete(`${list.abstractId}:${l}`);
    return n;
  };
  let i = 0;
  while (i < blocks.length) {
    if (!blocks[i].list) {
      out.push(blocks[i].node);
      i++;
      continue;
    }
    const run: { node: NodeJSON; list: ListInfo }[] = [];
    while (i < blocks.length && blocks[i].list) {
      const list = blocks[i].list!;
      const prevLevel = run.length ? run[run.length - 1].list.level : -1;
      run.push({ node: blocks[i].node, list: { ...list, level: Math.min(list.level, prevLevel + 1) } });
      i++;
    }
    let j = 0;
    const build = (level: number): NodeJSON => {
      const first = run[j].list;
      const items: NodeJSON[] = [];
      let order = 1;
      while (j < run.length && run[j].list.level >= level) {
        const entry = run[j];
        if (entry.list.level === level) {
          if (entry.list.kind !== first.kind && items.length) break;
          const n = number(entry.list);
          if (!items.length) order = n;
          items.push({ type: "list_item", content: [entry.node] });
          j++;
        } else {
          items[items.length - 1].content!.push(build(level + 1));
        }
      }
      return first.kind === "bullet"
        ? { type: "bullet_list", attrs: { bullet: first.bullet }, content: items }
        : {
            type: "ordered_list",
            attrs: { order, numFormat: first.format },
            content: items,
          };
    };
    while (j < run.length) out.push(build(0));
  }
  return out;
}

// —— Headers / footers ——

type HFRunProps = Omit<Extract<HFRun, { kind: "text" }>, "kind" | "text">;

function hfProps(run: RawRun, mark: ParagraphFormat): HFRunProps {
  const family = run.font ?? DEFAULT_FONT;
  const size = fontSizePt(run);
  return defined({
    bold: run.bold || undefined,
    italic: run.italic || undefined,
    underline: (run.underline && run.underline !== "none") || undefined,
    color: run.color && run.color !== "auto" ? run.color : undefined,
    fontFamily: family !== mark.fontFamily ? family : undefined,
    fontSize: size !== mark.fontSize ? size : undefined,
  });
}

function fieldName(instr: string): "PAGE" | "NUMPAGES" | null {
  const name = instr.trim().split(/\s+/)[0]?.toUpperCase();
  if (name === "PAGE") return "PAGE";
  if (name === "NUMPAGES" || name === "SECTIONPAGES") return "NUMPAGES";
  return null;
}

function parseHeaderFooter(xml: string, ctx: Context): HFParagraph[] {
  const root = parseXml(xml);
  const out: HFParagraph[] = [];
  const paragraphs: XElement[] = [];
  const collect = (el: XElement) => {
    for (const c of elements(el)) {
      if (c.localName === "p") paragraphs.push(c);
      else if (["tbl", "tr", "tc", "sdt", "sdtContent"].includes(c.localName ?? "")) collect(c);
    }
  };
  collect(root);

  for (const p of paragraphs) {
    const pPr = child(p, "pPr");
    const { styleId, format } = resolveParagraph(pPr, ctx);
    const runs: HFRun[] = [];
    const runProps = (r: XElement) => {
      const rPr = child(r, "rPr");
      return ctx.styles.run(
        styleId,
        attr(child(rPr, "rStyle"), "val") ?? undefined,
        parseRPr(rPr, ctx.styles.theme),
      );
    };
    // Complex field: the instruction is collected, then the cached result is
    // either replaced by a live PAGE/NUMPAGES field or shown as plain text.
    let field: { instr: string; props: HFRunProps; phase: "instr" | "skip" | "show" } | null =
      null;
    const emitField = () => {
      const name = field ? fieldName(field.instr) : null;
      if (field && name) runs.push({ kind: "field", field: name, ...field.props });
      return !!name;
    };

    const walk = (el: XElement) => {
      for (const c of elements(el)) {
        if (c.localName === "fldSimple") {
          const name = fieldName(attr(c, "instr") ?? "");
          const first = child(c, "r");
          if (name) {
            runs.push({ kind: "field", field: name, ...hfProps(first ? runProps(first) : {}, format) });
          } else {
            walk(c);
          }
          continue;
        }
        if (c.localName !== "r") {
          if (["ins", "moveTo", "hyperlink", "smartTag", "customXml", "sdt", "sdtContent"].includes(c.localName ?? "")) {
            walk(c);
          }
          continue;
        }
        const props = runProps(c);
        if (props.vanish) continue;
        const style = hfProps(props, format);
        for (const r of elements(c)) {
          if (r.localName === "fldChar") {
            const type = attr(r, "fldCharType");
            if (type === "begin") {
              field = { instr: "", props: style, phase: "instr" };
            } else if (type === "separate" && field) {
              field.phase = emitField() ? "skip" : "show";
            } else if (type === "end") {
              if (field?.phase === "instr") emitField();
              field = null;
            }
            continue;
          }
          if (r.localName === "instrText") {
            if (field?.phase === "instr") field.instr += r.textContent ?? "";
            continue;
          }
          if (field && field.phase !== "show") continue;
          if (r.localName === "t") runs.push({ kind: "text", text: r.textContent ?? "", ...style });
          else if (r.localName === "tab" || r.localName === "ptab") runs.push({ kind: "tab" });
        }
      }
    };
    walk(p);
    const frame = parseFrame(child(pPr, "framePr"));
    out.push(frame ? { format, runs, floating: true, frame } : { format, runs });
  }
  return out;
}

const ANCHORS: Record<string, HFFrame["hAnchor"]> = { page: "page", margin: "margin", text: "text" };

function parseFrame(framePr: XElement | undefined): HFFrame | undefined {
  if (!framePr) return undefined;
  const px = (name: string) => {
    const v = num(framePr, name);
    return v === undefined ? undefined : v / TWIPS_PER_PX;
  };
  return defined({
    hAnchor: ANCHORS[attr(framePr, "hAnchor") ?? ""] ?? "text",
    vAnchor: ANCHORS[attr(framePr, "vAnchor") ?? ""] ?? "margin",
    x: px("x"),
    y: px("y"),
    xAlign: (attr(framePr, "xAlign") ?? undefined) as HFFrame["xAlign"],
    yAlign: (attr(framePr, "yAlign") ?? undefined) as HFFrame["yAlign"],
    width: px("w"),
    height: px("h"),
  }) as HFFrame;
}

async function parseRels(zip: JSZip): Promise<Map<string, string>> {
  const xml = await zip.file("word/_rels/document.xml.rels")?.async("string");
  const rels = new Map<string, string>();
  if (!xml) return rels;
  for (const el of elements(parseXml(xml))) {
    const id = el.getAttribute("Id");
    const target = el.getAttribute("Target");
    if (id && target) rels.set(id, target.replace(/^\/?word\//, "").replace(/^\//, ""));
  }
  return rels;
}

async function parseHeaderFooterSet(
  sectPr: XElement,
  kind: "headerReference" | "footerReference",
  zip: JSZip,
  rels: Map<string, string>,
  ctx: Context,
): Promise<HeaderFooterSet | undefined> {
  const set: HeaderFooterSet = {};
  for (const ref of elements(sectPr).filter((c) => c.localName === kind)) {
    const type = attr(ref, "type") ?? "default";
    if (type !== "default" && type !== "first" && type !== "even") continue;
    const rid = ref.getAttributeNS(R_NS, "id") || ref.getAttribute("r:id");
    const target = rid ? rels.get(rid) : undefined;
    const xml = target ? await zip.file(`word/${target}`)?.async("string") : undefined;
    if (xml) set[type] = parseHeaderFooter(xml, ctx);
  }
  return set.default || set.first || set.even ? set : undefined;
}

/**
 * Word sections inherit each header/footer type they don't define from the previous
 * section; a title page or even pages without one stay blank.
 */
function inheritHeaderFooters(sections: ParsedSection[], evenAndOdd: boolean) {
  let headers: HeaderFooterSet = {};
  let footers: HeaderFooterSet = {};
  for (const { layout } of sections) {
    headers = { ...headers, ...layout.headers };
    footers = { ...footers, ...layout.footers };
    const complete = (set: HeaderFooterSet): HeaderFooterSet | undefined => {
      const out: HeaderFooterSet = { ...set };
      if (layout.titlePage && !out.first) out.first = [];
      if (evenAndOdd && !out.even) out.even = [];
      return out.default || out.first || out.even ? out : undefined;
    };
    layout.headers = complete(headers);
    layout.footers = complete(footers);
    if (!layout.headers) delete layout.headers;
    if (!layout.footers) delete layout.footers;
  }
}

const SECTION_STARTS: Record<string, SectionStart> = {
  nextPage: "nextPage",
  nextColumn: "nextPage",
  continuous: "continuous",
  evenPage: "evenPage",
  oddPage: "oddPage",
};

async function parseSection(
  sectPr: XElement,
  zip: JSZip,
  rels: Map<string, string>,
  ctx: Context,
): Promise<ParsedSection> {
  const pgMar = child(sectPr, "pgMar");
  const cols = child(sectPr, "cols");
  const colCount = num(cols, "num") ?? 1;
  const twipsPx = (v: number | undefined) => (v === undefined ? undefined : v / TWIPS_PER_PX);
  const layout = defined({
    page: parsePageSetup(sectPr),
    headerDistance: twipsPx(num(pgMar, "header")),
    footerDistance: twipsPx(num(pgMar, "footer")),
    headers: await parseHeaderFooterSet(sectPr, "headerReference", zip, rels, ctx),
    footers: await parseHeaderFooterSet(sectPr, "footerReference", zip, rels, ctx),
    titlePage: toggle(child(sectPr, "titlePg")) || undefined,
    lineNumbers: parseLineNumbers(sectPr, ctx.styles),
    pageNumberStart: num(child(sectPr, "pgNumType"), "start"),
    columns:
      colCount > 1 ? { count: colCount, space: twipsPx(num(cols, "space")) ?? 48 } : undefined,
  }) as SectionLayout;
  return {
    start: SECTION_STARTS[attr(child(sectPr, "type"), "val") ?? "nextPage"] ?? "nextPage",
    layout,
  };
}

/** Every `w:sectPr` in body order: those ending sections inside paragraphs, then the body's. */
function sectionProperties(body: XElement): XElement[] {
  const out: XElement[] = [];
  const visit = (el: XElement) => {
    for (const c of elements(el)) {
      if (c.localName === "p") {
        const sectPr = child(child(c, "pPr"), "sectPr");
        if (sectPr) out.push(sectPr);
      } else if (c.localName === "sdt") {
        const content = child(c, "sdtContent");
        if (content) visit(content);
      } else if (c.localName === "sectPr" && el === body) {
        out.push(c);
      }
    }
  };
  visit(body);
  return out;
}

function parsePageSetup(sectPr: XElement | undefined): PageSetup | undefined {
  if (!sectPr) return undefined;
  const pgSz = child(sectPr, "pgSz");
  const pgMar = child(sectPr, "pgMar");
  const px = (el: XElement | undefined, name: string, fallback: number) => {
    const v = Number(attr(el, name));
    return Number.isFinite(v) && v > 0 ? v / TWIPS_PER_PX : fallback;
  };
  const d = LETTER_PAGE_SETUP;
  return {
    width: px(pgSz, "w", d.width),
    height: px(pgSz, "h", d.height),
    marginTop: Math.abs(px(pgMar, "top", d.marginTop)),
    marginRight: px(pgMar, "right", d.marginRight),
    marginBottom: Math.abs(px(pgMar, "bottom", d.marginBottom)),
    marginLeft: px(pgMar, "left", d.marginLeft),
    gap: d.gap,
  };
}

function parseLineNumbers(sectPr: XElement, styles: StyleSheet): LineNumbering | null {
  const ln = child(sectPr, "lnNumType");
  if (!ln) return null;
  const countBy = num(ln, "countBy") ?? 1;
  if (countBy <= 0) return null;
  const restart = attr(ln, "restart");
  const distance = num(ln, "distance");
  const font = styles.charStyle("LineNumber");
  return {
    countBy,
    start: (num(ln, "start") ?? 0) + 1,
    restart: restart === "continuous" || restart === "newSection" ? restart : "newPage",
    distance: distance !== undefined ? distance / TWIPS_PER_PX : 24,
    fontFamily: font.font ?? DEFAULT_FONT,
    fontSize: fontSizePt(font),
  };
}

/** Parse a DOCX file, keeping Word revisions (w:ins / w:del) as tracked changes. */
export async function importDocx(
  buffer: ArrayBuffer | Uint8Array,
): Promise<DocxImportResult> {
  const zip = await JSZip.loadAsync(buffer);
  const read = (path: string) => zip.file(path)?.async("string");
  const documentXml = await read("word/document.xml");
  if (!documentXml) throw new Error("Not a DOCX file: word/document.xml is missing");

  const theme = parseTheme(await read("word/theme/theme1.xml"));
  const rels = await parseRels(zip);
  const ctx: Context = {
    styles: new StyleSheet(await read("word/styles.xml"), theme),
    settings: parseSettings(await read("word/settings.xml")),
    listLevel: parseNumbering(await read("word/numbering.xml")),
    ids: { next: 1 },
    sections: [],
    sectionIndex: 0,
    comments: parseComments(
      await read("word/comments.xml"),
      await read("word/commentsExtended.xml"),
    ),
    rels,
    images: await loadImages(zip, rels),
    link: null,
    fields: [],
  };

  const root = parseXml(documentXml);
  const body = child(root, "body");
  if (!body) throw new Error("Invalid DOCX: document body is missing");

  const sectPrs = sectionProperties(body);
  for (const sectPr of sectPrs) ctx.sections.push(await parseSection(sectPr, zip, rels, ctx));
  inheritHeaderFooters(ctx.sections, ctx.settings.evenAndOddHeaders);

  const blocks: Block[] = [];
  walkBody(body, blocks, ctx);
  const content = groupLists(blocks);
  if (!content.length) content.push({ type: "paragraph" });

  const last = ctx.sections[ctx.sections.length - 1];
  const bodySectPr = elements(body).some((c) => c.localName === "sectPr");
  const pageSetup = bodySectPr ? last?.layout.page : undefined;
  let layout: DocumentLayout | undefined;
  if (bodySectPr && pageSetup) {
    layout = defined({
      ...last.layout,
      defaultTabStop:
        ctx.settings.defaultTabStop === undefined
          ? undefined
          : ctx.settings.defaultTabStop / TWIPS_PER_PX,
      justifyShrink:
        (ctx.settings.compatibilityMode ?? 0) >= 15 ? JUSTIFY_SHRINK : undefined,
      evenAndOddHeaders: ctx.settings.evenAndOddHeaders || undefined,
      styles: ctx.styles.definitions(ctx.settings),
    }) as DocumentLayout;
  }

  const comments: Record<string, CommentThread> = {};
  for (const [wid, thread] of ctx.comments.threads) {
    if (ctx.comments.anchored.has(wid)) comments[thread.id] = thread;
  }
  const doc = schema.nodeFromJSON({
    type: "doc",
    attrs: { layout: layout ?? null, comments: Object.keys(comments).length ? comments : null },
    content,
  });
  doc.check();
  return { doc: doc.toJSON() as DocJSON, pageSetup, layout };
}
