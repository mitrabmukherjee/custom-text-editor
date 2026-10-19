import {
  Schema,
  type DOMOutputSpec,
  type Mark,
  type MarkSpec,
  type Node as PMNode,
  type NodeSpec,
} from "prosemirror-model";
import { addListNodes } from "prosemirror-schema-list";
import { tableNodes } from "prosemirror-tables";
import {
  cellBorderCss,
  cellMarginCss,
  paragraphCss,
  runCss,
  type CellBorderEdges,
  type CellMarginPx,
  type HFFrame,
  type ParagraphChange,
  type ParagraphFormat,
  type SectionLayout,
  type SectionStart,
} from "./format";
import { HIGHLIGHT_COLORS, displayColor, normalizeHex, textOn } from "./textColors";
import { authorColor } from "./track-changes/colors";

const UNDERLINE_CSS: Record<string, string> = {
  double: "text-decoration-style: double",
  thick: "text-decoration-thickness: 2px",
  dotted: "text-decoration-style: dotted",
  dash: "text-decoration-style: dashed",
  dotDash: "text-decoration-style: dashed",
  wave: "text-decoration-style: wavy",
};

function underlineCss(style: string): string | null {
  return UNDERLINE_CSS[style] ?? null;
}

/** "#rrggbb" for a CSS colour ("rgb(…)", "#hex"); null for transparent or unknown. */
function cssColorHex(value: string): string | null {
  const v = value.trim().toLowerCase();
  const m = v.match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?\s*\)$/);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
    return `#${[m[1], m[2], m[3]].map((c) => Number(c).toString(16).padStart(2, "0")).join("")}`;
  }
  return normalizeHex(v.startsWith("#") ? v : null);
}

/** `href` if it is a web, mail or phone link or an in-document anchor; null for anything else. */
export function safeHref(href: string | null | undefined): string | null {
  const value = href?.trim();
  return value && /^(https?:|mailto:|tel:|#)/i.test(value) ? value : null;
}

const CSS_ALIGN: Record<string, ParagraphFormat["align"]> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  justify: "justify",
};

/** Px from a CSS length in pt, px, in or cm (pasted HTML). */
function cssPx(value: string): number | undefined {
  const m = /^(-?[\d.]+)(pt|px|in|cm)$/.exec(value.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  const px = m[2] === "pt" ? n / 0.75 : m[2] === "in" ? n * 96 : m[2] === "cm" ? (n * 96) / 2.54 : n;
  return Math.round(px * 10) / 10 || undefined;
}

/** Paragraph formatting from inline styles of pasted HTML (Word and web pages). */
function cssFormat(dom: HTMLElement): ParagraphFormat | null {
  const style = dom.style;
  if (!style) return null;
  const format: ParagraphFormat = {};
  const align = CSS_ALIGN[style.textAlign];
  if (align && align !== "left") format.align = align;
  const left = cssPx(style.marginLeft);
  if (left && left > 0) format.indentLeft = left;
  const first = cssPx(style.textIndent);
  if (first) format.indentFirstLine = first;
  return Object.keys(format).length ? format : null;
}

function parseFormat(dom: HTMLElement): ParagraphFormat | null {
  const raw = dom.getAttribute("data-pm-format");
  if (!raw) return cssFormat(dom);
  try {
    return JSON.parse(raw) as ParagraphFormat;
  } catch {
    return null;
  }
}

function parseParaChange(dom: HTMLElement): ParagraphChange | null {
  const raw = dom.getAttribute("data-para-change");
  if (!raw) return null;
  try {
    return JSON.parse(raw) as ParagraphChange;
  } catch {
    return null;
  }
}

function parseFrameAttr(dom: HTMLElement): HFFrame | null {
  const raw = dom.getAttribute("data-frame");
  if (!raw) return null;
  try {
    return JSON.parse(raw) as HFFrame;
  } catch {
    return null;
  }
}

function blockAttrs(dom: HTMLElement) {
  const frame = parseFrameAttr(dom);
  return {
    format: parseFormat(dom),
    paraChange: parseParaChange(dom),
    frame,
    floating: dom.hasAttribute("data-floating"),
  };
}

function blockDataAttrs(node: PMNode): Record<string, string> {
  const format = node.attrs.format as ParagraphFormat | null;
  const change = node.attrs.paraChange as ParagraphChange | null;
  const attrs: Record<string, string> = {};
  if (format) {
    attrs.style = paragraphCss(format);
    attrs["data-pm-format"] = JSON.stringify(format);
  }
  if (change) {
    attrs["data-para-change"] = JSON.stringify(change);
    attrs["data-change-id"] = change.id;
    attrs.class = "tc-para-change";
    attrs.style = `${attrs.style ?? ""}; --tc-color: ${authorColor(change.authorId || "unknown")}`;
  }
  const frame = node.attrs.frame as HFFrame | null;
  if (frame) {
    attrs["data-frame"] = JSON.stringify(frame);
    attrs["data-floating"] = "";
    attrs.class = [attrs.class, "pm-body-frame"].filter(Boolean).join(" ");
  }
  return attrs;
}

function blockDOM(tag: string, node: PMNode): DOMOutputSpec {
  const attrs = blockDataAttrs(node);
  return Object.keys(attrs).length ? [tag, attrs, 0] : [tag, 0];
}

const SECTION_LABELS: Record<SectionStart, string> = {
  nextPage: "Section Break (Next Page)",
  continuous: "Section Break (Continuous)",
  evenPage: "Section Break (Even Page)",
  oddPage: "Section Break (Odd Page)",
};

const changeAttrs = {
  id: { default: "" },
  authorId: { default: "" },
  authorName: { default: "" },
  date: { default: "" },
};

function changeDomAttrs(mark: Mark, verb: string) {
  const { id, authorId, authorName, date } = mark.attrs as Record<string, string>;
  const when = date ? new Date(date).toLocaleString() : "";
  return {
    "data-change-id": id,
    "data-author-id": authorId,
    "data-author-name": authorName,
    "data-date": date,
    title: [authorName || "Unknown", verb, when].filter(Boolean).join(" · "),
    style: `--tc-color: ${authorColor(authorId || "unknown")}`,
  };
}

function parseChangeAttrs(dom: HTMLElement | string) {
  if (typeof dom === "string") return false;
  return {
    id: dom.getAttribute("data-change-id") ?? "",
    authorId: dom.getAttribute("data-author-id") ?? "",
    authorName: dom.getAttribute("data-author-name") ?? "",
    date: dom.getAttribute("data-date") ?? "",
  };
}

const cellNodes = tableNodes({
  tableGroup: "block",
  cellContent: "block+",
  cellAttributes: {
    /** Word cell shading, "#rrggbb". */
    background: {
      default: null,
      getFromDOM: (dom) => normalizeHex(dom.getAttribute("data-background")),
      setDOMAttr: (value, attrs) => {
        if (typeof value !== "string") return;
        attrs["data-background"] = value;
        attrs.style = `${attrs.style ?? ""}background-color: ${value}; color: ${textOn(value)};`;
      },
    },
    valign: {
      default: null,
      getFromDOM: (dom) => dom.getAttribute("data-valign"),
      setDOMAttr: (value, attrs) => {
        if (typeof value !== "string") return;
        attrs["data-valign"] = value;
        attrs.style = `${attrs.style ?? ""}vertical-align: ${value};`;
      },
    },
    edges: {
      default: null,
      getFromDOM: (dom) => {
        const raw = dom.getAttribute("data-cell-edges");
        if (!raw) return null;
        try {
          return JSON.parse(raw) as CellBorderEdges;
        } catch {
          return null;
        }
      },
      setDOMAttr: (value, attrs) => {
        if (!value || typeof value !== "object") return;
        attrs["data-cell-edges"] = "";
        attrs.style = `${attrs.style ?? ""}${cellBorderCss(value as CellBorderEdges)};`;
      },
    },
    cellMargin: {
      default: null,
      getFromDOM: (dom) => {
        const raw = dom.getAttribute("data-cell-margin");
        if (!raw) return null;
        try {
          return JSON.parse(raw) as CellMarginPx;
        } catch {
          return null;
        }
      },
      setDOMAttr: (value, attrs) => {
        if (!value || typeof value !== "object") return;
        attrs["data-cell-margin"] = JSON.stringify(value);
        const marginCss = cellMarginCss(value as CellMarginPx);
        if (marginCss) attrs.style = `${attrs.style ?? ""}${marginCss};`;
      },
    },
  },
});

function tableDomAttrs(node: PMNode): Record<string, string> {
  const attrs: Record<string, string> = { class: "pm-table" };
  if (!node.attrs.borders) attrs["data-borderless"] = "";
  const style: string[] = [];
  const width = node.attrs.width as number | null;
  const widthPct = node.attrs.widthPct as number | null;
  const indent = node.attrs.indent as number | null;
  const align = node.attrs.align as "left" | "center" | "right" | null;
  if (width) {
    style.push(`width: ${Math.round(width)}px`, "max-width: 100%");
    attrs["data-width"] = String(Math.round(width));
  } else if (widthPct) {
    style.push(`width: ${widthPct}%`);
    attrs["data-width-pct"] = String(widthPct);
  }
  if (indent) {
    style.push(`margin-left: ${Math.round(indent)}px`);
    attrs["data-indent"] = String(Math.round(indent));
  }
  if (align === "center") style.push("margin-left: auto", "margin-right: auto");
  else if (align === "right") style.push("margin-left: auto", "margin-right: 0");
  if (style.length) attrs.style = style.join("; ");
  return attrs;
}

function tableRowDom(node: PMNode): DOMOutputSpec {
  const style: string[] = [];
  const h = node.attrs.rowHeight as number | null;
  const attrs: Record<string, string> = {};
  if (h) {
    style.push(`height: ${Math.round(h)}px`, `min-height: ${Math.round(h)}px`);
    attrs["data-row-height"] = String(Math.round(h));
  }
  if (style.length) attrs.style = style.join("; ");
  if (node.attrs.cantSplit) attrs["data-cant-split"] = "";
  const rule = node.attrs.rowHeightRule as string | null;
  if (rule) attrs["data-row-height-rule"] = rule;
  return Object.keys(attrs).length ? ["tr", attrs, 0] : ["tr", 0];
}

/** Word tables: `borders` false for layout tables without gridlines. */
const tables: Record<string, NodeSpec> = {
  ...cellNodes,
  table_row: {
    ...cellNodes.table_row,
    attrs: {
      rowHeight: { default: null },
      rowHeightRule: { default: null },
      cantSplit: { default: false },
    },
    parseDOM: [
      {
        tag: "tr",
        getAttrs: (dom) => ({
          rowHeight: dom.getAttribute("data-row-height")
            ? Number(dom.getAttribute("data-row-height"))
            : null,
          rowHeightRule: dom.getAttribute("data-row-height-rule"),
          cantSplit: dom.hasAttribute("data-cant-split"),
        }),
      },
    ],
    toDOM: tableRowDom,
  },
  table: {
    ...cellNodes.table,
    attrs: {
      borders: { default: true },
      width: { default: null },
      widthPct: { default: null },
      indent: { default: null },
      align: { default: null },
    },
    parseDOM: [
      {
        tag: "table",
        getAttrs: (dom) => ({
          borders: !dom.hasAttribute("data-borderless"),
          width: dom.getAttribute("data-width") ? Number(dom.getAttribute("data-width")) : null,
          widthPct: dom.getAttribute("data-width-pct") ? Number(dom.getAttribute("data-width-pct")) : null,
          indent: dom.getAttribute("data-indent") ? Number(dom.getAttribute("data-indent")) : null,
          align: (dom.getAttribute("data-align") as "left" | "center" | "right" | null) ?? null,
        }),
      },
    ],
    toDOM: (node) => {
      const attrs = tableDomAttrs(node);
      if (node.attrs.align) attrs["data-align"] = node.attrs.align as string;
      return ["table", attrs, ["tbody", 0]];
    },
  },
};

const baseNodes: Record<string, NodeSpec> = {
  doc: {
    content: "block+",
    /**
     * `layout`: `DocumentLayout` imported from Word (page, header/footer, line numbers).
     * `lineNumbers`: line numbering for a document that has no Word layout (the last section).
     * `styles`: `StyleDefinition`s added or updated in the editor (override `layout.styles`).
     * `comments`: comment threads by id (`CommentThread`), anchored by `comment` marks.
     */
    attrs: {
      layout: { default: null },
      lineNumbers: { default: null },
      styles: { default: null },
      comments: { default: null },
    },
  },
  paragraph: {
    /** `paraChange`: a tracked paragraph formatting revision (`ParagraphChange`). */
    attrs: {
      format: { default: null },
      paraChange: { default: null },
      frame: { default: null },
      floating: { default: false },
    },
    content: "inline*",
    group: "block",
    parseDOM: [{ tag: "p", getAttrs: (dom) => blockAttrs(dom) }],
    toDOM: (node) => blockDOM("p", node),
  },
  heading: {
    attrs: { level: { default: 1 }, format: { default: null }, paraChange: { default: null } },
    content: "inline*",
    group: "block",
    defining: true,
    parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({
      tag: `h${level}`,
      getAttrs: (dom: HTMLElement) => ({ level, ...blockAttrs(dom) }),
    })),
    toDOM: (node) => blockDOM(`h${node.attrs.level}`, node),
  },
  page_break: {
    group: "block",
    atom: true,
    selectable: true,
    parseDOM: [{ tag: "hr[data-page-break]" }],
    toDOM: () => ["hr", { "data-page-break": "", class: "pm-page-break-node" }],
  },
  /**
   * End of a Word section. `section` is the layout of the section that ends here;
   * `start` is how the next section starts.
   */
  speaker_line: {
    attrs: { format: { default: null }, label: { default: "THE COURT:" } },
    content: "inline*",    
    group: "block",
    parseDOM: [
      {
        tag: "p[data-speaker-line]",
        getAttrs: (dom) => ({
          ...blockAttrs(dom),
          label: dom.getAttribute("data-label") ?? "THE COURT:",
        }),
      },
    ],
    toDOM: (node) => [
      "p",
      {
        "data-speaker-line": "",
        "data-label": node.attrs.label,
        class: "pm-speaker-line",
        ...blockDataAttrs(node),
      },
      ["span", { class: "pm-speaker-label" }, `${node.attrs.label} `],
      0,
    ],
  },
  qa_line: {
    attrs: { format: { default: null }, role: { default: "q" } },
    content: "inline*",
    group: "block",
    parseDOM: [
      {
        tag: "p[data-qa-line]",
        getAttrs: (dom) => ({
          ...blockAttrs(dom),
          role: dom.getAttribute("data-role") === "a" ? "a" : "q",
        }),
      },
    ],
    toDOM: (node) => {
      const prefix = node.attrs.role === "a" ? "A. " : "Q. ";
      return [
        "p",
        {
          "data-qa-line": "",
          "data-role": node.attrs.role,
          class: "pm-qa-line",
          ...blockDataAttrs(node),
        },
        ["span", { class: "pm-qa-prefix" }, prefix],
        0,
      ];
    },
  },
  timestamp: {
    inline: true,
    group: "inline",
    atom: true,
    attrs: { time: { default: "00:00:00" } },
    parseDOM: [{ tag: "span[data-timestamp]", getAttrs: (dom) => ({ time: dom.getAttribute("data-time") ?? "" }) }],
    toDOM: (node) => ["span", { "data-timestamp": "", class: "pm-timestamp", "data-time": node.attrs.time }, node.attrs.time],
  },
  section_break: {
    group: "block",
    atom: true,
    selectable: true,
    attrs: { start: { default: "nextPage" }, section: { default: null } },
    parseDOM: [
      {
        tag: "hr[data-section-break]",
        getAttrs: (dom) => {
          const raw = dom.getAttribute("data-section");
          let section: SectionLayout | null = null;
          try {
            section = raw ? (JSON.parse(raw) as SectionLayout) : null;
          } catch {
            section = null;
          }
          return { start: dom.getAttribute("data-section-break") || "nextPage", section };
        },
      },
    ],
    toDOM: (node) => [
      "hr",
      {
        "data-section-break": node.attrs.start,
        "data-label": SECTION_LABELS[node.attrs.start as SectionStart] ?? SECTION_LABELS.nextPage,
        "data-section": node.attrs.section ? JSON.stringify(node.attrs.section) : "",
        class: "pm-section-break-node",
      },
    ],
  },
  text: { group: "inline" },
  tab: {
    inline: true,
    group: "inline",
    leafText: () => "\t",
    parseDOM: [{ tag: "span.pm-tab" }],
    toDOM: () => ["span", { class: "pm-tab" }],
  },
  hard_break: {
    inline: true,
    group: "inline",
    selectable: false,
    parseDOM: [{ tag: "br" }],
    toDOM: () => ["br"],
  },
  /** Inline picture; `src` is a data URL for imported and inserted images. Size in px. */
  image: {
    inline: true,
    group: "inline",
    draggable: true,
    attrs: {
      src: {},
      alt: { default: "" },
      title: { default: null },
      width: { default: null },
      height: { default: null },
      /** VML absolute placement within the paragraph line (px). */
      anchorX: { default: null },
      anchorY: { default: null },
      /** Round-trip VML signature metadata (`VmlAnchorJSON`). */
      vmlAnchor: { default: null },
    },
    parseDOM: [
      {
        tag: "img[src]",
        getAttrs: (dom) => ({
          src: dom.getAttribute("src"),
          alt: dom.getAttribute("alt") ?? "",
          title: dom.getAttribute("title"),
          width: Number(dom.getAttribute("width")) || null,
          height: Number(dom.getAttribute("height")) || null,
          anchorX: Number(dom.getAttribute("data-anchor-x")) || null,
          anchorY: Number(dom.getAttribute("data-anchor-y")) || null,
          vmlAnchor: (() => {
            const raw = dom.getAttribute("data-vml-anchor");
            if (!raw) return null;
            try {
              return JSON.parse(raw);
            } catch {
              return null;
            }
          })(),
        }),
      },
    ],
    toDOM: (node) => {
      const { src, alt, title, width, height, anchorX, anchorY, vmlAnchor } = node.attrs;
      const img: Record<string, string | number | null> = {
        class: "pm-image",
        src,
        alt,
        title,
        width,
        height,
      };
      if (anchorX != null) img["data-anchor-x"] = anchorX;
      if (anchorY != null) img["data-anchor-y"] = anchorY;
      if (vmlAnchor) img["data-vml-anchor"] = JSON.stringify(vmlAnchor);
      return ["img", img];
    },
  },
  ...tables,
};

// Track-change marks come first so they wrap formatting marks in the DOM.
const marks: Record<string, MarkSpec> = {
  insertion: {
    attrs: changeAttrs,
    inclusive: false,
    parseDOM: [{ tag: "ins[data-change-id]", getAttrs: parseChangeAttrs }],
    toDOM: (mark) => ["ins", { class: "tc-insertion", ...changeDomAttrs(mark, "inserted") }, 0],
  },
  deletion: {
    attrs: changeAttrs,
    inclusive: false,
    parseDOM: [{ tag: "del[data-change-id]", getAttrs: parseChangeAttrs }],
    toDOM: (mark) => ["del", { class: "tc-deletion", ...changeDomAttrs(mark, "deleted") }, 0],
  },
  /** A formatting revision; `old` is the `RunFormat` before the change. */
  formatChange: {
    attrs: { ...changeAttrs, description: { default: "" }, old: { default: null } },
    inclusive: false,
    parseDOM: [
      {
        tag: "span[data-format-change]",
        getAttrs: (dom) => {
          const base = parseChangeAttrs(dom);
          if (!base) return false;
          const old = dom.getAttribute("data-old");
          return {
            ...base,
            description: dom.getAttribute("data-format-change") ?? "",
            old: old ? JSON.parse(old) : null,
          };
        },
      },
    ],
    toDOM: (mark) => [
      "span",
      {
        class: "tc-format",
        ...changeDomAttrs(mark, `formatted (${mark.attrs.description})`),
        "data-format-change": mark.attrs.description,
        "data-old": mark.attrs.old ? JSON.stringify(mark.attrs.old) : "",
      },
      0,
    ],
  },
  /** Text a comment is about; the thread lives in the doc's `comments` attribute. Comments may overlap. */
  comment: {
    attrs: { id: {}, authorId: { default: "" } },
    inclusive: false,
    excludes: "",
    parseDOM: [
      {
        tag: "span[data-comment-id]",
        getAttrs: (dom) => ({
          id: dom.getAttribute("data-comment-id"),
          authorId: dom.getAttribute("data-comment-author") ?? "",
        }),
      },
    ],
    toDOM: (mark) => [
      "span",
      {
        class: "pm-comment",
        "data-comment-id": mark.attrs.id,
        "data-comment-author": mark.attrs.authorId,
        style: `--tc-color: ${authorColor(mark.attrs.authorId || "unknown")}`,
      },
      0,
    ],
  },
  /** Hyperlink; `href` is kept as imported, but only safe schemes are rendered or followed. */
  link: {
    attrs: { href: {}, title: { default: null } },
    inclusive: false,
    parseDOM: [
      {
        tag: "a[data-href], a[href]",
        getAttrs: (dom) => ({
          href: dom.getAttribute("data-href") ?? dom.getAttribute("href"),
          title: dom.getAttribute("data-title"),
        }),
      },
    ],
    toDOM: (mark) => [
      "a",
      {
        class: "pm-link",
        href: safeHref(mark.attrs.href),
        "data-href": mark.attrs.href,
        "data-title": mark.attrs.title,
        title: `${mark.attrs.title ?? mark.attrs.href}\nCtrl+Click to follow link`,
        rel: "noopener noreferrer",
      },
      0,
    ],
  },
  textStyle: {
    attrs: {
      fontFamily: { default: null },
      fontSize: { default: null },
      fontKern: { default: null },
    },
    parseDOM: [
      {
        tag: "span[data-text-style]",
        getAttrs: (dom) => ({
          fontFamily: dom.getAttribute("data-font-family") || null,
          fontSize: Number(dom.getAttribute("data-font-size")) || null,
          fontKern: dom.getAttribute("data-font-kern") === "1" ? true : null,
        }),
      },
    ],
    toDOM: (mark) => [
      "span",
      {
        "data-text-style": "",
        "data-font-family": mark.attrs.fontFamily ?? "",
        "data-font-size": mark.attrs.fontSize ?? "",
        "data-font-kern": mark.attrs.fontKern ? "1" : "",
        style: runCss(mark.attrs),
      },
      0,
    ],
  },
  strong: {
    parseDOM: [
      { tag: "strong" },
      { tag: "b" },
      { style: "font-weight=bold" },
      { style: "font-weight=700" },
    ],
    toDOM: () => ["strong", 0],
  },
  em: {
    parseDOM: [{ tag: "i" }, { tag: "em" }, { style: "font-style=italic" }],
    toDOM: () => ["em", 0],
  },
  underline: {
    attrs: { style: { default: "single" } },
    parseDOM: [
      { tag: "u[data-underline]", getAttrs: (dom) => ({ style: dom.getAttribute("data-underline") }) },
      { tag: "u" },
      { tag: "ins:not([data-change-id])" },
      {
        style: "text-decoration",
        getAttrs: (value) => (/underline/.test(value) ? { style: "single" } : false),
      },
      {
        style: "text-decoration-line",
        getAttrs: (value) => (/underline/.test(value) ? { style: "single" } : false),
      },
    ],
    toDOM: (mark) => [
      "u",
      { class: "pm-underline", "data-underline": mark.attrs.style, style: underlineCss(mark.attrs.style) },
      0,
    ],
  },
  strike: {
    attrs: { style: { default: "single" } },
    parseDOM: [
      { tag: "s[data-strike]", getAttrs: (dom) => ({ style: dom.getAttribute("data-strike") }) },
      { tag: "s" },
      { tag: "strike" },
      { tag: "del:not([data-change-id])" },
      {
        style: "text-decoration",
        getAttrs: (value) => (/line-through/.test(value) ? { style: "single" } : false),
      },
      {
        style: "text-decoration-line",
        getAttrs: (value) => (/line-through/.test(value) ? { style: "single" } : false),
      },
    ],
    toDOM: (mark) => [
      "s",
      {
        class: "pm-strike",
        "data-strike": mark.attrs.style,
        style: mark.attrs.style === "double" ? "text-decoration-style: double" : null,
      },
      0,
    ],
  },
  highlight: {
    attrs: { color: { default: HIGHLIGHT_COLORS.yellow } },
    parseDOM: [
      {
        tag: "span[data-highlight]",
        getAttrs: (dom) => {
          const color = normalizeHex(dom.getAttribute("data-highlight"));
          return color ? { color } : false;
        },
      },
      {
        tag: "mark",
        getAttrs: (dom) => ({
          color: cssColorHex(dom.style.backgroundColor) ?? HIGHLIGHT_COLORS.yellow,
        }),
      },
      {
        style: "background-color",
        getAttrs: (value) => {
          const color = cssColorHex(value);
          return color && color !== "#ffffff" ? { color } : false;
        },
      },
    ],
    toDOM: (mark) => [
      "span",
      {
        class: "pm-highlight",
        "data-highlight": mark.attrs.color,
        style: `--pm-highlight: ${mark.attrs.color}; --pm-highlight-text: ${textOn(mark.attrs.color)}`,
      },
      0,
    ],
  },
  textColor: {
    attrs: { color: { default: "#000000" } },
    parseDOM: [
      {
        tag: "span[data-color]",
        getAttrs: (dom) => {
          const color = normalizeHex(dom.getAttribute("data-color"));
          return color ? { color } : false;
        },
      },
      {
        tag: "font[color]",
        getAttrs: (dom) => {
          const color = cssColorHex(dom.getAttribute("color") ?? "");
          return color ? { color } : false;
        },
      },
      {
        style: "color",
        getAttrs: (value) => {
          const color = cssColorHex(value);
          return color ? { color } : false;
        },
      },
    ],
    toDOM: (mark) => [
      "span",
      {
        class: "pm-color",
        "data-color": mark.attrs.color,
        style: `--pm-color: ${displayColor(mark.attrs.color)}; --pm-print-color: ${mark.attrs.color}`,
      },
      0,
    ],
  },
  caps: {
    attrs: { style: { default: "all" } },
    parseDOM: [
      { tag: "span[data-caps]", getAttrs: (dom) => ({ style: dom.getAttribute("data-caps") }) },
      { style: "font-variant", getAttrs: (v) => (/small-caps/.test(v) ? { style: "small" } : false) },
      { style: "text-transform", getAttrs: (v) => (v === "uppercase" ? { style: "all" } : false) },
    ],
    toDOM: (mark) => [
      "span",
      { class: `pm-caps pm-caps--${mark.attrs.style}`, "data-caps": mark.attrs.style },
      0,
    ],
  },
  vertAlign: {
    attrs: { align: { default: "superscript" } },
    excludes: "vertAlign",
    parseDOM: [
      { tag: "sup", attrs: { align: "superscript" } },
      { tag: "sub", attrs: { align: "subscript" } },
      { style: "vertical-align=super", attrs: { align: "superscript" } },
      { style: "vertical-align=sub", attrs: { align: "subscript" } },
    ],
    toDOM: (mark) => [mark.attrs.align === "subscript" ? "sub" : "sup", { class: "pm-vert" }, 0],
  },
};

const baseSchema = new Schema({ nodes: baseNodes, marks });

/** Number formats of an ordered list level; "legal" is outline numbering (1.1.1.). */
export type NumFormat = "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" | "legal";

export type BulletStyle = "disc" | "circle" | "square" | "dash" | "arrow" | "check" | "diamond";

const LIST_STYLE_TYPE: Record<Exclude<NumFormat, "legal">, string> = {
  decimal: "decimal",
  lowerLetter: "lower-alpha",
  upperLetter: "upper-alpha",
  lowerRoman: "lower-roman",
  upperRoman: "upper-roman",
};

export const BULLET_CHARS: Record<BulletStyle, string> = {
  disc: "\u2022",
  circle: "o",
  square: "\u25aa",
  dash: "\u2013",
  arrow: "\u27a2",
  check: "\u2713",
  diamond: "\u2756",
};

const listNodes = addListNodes(baseSchema.spec.nodes, "paragraph block*", "block")
  .update("ordered_list", {
    /** `numFormat` null: Word's default for the nesting depth (1. a. i.), see editor.css. */
    attrs: { order: { default: 1 }, numFormat: { default: null } },
    content: "list_item+",
    group: "block",
    parseDOM: [
      {
        tag: "ol",
        getAttrs: (dom: HTMLElement) => ({
          order: dom.hasAttribute("start") ? Number(dom.getAttribute("start")) || 1 : 1,
          numFormat: dom.getAttribute("data-num-format") || null,
        }),
      },
    ],
    toDOM: (node) => {
      const format = node.attrs.numFormat as NumFormat | null;
      const order = node.attrs.order as number;
      const attrs: Record<string, string> = {};
      if (order !== 1) attrs.start = String(order);
      if (format) {
        attrs["data-num-format"] = format;
        attrs.style =
          format === "legal"
            ? `counter-reset: pm-legal ${order - 1}`
            : `list-style-type: ${LIST_STYLE_TYPE[format]}`;
      }
      return ["ol", attrs, 0];
    },
  })
  .update("bullet_list", {
    attrs: { bullet: { default: null } },
    content: "list_item+",
    group: "block",
    parseDOM: [
      {
        tag: "ul",
        getAttrs: (dom: HTMLElement) => ({ bullet: dom.getAttribute("data-bullet") || null }),
      },
    ],
    toDOM: (node) => {
      const bullet = node.attrs.bullet as BulletStyle | null;
      if (!bullet) return ["ul", 0];
      const type = bullet === "disc" || bullet === "circle" || bullet === "square"
        ? bullet
        : `"${BULLET_CHARS[bullet]}  "`;
      return ["ul", { "data-bullet": bullet, style: `list-style-type: ${type}` }, 0];
    },
  });

export const schema = new Schema({
  nodes: listNodes,
  marks: baseSchema.spec.marks,
});

export type TranscriptSchema = typeof schema;
