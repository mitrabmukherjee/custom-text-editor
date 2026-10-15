import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import {
  type PageFrame,
  type PageInfo,
  type ParagraphFormat,
  type SectionStart,
  paragraphLineHeightPx,
} from "../format";
import { collectComments } from "../comments";
import { collectChanges, type ChangeType } from "../track-changes/changes";
import { SKIP_TRACKING } from "../track-changes/plugin";
import {
  computePageBreaks,
  type BreakKind,
  type LayoutLine,
  type LinePlacement,
} from "./computePageBreaks";
import { stackHeight, type PageSetup } from "./pageSetup";
import { columnCount, tableElementAt, trimTableColumnDom } from "../tableColumnDom";

export type PageBreakSpec = {
  pos: number;
  height: number;
  /** Page number that ends at this break. */
  page: number;
  forced: boolean;
};

export type TabWidth = { pos: number; width: number };

/** Word's Show Markup options: which revisions and comments are drawn. */
export type MarkupFilter = { insertions: boolean; formatting: boolean; comments: boolean };

export const ALL_MARKUP: MarkupFilter = { insertions: true, formatting: true, comments: true };

/** One tracked change or comment shown in the markup pane (Word balloon). */
export type ChangeBalloon = {
  id: string;
  changeType: ChangeType | "comment";
  /** 0-based, like `LinePlacement.page`. */
  page: number;
  top: number;
  bottom: number;
  /** px from the page's left edge to the start of the changed text. */
  left: number;
  authorId: string;
  authorName: string;
  /** Body line under the author, e.g. "Formatted: …" (Word All Markup wording). */
  label: string;
};

export type PaginationLayout = {
  pageCount: number;
  /** Section and displayed page number of every page. */
  pages: PageInfo[];
  /** Every numbered body line's final page and position (px from the page's top edge). */
  lines: LinePlacement[];
  /** Lines containing tracked changes, for Word's change bars. */
  changedLines: LinePlacement[];
  balloons: ChangeBalloon[];
};

type PaginationState = {
  breaks: PageBreakSpec[];
  tabs: TabWidth[];
  pageCount: number;
  decorations: DecorationSet;
};

type PaginationMeta = Partial<Pick<PaginationState, "breaks" | "tabs" | "pageCount">>;

export type PaginationConfig = {
  getSetup: () => PageSetup;
  getFrame: (page: PageInfo) => PageFrame;
  getDefaultTabStop: () => number;
  /** `pageNumberStart` of a section, if it restarts page numbering. */
  getNumberStart?: (section: number) => number | undefined;
  getMarkupFilter?: () => MarkupFilter;
  onLayout?: (layout: PaginationLayout) => void;
};

export const paginationKey = new PluginKey<PaginationState>("pagination");

const WIDGET_CLASS = "pm-page-break-widget";
const TAB_SELECTOR = "span.pm-tab";

function changeSelector(filter: MarkupFilter) {
  return [
    ...(filter.insertions ? ["ins.tc-insertion", "del.tc-deletion"] : []),
    ...(filter.formatting ? ["span.tc-format", ".tc-para-change"] : []),
  ].join(", ");
}

function showsChange(filter: MarkupFilter, type: ChangeType) {
  return type === "insertion" || type === "deletion" ? filter.insertions : filter.formatting;
}

const SECTION_BREAKS: Record<SectionStart, BreakKind> = {
  nextPage: true,
  continuous: false,
  evenPage: "even",
  oddPage: "odd",
};

/** A break between table rows has to be a row itself to be valid inside `<tbody>`. */
function rowSpacer(height: number, colspan: number) {
  const row = document.createElement("tr");
  const cell = row.appendChild(document.createElement("td"));
  cell.colSpan = Math.max(1, colspan);
  cell.className = "pm-page-break-cell";
  cell.style.height = `${height}px`;
  return row;
}

function buildDecorations(doc: PMNode, breaks: PageBreakSpec[], tabs: TabWidth[]) {
  const decos: Decoration[] = breaks.map((b) =>
    Decoration.widget(
      b.pos,
      () => {
        const $pos = doc.resolve(b.pos);
        const inTable = $pos.parent.type.spec.tableRole === "table";
        const colspan = inTable ? TableMap.get($pos.parent).width : 1;
        const el = inTable ? rowSpacer(b.height, colspan) : document.createElement("span");
        el.classList.add(WIDGET_CLASS);
        el.contentEditable = "false";
        if (!inTable) el.style.height = `${b.height}px`;
        el.dataset.page = String(b.page);
        if (b.forced) el.dataset.forced = "";
        return el;
      },
      {
        side: -1,
        ignoreSelection: true,
        marks: [],
        key: `pb:${b.page}:${b.pos}:${Math.round(b.height * 2)}`,
      },
    ),
  );
  for (const t of tabs) {
    const node = doc.nodeAt(t.pos);
    if (node?.type.name !== "tab") continue;
    decos.push(Decoration.node(t.pos, t.pos + 1, { style: `width: ${t.width}px` }));
  }
  return DecorationSet.create(doc, decos);
}

export function getPageCount(state: Parameters<typeof paginationKey.getState>[0]) {
  return paginationKey.getState(state)?.pageCount ?? 1;
}

export function paginationPlugin(config: PaginationConfig) {
  return new Plugin<PaginationState>({
    key: paginationKey,
    state: {
      init: () => ({ breaks: [], tabs: [], pageCount: 1, decorations: DecorationSet.empty }),
      apply(tr, prev, _old, newState) {
        const meta = tr.getMeta(paginationKey) as PaginationMeta | undefined;
        if (meta) {
          const next = { ...prev, ...meta };
          if (tr.docChanged && !meta.tabs) {
            next.tabs = prev.tabs.map((t) => ({ ...t, pos: tr.mapping.map(t.pos, 1) }));
          }
          return { ...next, decorations: buildDecorations(newState.doc, next.breaks, next.tabs) };
        }
        if (!tr.docChanged) return prev;
        return {
          ...prev,
          breaks: prev.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1) })),
          tabs: prev.tabs.map((t) => ({ ...t, pos: tr.mapping.map(t.pos, 1) })),
          decorations: prev.decorations.map(tr.mapping, tr.doc),
        };
      },
    },
    props: {
      decorations: (state) => paginationKey.getState(state)?.decorations,
    },
    view: (view) => new PaginationView(view, config),
  });
}

type LineRef =
  | { kind: "block"; pos: number }
  | { kind: "inline"; pos: number; el: HTMLElement; contentTop: number };

type MeasuredLine = LayoutLine & { ref: LineRef; numbered: boolean };

type Box = { top: number; bottom: number };

function textNodesOf(el: HTMLElement): Text[] {
  const out: Text[] = [];
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!n.textContent) continue;
    if (n.parentElement?.closest(`.${WIDGET_CLASS}`)) continue;
    out.push(n as Text);
  }
  return out;
}

/** Visual lines of a text block, as merged text rects in client coordinates. */
function textLines(el: HTMLElement): Box[] {
  const range = document.createRange();
  const rects: Box[] = [];
  for (const node of textNodesOf(el)) {
    range.selectNodeContents(node);
    for (const r of range.getClientRects()) {
      if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom });
    }
  }
  for (const inline of el.querySelectorAll<HTMLElement>(`${TAB_SELECTOR}, img`)) {
    // Anchored VML signatures paint above the line; measuring them splits the paragraph
    // into extra lines and page breaks shift when zoom changes layout/pagination.
    if (inline.closest(".pm-image-anchored")) continue;
    const r = inline.getBoundingClientRect();
    if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom });
  }
  rects.sort((a, b) => a.top - b.top);
  const lines: Box[] = [];
  for (const r of rects) {
    const last = lines[lines.length - 1];
    if (last && r.top < last.bottom - 2) {
      last.top = Math.min(last.top, r.top);
      last.bottom = Math.max(last.bottom, r.bottom);
    } else {
      lines.push({ ...r });
    }
  }
  return lines;
}

/** Content box of an element in client px (border box minus padding and border). */
function contentBox(el: HTMLElement, scale: number): Box {
  const r = el.getBoundingClientRect();
  const cs = getComputedStyle(el);
  return {
    top: r.top + (parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth)) * scale,
    bottom: r.bottom - (parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth)) * scale,
  };
}

/** Client px per CSS px, so a host page's CSS zoom or scale transform doesn't skew layout. */
function clientScale(dom: HTMLElement, rect: DOMRect): number {
  const scale = dom.offsetWidth ? rect.width / dom.offsetWidth : 1;
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** Document position of the first character rendered on the line at `contentTop`. */
function lineStartPos(view: EditorView, el: HTMLElement, contentTop: number): number | null {
  const range = document.createRange();
  const charTop = (node: Text, i: number) => {
    range.setStart(node, i);
    range.setEnd(node, i + 1);
    const rects = range.getClientRects();
    return rects.length ? rects[rects.length - 1].top : -Infinity;
  };
  const threshold = contentTop - 2;
  for (const node of textNodesOf(el)) {
    const len = node.length;
    if (charTop(node, len - 1) < threshold) continue;
    let lo = 0;
    let hi = len - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (charTop(node, mid) >= threshold) hi = mid;
      else lo = mid + 1;
    }
    return view.posAtDOM(node, lo);
  }
  return null;
}

/** Next Word tab stop after `x` (px from the paragraph's left edge). */
export function nextTabStop(x: number, format: ParagraphFormat | null, defaultTab: number) {
  const stops = (format?.tabs ?? []).map((t) => t.pos);
  if (format?.indentLeft && (format.indentFirstLine ?? 0) < 0) stops.push(format.indentLeft);
  const custom = stops.filter((p) => p > x + 0.5).sort((a, b) => a - b)[0];
  if (custom !== undefined) return custom;
  const step = defaultTab > 0 ? defaultTab : 48;
  return (Math.floor((x + 0.5) / step) + 1) * step;
}

/** Index of the last line starting at or above `y`, or -1. */
function lineAt(lines: LayoutLine[], y: number) {
  let lo = 0;
  let hi = lines.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (lines[mid].top <= y) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}

function balloonLabel(type: ChangeBalloon["changeType"], text: string, description?: string) {
  const snippet = (t: string) => {
    const s = t.replace(/\s+/g, " ").trim();
    return s.length > 72 ? `${s.slice(0, 72)}\u2026` : s || "(whitespace)";
  };
  if (type === "format" || type === "paragraph") return `Formatted: ${description ?? ""}`;
  return `Formatted: ${snippet(text)}`;
}

/** Change bars and balloons for every tracked change, placed with the lines' final page positions. */
function markupLayout(
  view: EditorView,
  dom: HTMLElement,
  lines: LayoutLine[],
  placements: LinePlacement[],
  natural: (y: number) => number,
  originX: number,
  scale: number,
  filter: MarkupFilter,
): Pick<PaginationLayout, "changedLines" | "balloons"> {
  const rects: Box[] = [];
  const selector = changeSelector(filter);
  for (const el of selector ? dom.querySelectorAll<HTMLElement>(selector) : []) {
    for (const r of el.getClientRects()) {
      if (r.height > 0) rects.push({ top: natural(r.top), bottom: natural(r.bottom) });
    }
  }

  rects.sort((a, b) => a.top - b.top);
  const changedLines: LinePlacement[] = [];
  let k = 0;
  lines.forEach((line, i) => {
    while (k < rects.length && rects[k].bottom <= line.top + 1) k++;
    if (k < rects.length && rects[k].top < line.bottom - 1) changedLines.push(placements[i]);
  });

  const balloons: ChangeBalloon[] = [];
  const place = (
    item: { id: string; from: number; authorId: string; authorName: string },
    changeType: ChangeBalloon["changeType"],
    label: string,
    selector: string,
  ) => {
    let coords: { top: number; bottom: number; left: number } | null = null;
    try {
      const c = view.coordsAtPos(item.from, 1);
      coords = { top: c.top, bottom: c.bottom, left: c.left };
    } catch {
      const r = dom.querySelector<HTMLElement>(selector)?.getClientRects()[0];
      if (r) coords = { top: r.top, bottom: r.bottom, left: r.left };
    }
    if (!coords) return;
    const i = lineAt(lines, natural((coords.top + coords.bottom) / 2));
    if (i < 0) return;
    const { page, top, bottom } = placements[i];
    balloons.push({
      id: item.id,
      changeType,
      page,
      top,
      bottom,
      left: (coords.left - originX) / scale,
      authorId: item.authorId,
      authorName: item.authorName,
      label,
    });
  };
  for (const change of collectChanges(view.state.doc)) {
    if (!showsChange(filter, change.type)) continue;
    place(
      change,
      change.type,
      balloonLabel(change.type, change.text, change.description),
      `[data-change-id="${CSS.escape(change.id)}"]`,
    );
  }
  if (filter.comments) {
    for (const comment of collectComments(view.state.doc)) {
      const replies = (comment.replies ?? []).map((r) => `${r.authorName}: ${r.text}`);
      place(
        comment,
        "comment",
        [comment.text, ...replies].join("\n"),
        `[data-comment-id="${CSS.escape(comment.id)}"]`,
      );
    }
    balloons.sort((a, b) => a.page - b.page || a.top - b.top || a.left - b.left);
  }
  return { changedLines, balloons };
}

function sameBreaks(a: PageBreakSpec[], b: PageBreakSpec[]) {
  return (
    a.length === b.length &&
    a.every(
      (x, i) =>
        x.pos === b[i].pos &&
        x.forced === b[i].forced &&
        Math.abs(x.height - b[i].height) < 1,
    )
  );
}

function sameTabs(a: TabWidth[], b: TabWidth[]) {
  return (
    a.length === b.length &&
    a.every((x, i) => x.pos === b[i].pos && Math.abs(x.width - b[i].width) < 0.5)
  );
}

/**
 * Maximum number of successive tab-width dispatches before we accept the current
 * widths and proceed to page-break computation.  Setting a tab's width can change
 * text wrapping which changes subsequent tabs' x-positions which changes their widths —
 * an oscillation that never converges.  Capping at a small number stops the loop while
 * still giving tabs two or three rounds to settle into a reasonable position.
 */
const MAX_TAB_ROUNDS = 4;

class PaginationView {
  private frame = 0;
  private resizeDebounce = 0;
  private lastSignature = "";
  private tabRounds = 0;
  private lastDocSize = 0;
  private readonly resizeObserver: ResizeObserver;

  constructor(
    private readonly view: EditorView,
    private readonly config: PaginationConfig,
  ) {
    this.resizeObserver = new ResizeObserver(() => {
      if (this.resizeDebounce) window.clearTimeout(this.resizeDebounce);
      this.resizeDebounce = window.setTimeout(() => {
        this.resizeDebounce = 0;
        this.schedule();
      }, 32);
    });
    this.resizeObserver.observe(view.dom);
    void document.fonts?.ready.then(() => this.schedule());
    this.schedule();
  }

  update(view: EditorView, prevState: EditorView["state"]) {
    if (
      view.state.doc !== prevState.doc ||
      paginationKey.getState(view.state) !== paginationKey.getState(prevState)
    ) {
      this.schedule();
    }
  }

  destroy() {
    cancelAnimationFrame(this.frame);
    if (this.resizeDebounce) window.clearTimeout(this.resizeDebounce);
    this.resizeObserver.disconnect();
  }

  /** Re-measure on the next frame (e.g. after a page setup change). */
  schedule() {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.measure();
    });
  }

  private dispatch(meta: PaginationMeta) {
    this.view.dispatch(
      this.view.state.tr
        .setMeta(paginationKey, meta)
        .setMeta("addToHistory", false)
        .setMeta(SKIP_TRACKING, true),
    );
  }

  /** Size every tab so it reaches its Word tab stop. Returns true if widths changed. */
  private layoutTabs(): boolean {
    const { view } = this;
    const defaultTab = this.config.getDefaultTabStop();
    const scale = clientScale(view.dom as HTMLElement, view.dom.getBoundingClientRect());
    const tabs: TabWidth[] = [];
    const formats = new Map<HTMLElement, ParagraphFormat | null>();
    view.state.doc.descendants((node, pos) => {
      if (node.type.name !== "tab") return true;
      const el = view.nodeDOM(pos) as HTMLElement | null;
      const block = el?.closest<HTMLElement>("p, h1, h2, h3, h4, h5, h6");
      if (!el || !block) return false;
      if (!formats.has(block)) {
        const raw = block.getAttribute("data-pm-format");
        formats.set(block, raw ? (JSON.parse(raw) as ParagraphFormat) : null);
      }
      const x = (el.getBoundingClientRect().left - block.getBoundingClientRect().left) / scale;
      const width = Math.max(0, nextTabStop(x, formats.get(block)!, defaultTab) - x);
      tabs.push({ pos, width: Math.round(width * 100) / 100 });
      return false;
    });
    const current = paginationKey.getState(view.state);
    if (!current || sameTabs(current.tabs, tabs)) {
      this.tabRounds = 0;
      return false;
    }
    // If tabs keep oscillating (due to wrapping feedback), accept current widths and move on.
    if (this.tabRounds >= MAX_TAB_ROUNDS) {
      this.tabRounds = 0;
      return false;
    }
    this.tabRounds += 1;
    this.dispatch({ tabs });
    return true;
  }

  private measure() {
    const { view } = this;
    const dom = view.dom as HTMLElement;
    // Reset the tab-round counter whenever the document changes size (new doc loaded).
    const docSize = view.state.doc.content.size;
    if (docSize !== this.lastDocSize) {
      this.lastDocSize = docSize;
      this.tabRounds = 0;
    }
    if (dom.querySelector("tbody col")) {
      view.state.doc.descendants((node, pos) => {
        if (node.type.spec.tableRole !== "table") return true;
        const el = tableElementAt(view, pos);
        if (el) trimTableColumnDom(el, columnCount(node));
        return false;
      });
    }
    const domRect = dom.getBoundingClientRect();
    if (domRect.height === 0) return;
    if (this.layoutTabs()) return;

    const setup = this.config.getSetup();
    const scale = clientScale(dom, domRect);
    const originY = domRect.top;
    const widgets: Box[] = Array.from(
      dom.querySelectorAll<HTMLElement>(`.${WIDGET_CLASS}`),
    ).map((el) => {
      const r = el.getBoundingClientRect();
      return { top: r.top, bottom: r.bottom };
    });
    /** Client y → CSS px from the top of page 1, ignoring page-break spacers. */
    const natural = (y: number) => {
      let above = 0;
      for (const w of widgets) if (w.bottom <= y + 0.5) above += w.bottom - w.top;
      return (y - originY - above) / scale;
    };

    const lines: MeasuredLine[] = [];
    let block = 0;
    let section = 0;
    view.state.doc.descendants((node, pos) => {
      const name = node.type.name;
      if (name === "page_break" || name === "section_break") {
        const breakAfter =
          name === "page_break"
            ? true
            : (SECTION_BREAKS[node.attrs.start as SectionStart] ?? true);
        if (name === "section_break") section += 1;
        if (breakAfter) {
          const el = view.nodeDOM(pos) as HTMLElement | null;
          if (el) {
            const r = el.getBoundingClientRect();
            lines.push({
              top: natural(r.top),
              bottom: natural(r.bottom),
              breakAfter,
              block: block++,
              section: name === "section_break" ? section - 1 : section,
              ref: { kind: "block", pos },
              numbered: false,
            });
          }
        }
        return false;
      }
      if (node.type.spec.tableRole === "table") {
        // Rows don't split across pages; a break before the first row goes before the table.
        node.forEach((row, offset, index) => {
          const rowPos = pos + 1 + offset;
          const el = view.nodeDOM(rowPos) as HTMLElement | null;
          if (!el) return;
          const r = el.getBoundingClientRect();
          let keepNext = false;
          row.descendants((cell) => {
            if (!cell.isTextblock) return;
            const f = cell.attrs.format as ParagraphFormat | null;
            if (f?.keepNext) keepNext = true;
          });
          lines.push({
            top: natural(r.top),
            bottom: natural(r.bottom),
            block: block++,
            section,
            ref: { kind: "block", pos: index === 0 ? pos : rowPos },
            numbered: false,
            keepNext,
          });
        });
        return false;
      }
      if (!node.isTextblock) return true;
      const el = view.nodeDOM(pos) as HTMLElement | null;
      if (!el) return false;

      if (node.type.name === "paragraph" && node.attrs.frame && node.attrs.floating) {
        const wrap = el.classList.contains("pm-body-frame-wrap")
          ? el
          : el.closest<HTMLElement>(".pm-body-frame-wrap");
        const anchor = wrap ?? el;
        const top = natural(anchor.getBoundingClientRect().top);
        lines.push({
          top,
          bottom: top,
          block: block++,
          section,
          numbered: false,
          ref: { kind: "block", pos },
        });
        return false;
      }

      const format = node.attrs.format as ParagraphFormat | null;
      const keep = {
        block: block++,
        section,
        widowControl: !!format?.widowControl,
        keepNext: !!format?.keepNext,
        keepLines: !!format?.keepLines,
        breakBefore: !!format?.pageBreakBefore,
        numbered: !format?.suppressLineNumbers,
      };
      const computedLh = parseFloat(getComputedStyle(el).lineHeight);
      const lineHeight =
        paragraphLineHeightPx(format) ??
        (Number.isFinite(computedLh) ? computedLh : undefined);
      const multiple =
        format?.lineRule === "auto" && (format.lineValue ?? 1) > 1 ? format.lineValue! : 1;
      const slackOf = (height: number) => height - height / multiple;
      const boxes = textLines(el);
      if (!boxes.length) {
        const box = contentBox(el, scale);
        const top = natural(box.top);
        const height = (box.bottom - box.top) / scale;
        lines.push({
          ...keep,
          top,
          bottom: top + height,
          slack: slackOf(height),
          ref: { kind: "block", pos },
        });
        return false;
      }
      boxes.forEach((box, i) => {
        const glyphs = (box.bottom - box.top) / scale;
        const height =
          lineHeight !== undefined ? Math.max(lineHeight, glyphs) : glyphs;
        const top = natural(box.top) - (height - glyphs) / 2;
        lines.push({
          ...keep,
          top,
          bottom: top + height,
          slack: slackOf(height),
          ref:
            i === 0
              ? { kind: "block", pos }
              : { kind: "inline", pos, el, contentTop: box.top },
        });
      });
      return false;
    });

    const layout = computePageBreaks(lines, setup, this.config.getFrame, {
      numberStart: this.config.getNumberStart,
    });
    const breaks: PageBreakSpec[] = [];
    layout.breaks.forEach((b, i) => {
      const ref = lines[b.lineIndex].ref;
      const pos =
        ref.kind === "block"
          ? ref.pos
          : (lineStartPos(view, ref.el, ref.contentTop) ?? ref.pos + 1);
      breaks.push({ pos, height: b.height, page: i + 1, forced: b.forced });
    });

    dom.style.minHeight = `${stackHeight(setup, layout.pageCount)}px`;

    const current = paginationKey.getState(view.state);
    if (
      current &&
      (!sameBreaks(current.breaks, breaks) || current.pageCount !== layout.pageCount)
    ) {
      this.dispatch({ breaks, pageCount: layout.pageCount });
      return;
    }

    const placed = layout.placements.filter((_, i) => lines[i].numbered);
    const markup = markupLayout(
      view,
      dom,
      lines,
      layout.placements,
      natural,
      domRect.left,
      scale,
      this.config.getMarkupFilter?.() ?? ALL_MARKUP,
    );
    const at = (p: LinePlacement) => `${p.page}:${Math.round(p.top)}:${Math.round(p.bottom)}`;
    const signature = [
      layout.pageCount,
      layout.pages.map((p) => `${p.section}:${p.first ? 1 : 0}:${p.number}`).join(","),
      placed.map((p) => `${at(p)}:${p.section ?? 0}`).join(","),
      markup.changedLines.map(at).join(","),
      markup.balloons
        .map(
          (b) =>
            `${b.changeType}:${b.id}:${at(b)}:${Math.round(b.left)}:${b.authorId}:${b.label}`,
        )
        .join(","),
    ].join("|");
    if (signature !== this.lastSignature) {
      this.lastSignature = signature;
      this.config.onLayout?.({
        pageCount: layout.pageCount,
        pages: layout.pages,
        lines: placed,
        ...markup,
      });
    }
  }
}
