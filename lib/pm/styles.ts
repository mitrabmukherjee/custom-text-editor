import type { MarkType, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { SKIP_TRACKING } from "./track-changes";
import { applyCharacterFormat, characterFormatAt, type CharacterPatch } from "./character";
import type { DocumentLayout, ParagraphFormat, StyleDefinition } from "./format";
import { schema } from "./schema";
import type { RunFormat } from "./track-changes/formatting";

const PT = 96 / 72;

/** Word's built-in quick styles, used when the document doesn't define them. */
export const BUILTIN_STYLES: StyleDefinition[] = [
  { id: "Normal", name: "Normal", type: "paragraph", format: {}, gallery: true },
  {
    id: "NoSpacing",
    name: "No Spacing",
    type: "paragraph",
    format: { spaceBefore: 0, spaceAfter: 0, lineRule: "auto", lineValue: 1 },
    gallery: true,
  },
  ...[16, 13, 12, 11, 11, 11].map(
    (size, i): StyleDefinition => ({
      id: `Heading${i + 1}`,
      name: `heading ${i + 1}`,
      type: "paragraph",
      format: {
        spaceBefore: (i === 0 ? 12 : 2) * PT,
        spaceAfter: 0,
        keepNext: true,
        keepLines: true,
        fontSize: size,
      },
      run: i < 3 ? { bold: true } : { bold: true, italic: i === 3 },
      next: "Normal",
      gallery: i < 3,
    }),
  ),
  {
    id: "Title",
    name: "Title",
    type: "paragraph",
    format: { spaceAfter: 0, lineRule: "auto", lineValue: 1, contextualSpacing: true, fontSize: 28 },
    next: "Normal",
    gallery: true,
  },
  {
    id: "Subtitle",
    name: "Subtitle",
    type: "paragraph",
    format: { spaceAfter: 8 * PT, fontSize: 11 },
    run: { italic: true },
    next: "Normal",
    gallery: true,
  },
  {
    id: "Quote",
    name: "Quote",
    type: "paragraph",
    format: {
      align: "center",
      indentLeft: 0.6 * 96,
      indentRight: 0.6 * 96,
      spaceBefore: 10 * PT,
      spaceAfter: 8 * PT,
    },
    run: { italic: true },
    next: "Normal",
    gallery: true,
  },
  {
    id: "ListParagraph",
    name: "List Paragraph",
    type: "paragraph",
    format: { indentLeft: 48, contextualSpacing: true },
    gallery: true,
  },
  { id: "Strong", name: "Strong", type: "character", run: { bold: true }, gallery: true },
  { id: "Emphasis", name: "Emphasis", type: "character", run: { italic: true }, gallery: true },
];

const GALLERY_ORDER = ["Normal", "NoSpacing", "Heading1", "Heading2", "Heading3", "Title", "Subtitle", "Quote"];

/** Word lower-cases built-in names ("heading 1"); the gallery shows them capitalised. */
export function displayName(style: StyleDefinition): string {
  return style.name.replace(/^./, (c) => c.toUpperCase());
}

export function headingLevelOf(styleId: string | undefined): number | null {
  const m = styleId?.match(/^Heading([1-6])$/);
  return m ? Number(m[1]) : null;
}

const cache = new WeakMap<object, StyleDefinition[]>();

/** All styles of a document: its own (imported, then edited) plus the built-ins it lacks. */
export function documentStyles(doc: PMNode): StyleDefinition[] {
  const hit = cache.get(doc.attrs);
  if (hit) return hit;
  const out = stylesOf(doc.attrs);
  cache.set(doc.attrs, out);
  return out;
}

/** `documentStyles` from the doc node's attrs (also used on JSON for export). */
export function stylesOf(attrs: { layout?: unknown; styles?: unknown } | undefined): StyleDefinition[] {
  const layout = (attrs?.layout as DocumentLayout | null | undefined) ?? null;
  const byId = new Map<string, StyleDefinition>();
  for (const s of layout?.styles ?? []) byId.set(s.id, s);
  for (const s of (attrs?.styles as StyleDefinition[] | null | undefined) ?? []) byId.set(s.id, s);
  const normal = byId.get("Normal")?.format ?? {};
  for (const s of BUILTIN_STYLES) {
    if (byId.has(s.id)) continue;
    // Built-ins missing from a Word document inherit its Normal style.
    byId.set(s.id, s.format ? { ...s, format: { ...normal, ...s.format } } : s);
  }
  return [...byId.values()];
}

export function findStyle(doc: PMNode, id: string | undefined): StyleDefinition | undefined {
  return id ? documentStyles(doc).find((s) => s.id === id) : undefined;
}

/** Styles shown in the gallery, in Word's order. */
export function galleryStyles(doc: PMNode): StyleDefinition[] {
  const rank = (s: StyleDefinition) => {
    const i = GALLERY_ORDER.indexOf(s.id);
    return i < 0 ? GALLERY_ORDER.length + (s.type === "character" ? 100 : 0) : i;
  };
  return documentStyles(doc)
    .filter((s) => s.gallery)
    .sort((a, b) => rank(a) - rank(b) || displayName(a).localeCompare(displayName(b)));
}

/** Style id of a paragraph: its `format.styleId`, else what its node type implies. */
export function paragraphStyleId(node: PMNode): string {
  const id = (node.attrs.format as ParagraphFormat | null)?.styleId;
  if (id) return id;
  return node.type === schema.nodes.heading ? `Heading${node.attrs.level}` : "Normal";
}

type RunKey = keyof CharacterPatch;

const RUN_MARKS: [RunKey, () => MarkType, (v: unknown) => Record<string, unknown>][] = [
  ["bold", () => schema.marks.strong, () => ({})],
  ["italic", () => schema.marks.em, () => ({})],
  ["underline", () => schema.marks.underline, (style) => ({ style })],
  ["strike", () => schema.marks.strike, (style) => ({ style })],
  ["color", () => schema.marks.textColor, (color) => ({ color })],
  ["highlight", () => schema.marks.highlight, (color) => ({ color })],
  ["caps", () => schema.marks.caps, (style) => ({ style })],
  ["vertAlign", () => schema.marks.vertAlign, (align) => ({ align })],
];

/** Swap the run formatting of `oldStyle` for that of `newStyle` over `from`–`to`. */
function swapStyleRuns(
  tr: Transaction,
  from: number,
  to: number,
  oldRun: Record<string, unknown> | undefined,
  newRun: Record<string, unknown> | undefined,
) {
  if (to <= from) return;
  for (const [key, type, attrs] of RUN_MARKS) {
    const before = oldRun?.[key];
    const after = newRun?.[key];
    if (before && !after) tr.removeMark(from, to, type());
    if (after && after !== before) tr.addMark(from, to, type().create(attrs(after)));
  }
}

function cleanFormat(format: ParagraphFormat): ParagraphFormat | null {
  const entries = Object.entries(format).filter(([, v]) => v !== undefined && v !== null);
  return entries.length ? (Object.fromEntries(entries) as ParagraphFormat) : null;
}

/** Paragraph format a style gives a paragraph. */
export function styleFormat(style: StyleDefinition): ParagraphFormat | null {
  const format: ParagraphFormat = { ...(style.format ?? {}) };
  // Normal paragraphs of plain (non-Word) documents carry no format at all.
  if (style.id !== "Normal" || cleanFormat(format)) format.styleId = style.id;
  return cleanFormat(format);
}

/**
 * Restyle one paragraph: node type (headings), `format`, and the style's run formatting.
 * Inside a list item the paragraph stays a paragraph.
 */
function restyle(tr: Transaction, pos: number, style: StyleDefinition) {
  const node = tr.doc.nodeAt(pos);
  if (!node?.isTextblock) return;
  const old = findStyle(tr.doc, paragraphStyleId(node));
  const format = styleFormat(style);
  const inList = tr.doc.resolve(pos).parent.type === schema.nodes.list_item;
  const level = inList ? null : headingLevelOf(style.id);
  const type = level ? schema.nodes.heading : schema.nodes.paragraph;
  if (type !== node.type || (level && node.attrs.level !== level)) {
    tr.setNodeMarkup(pos, type, { ...node.attrs, level: level ?? undefined, format });
  } else if (JSON.stringify(node.attrs.format) !== JSON.stringify(format)) {
    tr.setNodeAttribute(pos, "format", format);
  }
  swapStyleRuns(tr, pos + 1, pos + node.nodeSize - 1, old?.run, style.run);
}

function selectedTextblocks(state: EditorState) {
  const { from, to } = state.selection;
  const out: number[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock) {
      out.push(pos);
      return false;
    }
    return true;
  });
  return out;
}

/** Apply a style to the selection, like clicking it in Word's styles gallery. */
export function applyStyle(id: string): Command {
  return (state, dispatch) => {
    const style = findStyle(state.doc, id);
    if (!style) return false;
    if (style.type === "character") {
      return applyCharacterFormat((style.run ?? {}) as CharacterPatch)(state, dispatch);
    }
    const blocks = selectedTextblocks(state);
    if (!blocks.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const pos of blocks) restyle(tr, pos, style);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Style id of the paragraph at the selection start. */
export function currentStyleId(state: EditorState): string {
  const pos = selectedTextblocks(state)[0];
  const node = pos === undefined ? state.selection.$from.parent : state.doc.nodeAt(pos)!;
  return node.isTextblock ? paragraphStyleId(node) : "Normal";
}

/**
 * Word's "Update <style> to Match Selection": redefine the paragraph's style from its
 * paragraph formatting and the selection's character formatting, then restyle every
 * paragraph using it.
 */
export const updateStyleToMatch: Command = (state, dispatch) => {
  const id = currentStyleId(state);
  const style = findStyle(state.doc, id);
  if (!style || style.type !== "paragraph") return false;
  if (!dispatch) return true;
  const node = state.selection.$from.parent;
  const format: ParagraphFormat = { ...((node.attrs.format as ParagraphFormat | null) ?? {}) };
  delete format.styleId;
  const chars = characterFormatAt(state);
  const run: Partial<RunFormat> = {};
  for (const [key] of RUN_MARKS) {
    const v = chars[key as keyof RunFormat];
    if (v) Object.assign(run, { [key]: v });
  }
  const updated: StyleDefinition = {
    ...style,
    format: { ...format, fontFamily: chars.family, fontSize: chars.size },
    run: Object.keys(run).length ? run : undefined,
  };
  const overrides = ((state.doc.attrs.styles as StyleDefinition[] | null) ?? []).filter((s) => s.id !== id);
  const tr = state.tr.setDocAttribute("styles", [...overrides, updated]);
  const positions: number[] = [];
  tr.doc.descendants((n, pos) => {
    if (n.isTextblock) {
      if (paragraphStyleId(n) === id) positions.push(pos);
      return false;
    }
    return true;
  });
  for (const pos of positions) {
    const n = tr.doc.nodeAt(pos)!;
    const target = styleFormat(updated);
    if (JSON.stringify(n.attrs.format) !== JSON.stringify(target)) tr.setNodeAttribute(pos, "format", target);
    swapStyleRuns(tr, pos + 1, pos + n.nodeSize - 1, style.run, updated.run);
  }
  dispatch(tr);
  return true;
};

/** Node type and attrs for the paragraph Enter creates after `node` (Word's "Style for following paragraph"). */
export function nextParagraph(doc: PMNode, node: PMNode) {
  const style = findStyle(doc, paragraphStyleId(node));
  const next = style?.next && style.next !== style.id ? findStyle(doc, style.next) : undefined;
  if (!next) return null;
  return { type: schema.nodes.paragraph, attrs: { format: styleFormat(next) } };
}

/** Merge CRM template styles into the document (`te:load-template`). */
export function mergeTemplateStyles(state: EditorState, incoming: StyleDefinition[]): Transaction {
  const layout = (state.doc.attrs.layout as DocumentLayout | null) ?? null;
  const docStyles = new Map(
    ((state.doc.attrs.styles as StyleDefinition[] | null) ?? []).map((s) => [s.id, s]),
  );
  const layoutStyles = new Map((layout?.styles ?? []).map((s) => [s.id, s]));
  for (const s of incoming) {
    const prev = docStyles.get(s.id) ?? layoutStyles.get(s.id);
    const merged: StyleDefinition = {
      ...prev,
      ...s,
      format: { ...prev?.format, ...s.format },
      run: { ...prev?.run, ...s.run },
    };
    docStyles.set(s.id, merged);
    layoutStyles.set(s.id, merged);
  }
  let tr = state.tr.setDocAttribute("styles", [...docStyles.values()]).setMeta(SKIP_TRACKING, true);
  if (layout) {
    tr = tr.setDocAttribute("layout", { ...layout, styles: [...layoutStyles.values()] });
  }
  return tr;
}
