import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import type { DocumentLayout, ParagraphFormat } from "./format";
import { schema } from "./schema";
import { FORMAT_MARK_NAMES, runFormatOf, type RunFormat } from "./track-changes/formatting";

/** Word's font size list. */
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export const FONT_FAMILIES = [
  "Arial",
  "Calibri",
  "Cambria",
  "Courier New",
  "Garamond",
  "Georgia",
  "Helvetica",
  "Tahoma",
  "Times New Roman",
  "Verdana",
];

const DEFAULT_FAMILY = "Times New Roman";

/** Font and size text gets without a `textStyle` mark: the paragraph mark's, else the document default. */
function baseFont(state: EditorState, paragraph: PMNode | null) {
  const format = (paragraph?.attrs.format as ParagraphFormat | null) ?? null;
  const word = !!(state.doc.attrs.layout as DocumentLayout | null);
  return {
    fontFamily: format?.fontFamily ?? DEFAULT_FAMILY,
    fontSize: format?.fontSize ?? (word ? 10 : 12),
  };
}

/** Character formatting at the selection, with the effective font family and size filled in. */
export function characterFormatAt(state: EditorState): RunFormat & { family: string; size: number } {
  const { $from, empty } = state.selection;
  let marks: readonly Mark[] = state.storedMarks ?? $from.marks();
  if (!empty) {
    // Like Word, report the formatting of the first selected character.
    let found: readonly Mark[] | null = null;
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
      if (found) return false;
      if (node.isText) found = node.marks;
      return !found;
    });
    if (found) marks = found;
  }
  const format = runFormatOf(marks);
  const base = baseFont(state, $from.parent.isTextblock ? $from.parent : null);
  return {
    ...format,
    family: format.fontFamily ?? base.fontFamily,
    size: format.fontSize ?? base.fontSize,
  };
}

type Segment = { from: number; to: number; node: PMNode; paragraph: PMNode | null };

function textSegments(state: EditorState): Segment[] {
  const { from, to } = state.selection;
  const out: Segment[] = [];
  state.doc.nodesBetween(from, to, (node, pos, parent) => {
    if (!node.isInline) return true;
    out.push({
      from: Math.max(pos, from),
      to: Math.min(pos + node.nodeSize, to),
      node,
      paragraph: parent?.isTextblock ? parent : null,
    });
    return false;
  });
  return out;
}

type FontPatch = { fontFamily?: string | null; fontSize?: number | null };

/**
 * Change the font family and/or size of the selection. `patch` gets each run's effective
 * font; values equal to the paragraph's font are stored as "no override".
 */
export function setTextStyle(patch: (current: { family: string; size: number }) => FontPatch): Command {
  return (state, dispatch) => {
    const { textStyle } = schema.marks;
    if (state.selection.empty) {
      if (!dispatch) return true;
      const current = characterFormatAt(state);
      const base = baseFont(state, state.selection.$from.parent);
      const next = { fontFamily: current.fontFamily, fontSize: current.fontSize, ...patch(current) };
      const attrs = {
        fontFamily: next.fontFamily === base.fontFamily ? null : (next.fontFamily ?? null),
        fontSize: next.fontSize === base.fontSize ? null : (next.fontSize ?? null),
      };
      const tr = state.tr.removeStoredMark(textStyle);
      if (attrs.fontFamily || attrs.fontSize) tr.addStoredMark(textStyle.create(attrs));
      dispatch(tr);
      return true;
    }
    const segments = textSegments(state).filter((s) => s.node.isText);
    if (!segments.length) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    for (const seg of segments) {
      const style = textStyle.isInSet(seg.node.marks);
      const base = baseFont(state, seg.paragraph);
      const current = {
        family: (style?.attrs.fontFamily as string | null) ?? base.fontFamily,
        size: (style?.attrs.fontSize as number | null) ?? base.fontSize,
      };
      const next = {
        fontFamily: current.family,
        fontSize: current.size,
        ...patch(current),
      };
      const attrs = {
        fontFamily: next.fontFamily === base.fontFamily ? null : (next.fontFamily ?? null),
        fontSize: next.fontSize === base.fontSize ? null : (next.fontSize ?? null),
      };
      const same =
        (style?.attrs.fontFamily ?? null) === attrs.fontFamily &&
        (style?.attrs.fontSize ?? null) === attrs.fontSize;
      if (same) continue;
      if (style) tr.removeMark(seg.from, seg.to, textStyle);
      if (attrs.fontFamily || attrs.fontSize) tr.addMark(seg.from, seg.to, textStyle.create(attrs));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const setFontFamily = (family: string) => setTextStyle(() => ({ fontFamily: family }));

export const setFontSize = (size: number) =>
  setTextStyle(() => ({ fontSize: Math.max(1, Math.min(1638, Math.round(size * 2) / 2)) }));

/** Word's Ctrl+Shift+> / < (next size in the list) and Ctrl+] / [ (one point). */
export function growFont(direction: 1 | -1, byPoint = false): Command {
  return setTextStyle(({ size }) => {
    if (byPoint) return { fontSize: Math.max(1, size + direction) };
    const list = direction > 0 ? FONT_SIZES : [...FONT_SIZES].reverse();
    const next = list.find((s) => (direction > 0 ? s > size : s < size));
    return { fontSize: next ?? (direction > 0 ? size + 10 : Math.max(1, size - 1)) };
  });
}

function markAt(state: EditorState, type: MarkType): Mark | null {
  const { from, to, empty, $from } = state.selection;
  if (empty) return type.isInSet(state.storedMarks ?? $from.marks()) ?? null;
  let found: Mark | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (found) return false;
    if (node.isText) found = type.isInSet(node.marks) ?? null;
    return !found && !node.isText;
  });
  return found;
}

/** Set (or with `null`, remove) a mark with attributes on the selection or the stored marks. */
export function setMark(type: MarkType, attrs: Record<string, unknown> | null): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { from, to, empty } = state.selection;
    let tr: Transaction = state.tr;
    if (empty) {
      tr = tr.removeStoredMark(type);
      if (attrs) tr.addStoredMark(type.create(attrs));
      dispatch(tr);
      return true;
    }
    const segments = textSegments(state).filter((s) => s.node.isText);
    for (const seg of segments) {
      const existing = type.isInSet(seg.node.marks);
      if (attrs && existing && type.create(attrs).eq(existing)) continue;
      if (!attrs && !existing) continue;
      if (existing) tr.removeMark(seg.from, seg.to, type);
      if (attrs) tr.addMark(seg.from, seg.to, type.create(attrs));
    }
    if (tr.docChanged) dispatch(tr.scrollIntoView());
    return from !== to;
  };
}

/** Toggle a mark with attributes: off when the selection starts with it (same `key` attr), on otherwise. */
export function toggleAttrMark(
  type: MarkType,
  attrs: Record<string, unknown>,
  key?: string,
): Command {
  return (state, dispatch) => {
    const mark = markAt(state, type);
    const on = mark && (!key || mark.attrs[key] === attrs[key]);
    return setMark(type, on ? null : attrs)(state, dispatch);
  };
}

export const toggleUnderline = (style = "single") =>
  toggleAttrMark(schema.marks.underline, { style }, "style");
export const toggleStrike = (style = "single") =>
  toggleAttrMark(schema.marks.strike, { style }, "style");
export const toggleSuperscript = toggleAttrMark(schema.marks.vertAlign, { align: "superscript" }, "align");
export const toggleSubscript = toggleAttrMark(schema.marks.vertAlign, { align: "subscript" }, "align");
export const toggleCaps = (style: "all" | "small") =>
  toggleAttrMark(schema.marks.caps, { style }, "style");

export const setTextColor = (color: string | null) =>
  setMark(schema.marks.textColor, color ? { color } : null);
export const setHighlight = (color: string | null) =>
  setMark(schema.marks.highlight, color ? { color } : null);

/** Run commands one after another as a single transaction (one undo step, one revision). */
export function sequence(...commands: Command[]): Command {
  return (state, dispatch) => {
    if (!dispatch) return commands.every((c) => c(state));
    let current = state;
    const tr = state.tr;
    for (const command of commands) {
      command(current, (t) => {
        for (const step of t.steps) tr.step(step);
        if (t.storedMarksSet) tr.setStoredMarks(t.storedMarks);
        current = current.apply(t);
      });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export type CharacterPatch = Partial<
  Pick<RunFormat, "bold" | "italic" | "underline" | "strike" | "color" | "highlight" | "caps" | "vertAlign">
> & { family?: string; size?: number };

/** Apply the fields present in `patch` (Font dialog). */
export function applyCharacterFormat(patch: CharacterPatch): Command {
  const m = schema.marks;
  const commands: Command[] = [];
  if (patch.family !== undefined || patch.size !== undefined) {
    commands.push(
      setTextStyle(() => ({
        ...(patch.family !== undefined ? { fontFamily: patch.family } : {}),
        ...(patch.size !== undefined ? { fontSize: patch.size } : {}),
      })),
    );
  }
  const attr = <T>(key: keyof CharacterPatch, type: MarkType, attrs: (v: T) => Record<string, unknown>) => {
    if (!(key in patch)) return;
    const v = patch[key] as T | null | undefined;
    commands.push(setMark(type, v ? attrs(v) : null));
  };
  attr<boolean>("bold", m.strong, () => ({}));
  attr<boolean>("italic", m.em, () => ({}));
  attr<string>("underline", m.underline, (style) => ({ style }));
  attr<string>("strike", m.strike, (style) => ({ style }));
  attr<string>("color", m.textColor, (color) => ({ color }));
  attr<string>("highlight", m.highlight, (color) => ({ color }));
  attr<string>("caps", m.caps, (style) => ({ style }));
  attr<string>("vertAlign", m.vertAlign, (align) => ({ align }));
  return sequence(...commands);
}

/** Word's Clear Formatting (Ctrl+Space): remove every character formatting mark. */
export const clearFormatting: Command = (state, dispatch) => {
  const { from, to, empty } = state.selection;
  const types = FORMAT_MARK_NAMES.map((n) => schema.marks[n]);
  if (empty) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    dispatch?.(state.tr.setStoredMarks(marks.filter((m) => !types.includes(m.type))));
    return true;
  }
  if (!types.some((t) => state.doc.rangeHasMark(from, to, t))) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const type of types) {
      if (state.doc.rangeHasMark(from, to, type)) tr.removeMark(from, to, type);
    }
    dispatch(tr.scrollIntoView());
  }
  return true;
};

export function isAttrMarkActive(state: EditorState, type: MarkType, key?: string, value?: unknown) {
  const mark = markAt(state, type);
  return !!mark && (!key || mark.attrs[key] === value);
}
