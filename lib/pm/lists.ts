import type { Node as PMNode, NodeType } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { sinkListItem, wrapInList } from "prosemirror-schema-list";
import { schema, type BulletStyle, type NumFormat } from "./schema";

export type MultilevelPreset = { id: string; label: string; levels: (NumFormat | null)[] };

/** Word's multilevel list library (the formats of levels 1, 2, 3…). */
export const MULTILEVEL_PRESETS: MultilevelPreset[] = [
  { id: "default", label: "1. a. i.", levels: ["decimal", "lowerLetter", "lowerRoman"] },
  { id: "legal", label: "1. 1.1. 1.1.1.", levels: ["legal"] },
  {
    id: "outline",
    label: "I. A. 1. a. i.",
    levels: ["upperRoman", "upperLetter", "decimal", "lowerLetter", "lowerRoman"],
  },
];

export const NUMBER_FORMATS: { format: NumFormat; label: string }[] = [
  { format: "decimal", label: "1. 2. 3." },
  { format: "lowerLetter", label: "a. b. c." },
  { format: "upperLetter", label: "A. B. C." },
  { format: "lowerRoman", label: "i. ii. iii." },
  { format: "upperRoman", label: "I. II. III." },
  { format: "legal", label: "1. 1.1. 1.1.1." },
];

export const BULLET_STYLES: { bullet: BulletStyle; label: string }[] = [
  { bullet: "disc", label: "\u2022 Disc" },
  { bullet: "circle", label: "o Circle" },
  { bullet: "square", label: "\u25aa Square" },
  { bullet: "dash", label: "\u2013 Dash" },
  { bullet: "arrow", label: "\u27a2 Arrow" },
  { bullet: "check", label: "\u2713 Check" },
  { bullet: "diamond", label: "\u2756 Diamond" },
];

const isList = (type: NodeType) => type === schema.nodes.ordered_list || type === schema.nodes.bullet_list;

/** Depth of the innermost list (of `type`, if given) around the selection; 0 when none. */
function listDepth(state: EditorState, type?: NodeType) {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const t = $from.node(d).type;
    if (type ? t === type : isList(t)) return d;
  }
  return 0;
}

/** Put the selection in a list of `type`: wrap, or switch the innermost list's type. */
function ensureList(type: NodeType, attrs: Record<string, unknown>): Command {
  return (state, dispatch) => {
    const depth = listDepth(state);
    if (!depth) {
      return wrapInList(type, attrs)(state, dispatch);
    }
    if (!dispatch) return true;
    const $from = state.selection.$from;
    const list = $from.node(depth);
    const pos = $from.before(depth);
    if (list.type === type) {
      dispatch(state.tr.setNodeMarkup(pos, type, { ...list.attrs, ...attrs }));
    } else {
      dispatch(state.tr.setNodeMarkup(pos, type, attrs));
    }
    return true;
  };
}

/** Numbering library: number the selection with `format` (null: depth default). */
export const setNumbering = (format: NumFormat | null) =>
  ensureList(schema.nodes.ordered_list, { numFormat: format });

/** Bullet library. */
export const setBullet = (bullet: BulletStyle | null) => ensureList(schema.nodes.bullet_list, { bullet });

/** Apply a multilevel list: number the selection, then set every nested level's format. */
export function applyMultilevel(preset: MultilevelPreset): Command {
  return (state, dispatch) => {
    const tr = state.tr;
    let current = state;
    if (!listDepth(state, schema.nodes.ordered_list)) {
      const wrapped = setNumbering(preset.levels[0] ?? null)(state, (t) => {
        for (const step of t.steps) tr.step(step);
        current = state.apply(t);
      });
      if (!wrapped) return false;
    }
    if (!dispatch) return true;
    const { $from } = current.selection;
    let outer = 0;
    for (let d = 1; d <= $from.depth; d++) {
      if ($from.node(d).type === schema.nodes.ordered_list) {
        outer = d;
        break;
      }
    }
    if (!outer) return false;
    const start = $from.before(outer);
    const root = tr.doc.nodeAt(start)!;
    const visit = (node: PMNode, pos: number, level: number) => {
      if (node.type === schema.nodes.ordered_list) {
        // Legal numbering continues into nested levels that have no format of their own.
        const format = preset.levels[level] ?? (preset.id === "legal" ? null : node.attrs.numFormat);
        tr.setNodeMarkup(pos, null, { ...node.attrs, numFormat: format });
      }
      const nested = isList(node.type) ? level + 1 : level;
      node.forEach((child, offset) => visit(child, pos + 1 + offset, nested));
    };
    visit(root, start, 0);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Innermost ordered list around the selection and the index of the selected item. */
function orderedListAt(state: EditorState) {
  const depth = listDepth(state, schema.nodes.ordered_list);
  if (!depth) return null;
  const { $from } = state.selection;
  return { depth, node: $from.node(depth), pos: $from.before(depth), index: $from.index(depth) };
}

/** Word's "Restart at 1": from the selected item on, numbering starts again. */
export const restartNumbering: Command = (state, dispatch) => {
  const list = orderedListAt(state);
  if (!list) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  if (list.index === 0) {
    tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, order: 1 });
  } else {
    let itemPos = list.pos + 1;
    for (let i = 0; i < list.index; i++) itemPos += list.node.child(i).nodeSize;
    tr.split(itemPos);
    tr.setNodeMarkup(itemPos + 1, null, { ...list.node.attrs, order: 1 });
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/** Ordered lists at the same nesting depth, in document order. */
function listsAtDepth(doc: PMNode, depth: number) {
  const out: { node: PMNode; pos: number }[] = [];
  const visit = (node: PMNode, pos: number, level: number) => {
    node.forEach((child, offset) => {
      const childPos = pos + offset + (node.type === doc.type ? 0 : 1);
      if (child.type === schema.nodes.ordered_list && level === depth) out.push({ node: child, pos: childPos });
      if (!child.isTextblock) visit(child, childPos, isList(child.type) ? level + 1 : level);
    });
  };
  visit(doc, 0, 0);
  return out;
}

/** Word's "Continue Numbering": carry on from the previous list at the same level. */
export const continueNumbering: Command = (state, dispatch) => {
  const list = orderedListAt(state);
  if (!list) return false;
  const { $from } = state.selection;
  let level = 0;
  for (let d = 1; d < list.depth; d++) if (isList($from.node(d).type)) level++;
  const lists = listsAtDepth(state.doc, level);
  const i = lists.findIndex((l) => l.pos === list.pos);
  const prev = i > 0 ? lists[i - 1] : null;
  if (!prev) return false;
  const order = (prev.node.attrs.order as number) + prev.node.childCount;
  if (dispatch) dispatch(state.tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, order }));
  return true;
};

/** Word's "Set Numbering Value". */
export function setNumberingValue(value: number): Command {
  return (state, dispatch) => {
    const list = orderedListAt(state);
    if (!list || value < 0) return false;
    dispatch?.(state.tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, order: Math.round(value) }));
    return true;
  };
}

export function listAttrsAt(state: EditorState) {
  const depth = listDepth(state);
  if (!depth) return null;
  const node = state.selection.$from.node(depth);
  return { type: node.type, attrs: node.attrs, index: state.selection.$from.index(depth) };
}

/** Tab at the start of a list item demotes it, like Word. */
export const tabInList: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!empty || $from.parentOffset !== 0 || !listDepth(state)) return false;
  if ($from.node($from.depth - 1).type !== schema.nodes.list_item) return false;
  return sinkListItem(schema.nodes.list_item)(state, dispatch);
};
