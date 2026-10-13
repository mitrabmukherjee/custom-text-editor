import type { Node as PMNode, NodeType } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { liftListItem, sinkListItem, wrapInList } from "prosemirror-schema-list";
import type { ParagraphFormat } from "./format";
import { schema } from "./schema";

/** Word's default indent step (0.5 inch). */
export const INDENT_STEP = 48;

export type FormatPatch =
  | Partial<ParagraphFormat>
  | ((format: ParagraphFormat) => Partial<ParagraphFormat>);

function textblocksInSelection(state: EditorState) {
  const { from, to } = state.selection;
  const blocks: { node: PMNode; pos: number }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock) {
      blocks.push({ node, pos });
      return false;
    }
    return true;
  });
  return blocks;
}

function cleanFormat(format: ParagraphFormat): ParagraphFormat | null {
  const entries = Object.entries(format).filter(([, v]) => v !== undefined && v !== null);
  return entries.length ? (Object.fromEntries(entries) as ParagraphFormat) : null;
}

/** Format of the paragraph holding the selection start ({} when unformatted). */
export function paragraphFormatAt(state: EditorState): ParagraphFormat {
  const block = textblocksInSelection(state)[0]?.node ?? state.selection.$from.parent;
  return (block.attrs.format as ParagraphFormat | null) ?? {};
}

/** Merge `patch` into the format of every paragraph in the selection, like Word's Paragraph dialog. */
export function setParagraphFormat(patch: FormatPatch): Command {
  return (state, dispatch) => {
    const blocks = textblocksInSelection(state);
    if (!blocks.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of blocks) {
        const current = (node.attrs.format as ParagraphFormat | null) ?? {};
        const changes = typeof patch === "function" ? patch(current) : patch;
        tr.setNodeAttribute(pos, "format", cleanFormat({ ...current, ...changes }));
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function listDepth(state: EditorState) {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const type = $from.node(d).type;
    if (type === schema.nodes.bullet_list || type === schema.nodes.ordered_list) return d;
  }
  return 0;
}

/** Increase/decrease indent: nest list items, otherwise move the left indent to the next 0.5" stop. */
export function indent(direction: 1 | -1): Command {
  return (state, dispatch) => {
    if (listDepth(state)) {
      const command =
        direction > 0 ? sinkListItem(schema.nodes.list_item) : liftListItem(schema.nodes.list_item);
      return command(state, dispatch);
    }
    return setParagraphFormat((f) => {
      const left = f.indentLeft ?? 0;
      const next =
        direction > 0
          ? (Math.floor(left / INDENT_STEP + 1e-6) + 1) * INDENT_STEP
          : (Math.ceil(left / INDENT_STEP - 1e-6) - 1) * INDENT_STEP;
      return { indentLeft: Math.max(0, next) || undefined };
    })(state, dispatch);
  };
}

/** Word's Ctrl+T: hang the first line by one more step; `-1` takes a step back. */
export function hangingIndent(direction: 1 | -1): Command {
  return setParagraphFormat((f) => {
    const left = f.indentLeft ?? 0;
    const hanging = Math.max(0, -(f.indentFirstLine ?? 0));
    const nextLeft = Math.max(0, left + direction * INDENT_STEP);
    const nextHanging = Math.max(0, Math.min(nextLeft, hanging + direction * INDENT_STEP));
    return { indentLeft: nextLeft || undefined, indentFirstLine: nextHanging ? -nextHanging : undefined };
  });
}

/** Bullets / numbering button: wrap in a list, switch list type, or remove the list. */
export function toggleList(listType: NodeType): Command {
  return (state, dispatch) => {
    const depth = listDepth(state);
    if (!depth) return wrapInList(listType)(state, dispatch);
    const list = state.selection.$from.node(depth);
    if (list.type === listType) return liftListItem(schema.nodes.list_item)(state, dispatch);
    dispatch?.(state.tr.setNodeMarkup(state.selection.$from.before(depth), listType));
    return true;
  };
}

export function activeList(state: EditorState): NodeType | null {
  const depth = listDepth(state);
  return depth ? state.selection.$from.node(depth).type : null;
}
