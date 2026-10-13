import type { MarkType } from "prosemirror-model";
import { TextSelection, type Command, type EditorState } from "prosemirror-state";
import { schema } from "./schema";

export function isMarkActive(state: EditorState, type: MarkType): boolean {
  const { from, $from, to, empty } = state.selection;
  if (empty) return !!type.isInSet(state.storedMarks ?? $from.marks());
  return state.doc.rangeHasMark(from, to, type);
}

/** Insert a hard page break at the cursor's top-level block (Ctrl/Cmd+Enter). */
export const insertPageBreak: Command = (state, dispatch) => {
  const { $from } = state.selection;
  if ($from.depth < 1) return false;
  if (!dispatch) return true;

  const tr = state.tr;
  const blockStart = $from.before(1);
  let breakPos: number;
  if ($from.parentOffset === 0 && $from.parent.content.size > 0) {
    breakPos = blockStart;
  } else {
    if ($from.parent.isTextblock && $from.parentOffset < $from.parent.content.size) {
      tr.split($from.pos);
    }
    breakPos = blockStart + tr.doc.nodeAt(blockStart)!.nodeSize;
  }
  tr.insert(breakPos, schema.nodes.page_break.create());
  const after = breakPos + 1;
  if (after >= tr.doc.content.size) {
    tr.insert(after, schema.nodes.paragraph.create());
  }
  tr.setSelection(TextSelection.near(tr.doc.resolve(after + 1)));
  dispatch(tr.scrollIntoView());
  return true;
};
