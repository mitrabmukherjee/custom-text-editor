import { NodeSelection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { DocumentLayout, SectionLayout, SectionStart } from "./format";
import { sectionLayouts } from "./format";
import { schema } from "./schema";
import { SKIP_TRACKING } from "./track-changes";

/** Section index (0-based) for document position `pos`. */
export function sectionIndexAt(state: EditorState, pos: number): number {
  let index = 0;
  state.doc.nodesBetween(0, pos, (node, nodePos) => {
    if (node.type === schema.nodes.section_break && nodePos < pos) index += 1;
  });
  return index;
}

export function sectionsOf(state: EditorState): SectionLayout[] {
  const layout = state.doc.attrs.layout as DocumentLayout | null;
  const breaks: (SectionLayout | null)[] = [];
  state.doc.descendants((node) => {
    if (node.type === schema.nodes.section_break) {
      breaks.push(node.attrs.section as SectionLayout | null);
    }
    return !node.isTextblock;
  });
  return sectionLayouts(layout, breaks);
}

/** Insert a Word section break; `start` is how the following section begins. */
export function insertSectionBreak(start: SectionStart): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    const idx = sectionIndexAt(state, $from.pos);
    const section = { ...sectionsOf(state)[idx] };
    const node = schema.nodes.section_break.create({ start, section });
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(node).setMeta(SKIP_TRACKING, true);
      let breakPos = -1;
      tr.doc.descendants((node, pos) => {
        if (breakPos < 0 && node.type === schema.nodes.section_break) breakPos = pos;
      });
      const sel =
        breakPos >= 0
          ? NodeSelection.create(tr.doc, breakPos)
          : TextSelection.near(tr.doc.resolve(Math.min(tr.mapping.map($from.pos) + 1, tr.doc.content.size)), 1);
      dispatch(tr.setSelection(sel).scrollIntoView());
    }
    return true;
  };
}

export function patchSectionLayout(
  state: EditorState,
  sectionIndex: number,
  patch: (section: SectionLayout) => SectionLayout,
): Transaction | null {
  const layout = state.doc.attrs.layout as DocumentLayout | null;
  const breaks: { pos: number; section: SectionLayout | null; start: SectionStart }[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type === schema.nodes.section_break) {
      breaks.push({
        pos,
        section: node.attrs.section as SectionLayout | null,
        start: node.attrs.start as SectionStart,
      });
    }
    return !node.isTextblock;
  });
  if (sectionIndex < breaks.length) {
    const { pos, section, start } = breaks[sectionIndex];
    const next = patch(section ?? {});
    return state.tr
      .setNodeMarkup(pos, undefined, { start, section: next })
      .setMeta(SKIP_TRACKING, true);
  }
  if (!layout) {
    const base: DocumentLayout = {
      page: { width: 816, height: 1056, marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96, gap: 24 },
    };
    return state.tr.setDocAttribute("layout", patch(base)).setMeta(SKIP_TRACKING, true);
  }
  return state.tr.setDocAttribute("layout", patch(layout)).setMeta(SKIP_TRACKING, true);
}

export function setSectionColumns(
  state: EditorState,
  sectionIndex: number,
  count: number,
  space: number,
): Transaction | null {
  const cols = count <= 1 ? undefined : { count, space };
  return patchSectionLayout(state, sectionIndex, (section) => ({
    ...section,
    columns: cols,
  }));
}
