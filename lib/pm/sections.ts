import { NodeSelection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { DocumentLayout, LineNumbering, SectionLayout, SectionStart } from "./format";
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
  const sections = sectionLayouts(layout, breaks);
  if (!layout) {
    const last = sections.length - 1;
    sections[last] = {
      ...sections[last],
      lineNumbers: (state.doc.attrs.lineNumbers as LineNumbering | null) ?? null,
    };
  }
  return sections;
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

/** How line numbers restart, or `"none"` to turn them off. */
export type LineNumberMode = "none" | LineNumbering["restart"];

const DEFAULT_LINE_NUMBERS: LineNumbering = {
  countBy: 1,
  start: 1,
  restart: "newPage",
  distance: 24,
};

function lineNumbersFor(prev: LineNumbering | null | undefined, mode: LineNumberMode): LineNumbering | null {
  if (mode === "none") return null;
  return {
    countBy: prev?.countBy ?? DEFAULT_LINE_NUMBERS.countBy,
    start: prev?.start ?? DEFAULT_LINE_NUMBERS.start,
    restart: mode,
    distance: prev?.distance ?? DEFAULT_LINE_NUMBERS.distance,
    ...(prev?.fontFamily ? { fontFamily: prev.fontFamily } : {}),
    ...(prev?.fontSize !== undefined ? { fontSize: prev.fontSize } : {}),
  };
}

/**
 * Turn line numbers on or off for one section. Restart changes keep imported count, start, distance, and font.
 * A document with no Word layout stores them on the doc, so the page stays as it is.
 */
export function setSectionLineNumbers(
  state: EditorState,
  sectionIndex: number,
  mode: LineNumberMode,
): Transaction | null {
  const layout = state.doc.attrs.layout as DocumentLayout | null;
  let breaks = 0;
  state.doc.descendants((node) => {
    if (node.type === schema.nodes.section_break) breaks += 1;
    return !node.isTextblock;
  });
  if (!layout && sectionIndex >= breaks) {
    const prev = state.doc.attrs.lineNumbers as LineNumbering | null;
    return state.tr
      .setDocAttribute("lineNumbers", lineNumbersFor(prev, mode))
      .setMeta(SKIP_TRACKING, true);
  }
  return patchSectionLayout(state, sectionIndex, (section) => ({
    ...section,
    lineNumbers: lineNumbersFor(section.lineNumbers, mode),
  }));
}
