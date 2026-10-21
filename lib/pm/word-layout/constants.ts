import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { DocumentLayout, ParagraphFormat } from "../format";
import { sectionsOf } from "../sections";

/** When true, Word-layout docs may use the layout engine (see `wordLayoutEngineDrivesUI`). */
function envStrict(): boolean {
  if (typeof process !== "undefined" && process.env?.WORD_LAYOUT_STRICT === "0") return false;
  if (typeof localStorage !== "undefined" && localStorage.getItem("WORD_LAYOUT_STRICT") === "0") {
    return false;
  }
  return true;
}

export const WORD_LAYOUT_STRICT = envStrict();

/**
 * When true, engine breaks and line wraps drive the editor (not DOM measurement).
 * Default is on for Word-layout documents; set `WORD_LAYOUT_ENGINE=0` to fall back to DOM.
 */
export function wordLayoutEngineDrivesUI(): boolean {
  if (!WORD_LAYOUT_STRICT) return false;
  if (typeof localStorage !== "undefined" && localStorage.getItem("WORD_LAYOUT_ENGINE") === "0") {
    return false;
  }
  if (typeof process !== "undefined" && process.env?.WORD_LAYOUT_ENGINE === "0") return false;
  return true;
}

export function wordLayoutForDocument(hasWordLayout: boolean): boolean {
  return hasWordLayout && WORD_LAYOUT_STRICT;
}

/** Run engine layout alongside DOM for comparison logging. */
export function wordLayoutCompareMode(hasWordLayout: boolean): boolean {
  return wordLayoutForDocument(hasWordLayout) && !wordLayoutEngineDrivesUI();
}

function hfSetNonEmpty(set: { default?: unknown[]; first?: unknown[]; even?: unknown[] } | undefined): boolean {
  if (!set) return false;
  return !!(set.default?.length || set.first?.length || set.even?.length);
}

function legacyLayoutLooksImported(doc: PMNode): boolean {
  const state = EditorState.create({ doc });
  for (const sec of sectionsOf(state)) {
    if (sec.lineNumbers || sec.docGridLinePitch) return true;
    if ((sec.columns?.count ?? 1) > 1) return true;
    if (hfSetNonEmpty(sec.headers) || hfSetNonEmpty(sec.footers)) return true;
  }
  let imported = false;
  doc.descendants((node) => {
    if (imported) return false;
    if (node.type.name === "section_break") {
      imported = true;
      return false;
    }
    if (node.isTextblock) {
      const f = node.attrs.format as ParagraphFormat | null;
      if (f?.docGridSectionPitchTwips) imported = true;
    }
    return !imported;
  });
  return imported;
}

/** Whether the word-layout engine may drive pagination for this document. */
export function documentUsesLayoutEngine(doc: PMNode): boolean {
  const layout = doc.attrs.layout as DocumentLayout | null;
  if (!layout) return false;
  if (layout.layoutEngine === true) return true;
  if (layout.layoutEngine === false) return false;
  return legacyLayoutLooksImported(doc);
}
