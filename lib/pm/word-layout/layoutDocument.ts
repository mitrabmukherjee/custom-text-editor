import type { Node as PMNode } from "prosemirror-model";
import type { DocumentLayout } from "../format";
import { sectionsOf } from "../sections";
import { EditorState } from "prosemirror-state";
import { computePageBreaksWithContext } from "../pagination/computePageBreaks";
import { LETTER_PAGE_SETUP } from "../pagination/pageSetup";
import { buildPaginationContext } from "./pageGeometry";
import { layoutBody } from "./layoutBody";
import { collectLineWrapPositions } from "./positions";
import type { LayoutBodyOptions, LayoutResult } from "./types";

export function layoutDocument(
  doc: PMNode,
  options: Partial<LayoutBodyOptions> = {},
): LayoutResult {
  const state = EditorState.create({ doc });
  const sections = sectionsOf(state);
  const layout = doc.attrs.layout as DocumentLayout | null;
  const defaultSetup = options.defaultSetup ?? layout?.page ?? LETTER_PAGE_SETUP;
  const bodyOpts: LayoutBodyOptions = {
    defaultSetup,
    defaultTabStop: layout?.defaultTabStop ?? 48,
    justifyShrink: layout?.justifyShrink,
    getDocGridLinePitchTwips: (s) => sections[s]?.docGridLinePitch,
    includeDeletionText: options.includeDeletionText,
    ...options,
  };
  const body = layoutBody(doc, bodyOpts);
  const ctx = buildPaginationContext(layout, sections, defaultSetup);
  const pagination = computePageBreaksWithContext(body.lines, ctx, {
    numberStart: (section) => sections[section]?.pageNumberStart,
  }, defaultSetup);
  const lineWrapPositions = collectLineWrapPositions(doc, body.lines);
  return {
    lines: body.lines,
    pagination,
    lineWrapPositions,
    tabWidths: body.tabWidths,
  };
}
