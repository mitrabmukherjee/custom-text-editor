import type { Node as PMNode } from "prosemirror-model";
import type { DocumentLayout, ParagraphFormat, SectionLayout, SectionStart } from "../format";
import { docGridPitchPxFromTwips } from "../format";
import { TWIPS_PER_PX } from "../pagination/pageSetup";
import { sectionsOf } from "../sections";
import { EditorState } from "prosemirror-state";
import { LETTER_PAGE_SETUP, type PageSetup } from "../pagination/pageSetup";
import { buildPaginationContext, geometryForSection } from "./pageGeometry";
import { ColumnFlow } from "./columnFlow";
import { layoutTableRows } from "./layoutTables";
import { layoutParagraph } from "./paragraphLayout";
import type { EngineLayoutLine, LayoutBodyOptions, LayoutBodyResult, LayoutTabWidth } from "./types";

const SECTION_BREAKS: Record<SectionStart, boolean | "odd" | "even"> = {
  nextPage: true,
  continuous: false,
  evenPage: "even",
  oddPage: "odd",
};

/** Doc-grid pitch from section or margin line-number row height (title page without w:docGrid). */
function effectivePitchTwips(
  section: number,
  sections: SectionLayout[],
  format: ParagraphFormat | null,
): number | undefined {
  const sec = sections[section];
  if (sec?.docGridLinePitch) return sec.docGridLinePitch;
  if (format?.docGridSectionPitchTwips) return format.docGridSectionPitchTwips;
  const dist = sec?.lineNumbers?.distance;
  if (dist == null) return undefined;
  const pitch360Px = docGridPitchPxFromTwips(360);
  if (Math.abs(dist - pitch360Px) < 1) return 360;
  return Math.round(dist * TWIPS_PER_PX);
}

export function layoutBody(
  doc: PMNode,
  options: LayoutBodyOptions,
): LayoutBodyResult {
  const state = EditorState.create({ doc });
  const sections = sectionsOf(state);
  const layout = doc.attrs.layout as DocumentLayout | null;
  const defaultSetup = options.defaultSetup ?? LETTER_PAGE_SETUP;
  const defaultTab = options.defaultTabStop ?? layout?.defaultTabStop ?? 48;
  const pagCtx = buildPaginationContext(layout, sections, defaultSetup);
  const firstFrame = pagCtx.getFrame({ section: 0, first: true, number: 1 });

  const lines: EngineLayoutLine[] = [];
  const blocks: LayoutBodyResult["blocks"] = [];
  const tabWidths: LayoutTabWidth[] = [];
  let naturalY = firstFrame.top;
  let block = 0;
  let section = 0;
  const columnFlows = new Map<number, ColumnFlow>();

  const sectionWidth = (col = 0) => {
    const geo = geometryForSection(section, sections, defaultSetup);
    if (options.contentWidthPx != null && geo.frames.length <= 1) {
      return options.contentWidthPx;
    }
    return geo.frames[col]?.width ??
      defaultSetup.width - defaultSetup.marginLeft - defaultSetup.marginRight;
  };

  const flowForSection = () => {
    let flow = columnFlows.get(section);
    if (!flow) {
      const geo = geometryForSection(section, sections, defaultSetup);
      const frame = pagCtx.getFrame({ section, first: section === 0, number: 1 });
      flow = new ColumnFlow(geo, frame, naturalY);
      columnFlows.set(section, flow);
    }
    return flow;
  };

  doc.descendants((node, pos) => {
    const name = node.type.name;
    if (name === "page_break") {
      lines.push({
        top: naturalY,
        bottom: naturalY + 8,
        breakAfter: true,
        block: block++,
        section,
        numbered: false,
        anchor: { pos, textOffset: 0 },
      });
      naturalY += 8;
      return false;
    }
    if (name === "section_break") {
      const breakAfter = SECTION_BREAKS[node.attrs.start as SectionStart] ?? true;
      if (breakAfter) {
        lines.push({
          top: naturalY,
          bottom: naturalY,
          breakAfter,
          block: block++,
          section,
          numbered: false,
          anchor: { pos, textOffset: 0 },
        });
      }
      section += 1;
      columnFlows.clear();
      return false;
    }
    if (node.type.spec.tableRole === "table") {
      const pitch = options.getDocGridLinePitchTwips?.(section);
      const geo = geometryForSection(section, sections, defaultSetup);
      const flow = geo.frames.length > 1 ? flowForSection() : null;
      const tableTop = flow ? flow.blockTop() : naturalY;
      const rowLines = layoutTableRows(
        node,
        pos,
        section,
        block,
        tableTop,
        sectionWidth(0),
        pitch ? docGridPitchPxFromTwips(pitch) : undefined,
      );
      if (flow) {
        const { column, deltaTop } = flow.advance(rowLines.height, () => {
          const br: EngineLayoutLine = {
            top: naturalY,
            bottom: naturalY,
            breakAfter: true,
            block: block++,
            section,
            numbered: false,
            anchor: { pos, textOffset: 0 },
          };
          lines.push(br);
          return br;
        });
        for (const rl of rowLines.lines) {
          rl.top += deltaTop;
          rl.bottom += deltaTop;
          rl.column = column;
          lines.push(rl);
          block = (rl.block ?? block) + 1;
        }
        naturalY = flow.syncNaturalY();
      } else {
        for (const rl of rowLines.lines) {
          lines.push(rl);
          block = (rl.block ?? block) + 1;
        }
        naturalY = rowLines.bottom;
      }
      blocks.push({ pos, top: naturalY - rowLines.height, bottom: naturalY, kind: "table_row" });
      return false;
    }
    if (!node.isTextblock) return true;
    if (name === "paragraph" && node.attrs.frame && node.attrs.floating) {
      lines.push({
        top: naturalY,
        bottom: naturalY,
        block: block++,
        section,
        numbered: false,
        anchor: { pos, textOffset: 0 },
      });
      return false;
    }

    const format = node.attrs.format as ParagraphFormat | null;
    const pitchTwips =
      options.getDocGridLinePitchTwips?.(section) ??
      effectivePitchTwips(section, sections, format);
    const numbered = !format?.suppressLineNumbers;
    const keep = {
      widowControl: !!format?.widowControl,
      keepNext: !!format?.keepNext,
      keepLines: !!format?.keepLines,
      breakBefore: !!format?.pageBreakBefore,
      numbered,
    };
    const bid = block++;
    const geo = geometryForSection(section, sections, defaultSetup);
    const multiCol = geo.frames.length > 1;
    const flow = multiCol ? flowForSection() : null;
    const blockTop = flow ? flow.blockTop() : naturalY;
    const { lines: plines, height, tabWidths: pTabs } = layoutParagraph(
      node,
      pos,
      sectionWidth(0),
      blockTop,
      bid,
      section,
      {
        defaultTabStop: defaultTab,
        justifyShrink: options.justifyShrink ?? layout?.justifyShrink,
        pitchTwips,
        includeDeletionText: options.includeDeletionText,
      },
      keep,
    );
    tabWidths.push(...pTabs);
    if (flow) {
      const { column, deltaTop } = flow.advance(height, () => {
        const br: EngineLayoutLine = {
          top: naturalY,
          bottom: naturalY,
          breakAfter: true,
          block: block++,
          section,
          numbered: false,
          anchor: { pos, textOffset: 0 },
        };
        lines.push(br);
        return br;
      });
      for (const ln of plines) {
        ln.top += deltaTop;
        ln.bottom += deltaTop;
        ln.column = column;
        lines.push(ln);
      }
      naturalY = flow.syncNaturalY();
    } else {
      for (const ln of plines) lines.push(ln);
      naturalY += height;
    }
    blocks.push({
      pos,
      top: naturalY - height,
      bottom: naturalY,
      kind: name.startsWith("heading") ? "heading" : "paragraph",
    });
    return false;
  });

  return { lines, blocks, tabWidths };
}
