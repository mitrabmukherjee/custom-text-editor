/**
 * Node-side: engine layout vs import stats (DOM must be checked in browser with debugPagination).
 */
import { existsSync, readFileSync } from "node:fs";
import { importDocx } from "../lib/docx/import";
import { layoutDocument } from "../lib/pm/word-layout/layoutDocument";
import { lineBoxPx, lineMetricsContextFromFormat } from "../lib/pm/format";
import { schema } from "../lib/pm/schema";
import { sectionsOf } from "../lib/pm/sections";
import { EditorState } from "prosemirror-state";

const path =
  process.argv[2] ??
  "C:/Users/Steora/Downloads/file_1783367057023_Compared_MISSONS__Wayne__Dorothy_Pace_-_May_16__2024__ID1518937.3-_36_pgs_-_Shoumi_draft.docx";

async function main() {
  if (!existsSync(path)) {
    console.error("Missing:", path);
    process.exit(1);
  }
  const { doc, pageSetup, layout } = await importDocx(readFileSync(path));
  const root = schema.nodeFromJSON(doc);
  const setup = pageSetup ?? layout?.page!;
  const state = EditorState.create({ doc: root });
  const sections = sectionsOf(state);
  const pitch = sections[0]?.docGridLinePitch ?? 272;

  let paras = 0;
  let empty = 0;
  let auto15 = 0;
  let auto1 = 0;
  const boxHeights = new Map<number, number>();

  root.descendants((node) => {
    if (!node.isTextblock || node.type.name === "heading") return;
    if (!node.attrs.format?.docGridSectionPitchTwips) return;
    paras++;
    if (!node.textContent.trim()) empty++;
    const f = node.attrs.format;
    if (f.lineValue === 1.5) auto15++;
    if (f.lineValue === 1) auto1++;
    const ctx = lineMetricsContextFromFormat(f);
    const box = lineBoxPx(f, ctx);
    if (box != null) boxHeights.set(box, (boxHeights.get(box) ?? 0) + 1);
  });

  const { pagination, lines } = layoutDocument(root, { defaultSetup: setup });

  console.log("=== MISSIONS import + engine (layoutDocument) ===");
  console.log({
    docGridLinePitch: pitch,
    pitchPx: pitch / 15,
    enginePageCount: pagination.pageCount,
    engineLineCount: lines.length,
    paragraphsOnGrid: paras,
    emptyParagraphs: empty,
    auto15,
    auto1,
    lineBoxPxDistribution: Object.fromEntries([...boxHeights.entries()].sort((a, b) => a[0] - b[0])),
    sampleBreakLineIndices: pagination.breaks.slice(0, 8).map((b) => b.lineIndex),
  });
  console.log(
    "\nIn browser: localStorage.debugPagination=1, reload, import doc — console [pagination] engine vs DOM",
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
