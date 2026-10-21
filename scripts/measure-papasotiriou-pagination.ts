import { existsSync, readFileSync } from "node:fs";
import { importDocx } from "../lib/docx/import";
import { pageFrames } from "../lib/pm/format";
import { computePageBreaks } from "../lib/pm/pagination/computePageBreaks";
import { buildSimulatedGridLayoutLines } from "../lib/pm/pagination/simulatedGridLines";
import { sectionsOf } from "../lib/pm/sections";
import { schema } from "../lib/pm/schema";
import { EditorState } from "prosemirror-state";

const DEFAULT_PATH =
  "C:/Users/Steora/Downloads/MASTER-BATCH-2013/MASTER-BATCH-2013/TF-plus-data-v2-2013-FLAT-B2/PAPASOTIRIOU - MAY 10  2013 - ID 61861 - COMPLETE PROCEEDING - MARGARET OSADET - JG - RELEASED.docx";

const MARKERS = ["FRIDAY", "WHEREUPON", "FORM 2"] as const;

function snippet(text: string, width = 72): string {
  const s = text.replace(/\s+/g, " ").trim();
  if (!s) return "(empty)";
  return s.length > width ? `${s.slice(0, width)}…` : s;
}

function lineContext(lineTexts: string[], index: number) {
  const prev = index > 0 ? lineTexts[index - 1] : "";
  const cur = lineTexts[index] ?? "";
  const next = lineTexts[index + 1] ?? "";
  return {
    prev: snippet(prev),
    at: snippet(cur),
    next: snippet(next),
  };
}

function markerLineIndices(lineTexts: string[], marker: string): number[] {
  const out: number[] = [];
  const upper = marker.toUpperCase();
  lineTexts.forEach((t, i) => {
    if (t.toUpperCase().includes(upper)) out.push(i);
  });
  return out;
}

async function main() {
  const envPath = process.env.PAPASOTIRIOU_DOCX;
  const path = envPath && existsSync(envPath) ? envPath : DEFAULT_PATH;
  if (!existsSync(path)) {
    console.error("DOCX not found:", path);
    console.error("Set PAPASOTIRIOU_DOCX to the full file path.");
    process.exit(1);
  }

  console.log("DOCX:", path);
  console.log(
    "Note: EditorController/measure() needs a browser DOM; this script uses simulated grid lines (same as papasotiriouPagination.test.ts).\n",
  );

  const { doc, pageSetup, layout } = await importDocx(readFileSync(path));
  const root = schema.nodeFromJSON(doc);
  const sections = sectionsOf(EditorState.create({ doc: root }));
  const setup = pageSetup ?? layout?.page!;
  const frame = pageFrames(setup, layout ?? null, sections);
  const { lines, lineTexts } = buildSimulatedGridLayoutLines(root, setup, sections);
  const result = computePageBreaks(lines, setup, frame);

  console.log("totalSimulatedLines:", lines.length);
  console.log("pageCount:", result.pageCount);
  console.log("breakCount:", result.breaks.length);
  console.log("");

  for (const [bi, b] of result.breaks.entries()) {
    const ctx = lineContext(lineTexts, b.lineIndex);
    console.log(`break[${bi}] lineIndex=${b.lineIndex} height=${b.height.toFixed(2)} forced=${b.forced}`);
    console.log(`  prev: ${ctx.prev}`);
    console.log(`  at:   ${ctx.at}`);
    console.log(`  next: ${ctx.next}`);
    console.log("");
  }

  console.log("— markers in simulated line text —");
  for (const marker of MARKERS) {
    const indices = markerLineIndices(lineTexts, marker);
    console.log(`${marker}: line indices ${indices.join(", ") || "(none)"}`);
    for (const i of indices) {
      const page = result.placements[i]?.page ?? "?";
      console.log(`  line ${i} page ${page}: ${snippet(lineTexts[i])}`);
    }
  }
}

void main();
