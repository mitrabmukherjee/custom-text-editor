import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { importDocx } from "../lib/docx/import";
import { schema } from "../lib/pm/schema";
import { layoutDocument } from "../lib/pm/word-layout/layoutDocument";
import { hashLineAnchors, layoutGoldenFrom } from "../lib/pm/word-layout/golden";
import type { EngineLayoutLine, LayoutGolden } from "../lib/pm/word-layout/types";

const DEFAULT_PATH =
  "C:/Users/Steora/Downloads/MASTER-BATCH-2013/MASTER-BATCH-2013/TF-plus-data-v2-2013-FLAT-B2/PAPASOTIRIOU - MAY 10  2013 - ID 61861 - COMPLETE PROCEEDING - MARGARET OSADET - JG - RELEASED.docx";

const FIXTURE_NAME = "papasotiriou-may-2013.layout-golden.json";
const OUT = path.join(process.cwd(), "public", "fixtures", FIXTURE_NAME);

function pages12Golden(
  lines: EngineLayoutLine[],
  pagination: ReturnType<typeof layoutDocument>["pagination"],
): LayoutGolden["pages12"] {
  const perPageLines = pagination.placements.reduce<number[]>((acc, p) => {
    if (p.page <= 1) acc[p.page] = (acc[p.page] ?? 0) + 1;
    return acc;
  }, []);
  const page0Lines = lines.filter((_, i) => pagination.placements[i]?.page === 0);
  const page1Lines = lines.filter((_, i) => pagination.placements[i]?.page === 1);
  const breaksThroughPage2 = pagination.breaks
    .map((b, i) => ({ i, lineIndex: b.lineIndex }))
    .filter(({ lineIndex }) => (pagination.placements[lineIndex]?.page ?? 0) <= 1)
    .map((b) => b.lineIndex);
  return {
    pageCount: pagination.pageCount,
    perPageLines: [perPageLines[0] ?? 0, perPageLines[1] ?? 0],
    breakLineIndices: breaksThroughPage2,
    lineAnchorHashPages12: hashLineAnchors([...page0Lines, ...page1Lines]),
  };
}

async function main() {
  const envPath = process.env.PAPASOTIRIOU_DOCX;
  const docxPath = envPath && existsSync(envPath) ? envPath : DEFAULT_PATH;
  if (!existsSync(docxPath)) {
    console.error("DOCX not found:", docxPath);
    console.error("Set PAPASOTIRIOU_DOCX to the full file path.");
    process.exit(1);
  }

  const { doc, pageSetup, layout } = await importDocx(readFileSync(docxPath));
  const node = schema.nodeFromJSON(doc);
  const result = layoutDocument(node, { defaultSetup: pageSetup ?? layout?.page });
  const golden: LayoutGolden = {
    ...layoutGoldenFrom(FIXTURE_NAME, result.lines, result.pagination),
    pages12: pages12Golden(result.lines, result.pagination),
  };
  writeFileSync(OUT, JSON.stringify(golden, null, 2) + "\n");
  console.log("wrote", OUT);
  console.log("pageCount:", golden.pageCount, "lineCount:", golden.lineCount);
  console.log("pages12:", golden.pages12);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
