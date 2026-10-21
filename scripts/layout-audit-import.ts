import { readFileSync } from "fs";
import { importDocx } from "../lib/docx/import";
import {
  hfSetParagraphs,
  isHFTable,
  lineBoxPx,
  lineMetricsContextFromFormat,
  paragraphCss,
} from "../lib/pm/format";
import { sectionsOf } from "../lib/pm/sections";
import { schema } from "../lib/pm/schema";
import { EditorState } from "prosemirror-state";

const path = process.argv[2] || "public/test.docx";

async function main() {
  const result = await importDocx(readFileSync(path));
  const node = schema.nodeFromJSON(result.doc);
  const state = EditorState.create({ doc: node });

  let tables = 0;
  let tablesWithWidth = 0;
  let tablesWithIndent = 0;
  let tablesCentered = 0;
  let paraBorders = 0;
  let cellEdges = 0;
  let images = 0;
  let anchoredImages = 0;
  let vmlAnchors = 0;
  let framed = 0;
  let hfTableBlocks = 0;
  let hfBoxBorderParas = 0;

  node.descendants((n) => {
    if (n.type.name === "table") {
      tables += 1;
      if (n.attrs.width || n.attrs.widthPct) tablesWithWidth += 1;
      if (n.attrs.indentLeft || n.attrs.indentRight) tablesWithIndent += 1;
      if (n.attrs.align === "center") tablesCentered += 1;
    }
    if (n.type.name === "table_cell" && n.attrs.edges) cellEdges += 1;
    if (n.isTextblock && n.attrs.format?.boxBorder) paraBorders += 1;
    if (n.type.name === "image") {
      images += 1;
      if (n.attrs.anchorX != null || n.attrs.anchorY != null) anchoredImages += 1;
      if (n.attrs.vmlAnchor) vmlAnchors += 1;
    }
    if (n.type.name === "paragraph" && n.attrs.frame) framed += 1;
  });

  const sections = sectionsOf(state);
  const docLayout = result.layout;
  console.log("— After import (document model) —");
  console.log("  tables:", tables, `(width: ${tablesWithWidth}, indent: ${tablesWithIndent}, centered: ${tablesCentered})`);
  console.log("  cells with tcBorders:", cellEdges);
  console.log("  paragraphs with boxBorder:", paraBorders);
  console.log("  images:", images, `(anchored: ${anchoredImages}, vmlAnchor: ${vmlAnchors})`);
  console.log("  body frame paragraphs:", framed);
  console.log("  doc layout pageNumberStart:", docLayout?.pageNumberStart ?? "(none)");
  console.log("  doc layout docGridLinePitch:", docLayout?.docGridLinePitch ?? "(none)");
  if (docLayout?.page) {
    const p = docLayout.page;
    console.log(
      "  page margins (px):",
      `T${p.marginTop} R${p.marginRight} B${p.marginBottom} L${p.marginLeft}`,
    );
  }
  for (const [i, sec] of sections.entries()) {
    if (sec.pageNumberStart != null) console.log(`  section[${i}] pageNumberStart:`, sec.pageNumberStart);
    if (sec.docGridLinePitch != null) {
      console.log(`  section[${i}] docGridLinePitch:`, sec.docGridLinePitch);
    }
    if (sec.lineNumbers) {
      console.log(
        `  section[${i}] lineNumbers: countBy=${sec.lineNumbers.countBy} start=${sec.lineNumbers.start}`,
      );
    }
    for (const key of ["default", "first", "even"] as const) {
      const blocks = sec.headers?.[key];
      if (!blocks?.length) continue;
      for (const block of blocks) {
        if (isHFTable(block)) hfTableBlocks += 1;
        hfBoxBorderParas += hfSetParagraphs([block]).filter((p) => p.format.boxBorder).length;
      }
    }
  }
  console.log("  header table blocks:", hfTableBlocks);
  console.log("  HF paragraphs with boxBorder:", hfBoxBorderParas);

  let sampleGridPara: string | null = null;
  node.descendants((n) => {
    if (sampleGridPara || !n.isTextblock) return;
    const f = n.attrs.format as { docGridSectionPitchTwips?: number } | null;
    if (!f?.docGridSectionPitchTwips) return;
    sampleGridPara = paragraphCss(f);
    console.log(
      "  sample grid paragraph lineBoxPx:",
      lineBoxPx(f, lineMetricsContextFromFormat(f)),
    );
    console.log("  sample grid paragraphCss:", sampleGridPara);
  });
}

void main();
