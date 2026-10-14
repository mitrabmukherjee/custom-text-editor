import { readFileSync } from "fs";
import { importDocx } from "../lib/docx/import";
import { schema } from "../lib/pm/schema";

const path = process.argv[2] || "public/test.docx";

async function main() {
  const result = await importDocx(readFileSync(path));
  const node = schema.nodeFromJSON(result.doc);

  let tables = 0;
  let tablesWithWidth = 0;
  let tablesCentered = 0;
  let paraBorders = 0;
  let cellEdges = 0;
  let images = 0;
  let framed = 0;

  node.descendants((n) => {
    if (n.type.name === "table") {
      tables += 1;
      if (n.attrs.width || n.attrs.widthPct) tablesWithWidth += 1;
      if (n.attrs.align === "center") tablesCentered += 1;
    }
    if (n.type.name === "table_cell" && n.attrs.edges) cellEdges += 1;
    if (n.isTextblock && n.attrs.format?.boxBorder) paraBorders += 1;
    if (n.type.name === "image") images += 1;
    if (n.type.name === "paragraph" && n.attrs.frame) framed += 1;
  });

  console.log("— After import (document model) —");
  console.log("  tables:", tables, `(width: ${tablesWithWidth}, centered: ${tablesCentered})`);
  console.log("  cells with tcBorders:", cellEdges);
  console.log("  paragraphs with boxBorder:", paraBorders);
  console.log("  images:", images);
  console.log("  body frame paragraphs:", framed);
}

void main();
