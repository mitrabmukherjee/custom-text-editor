import { readFileSync, readdirSync } from "fs";
import path from "path";
import JSZip from "jszip";
import { importDocx } from "../lib/docx/import";
import { schema } from "../lib/pm/schema";

const dir = process.argv[2] || "public/fixtures/batch";
const maxFiles = Number(process.argv[3] ?? 50);
const files = readdirSync(dir)
  .filter((f) => f.toLowerCase().endsWith(".docx"))
  .slice(0, maxFiles);

async function main() {
let mismatches = 0;
for (const name of files) {
  const full = path.join(dir, name);
  let xml: string | undefined;
  try {
    const zip = await JSZip.loadAsync(readFileSync(full));
    xml = await zip.file("word/document.xml")?.async("string");
  } catch {
    continue;
  }
  if (!xml) continue;
  const xmlTc = (xml.match(/<w:tcBorders\b/g) || []).length;
  const xmlPb = (xml.match(/<w:pBdr\b/g) || []).length;
  const node = schema.nodeFromJSON((await importDocx(readFileSync(full))).doc);
  let cellEdges = 0;
  let paraBorders = 0;
  node.descendants((n) => {
    if (n.type.name === "table_cell" && n.attrs.edges) cellEdges += 1;
    if (n.isTextblock && n.attrs.format?.boxBorder) paraBorders += 1;
  });
  if (xmlTc > 0 && cellEdges === 0) {
    console.error("tcBorders mismatch:", name, { xmlTc, cellEdges });
    mismatches++;
  }
  if (xmlPb > 0 && paraBorders === 0) {
    console.error("pBdr mismatch:", name, { xmlPb, paraBorders });
    mismatches++;
  }
}
if (mismatches) {
  console.error(`${mismatches} file(s) failed layout import coverage`);
  process.exit(1);
}
console.log(`OK: ${files.length} file(s) checked`);
}

void main();
