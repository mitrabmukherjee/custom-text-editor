import { readFileSync, writeFileSync } from "node:fs";
import JSZip from "jszip";

const file = process.argv[2];
const zip = await JSZip.loadAsync(readFileSync(file));
const xml = await zip.file("word/document.xml")?.async("string");
if (!xml) throw new Error("no document.xml");
const i = xml.search(/FORM\s*3/i);
if (i < 0) {
  console.error("FORM 3 not found");
  process.exit(1);
}
const slice = xml.slice(Math.max(0, i - 2000), i + 25000);
writeFileSync("scripts/.form3-snippet.xml", slice);
console.log("wrote scripts/.form3-snippet.xml", slice.length, "chars from index", i);

const after = xml.slice(i);
const counts = {
  tbl: (after.match(/<w:tbl/g) || []).length,
  pBdr: (after.match(/pBdr/g) || []).length,
  framePr: (after.match(/framePr/g) || []).length,
  pict: (after.match(/<w:pict/g) || []).length,
  drawing: (after.match(/<w:drawing/g) || []).length,
  tcBorders: (after.match(/tcBorders/g) || []).length,
};
console.log("after FORM 3:", counts);
