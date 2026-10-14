import { readFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";

const file = process.argv[2];
if (!file) {
  console.error("usage: node scripts/peek-docx.mjs <path.docx>");
  process.exit(1);
}

const zip = await JSZip.loadAsync(readFileSync(file));
const xml = await zip.file("word/document.xml")?.async("string");
if (!xml) throw new Error("no document.xml");

const plain = xml.replace(/<w:tab[^/]*\/>/g, "\t").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
console.log("plain (first 800):", plain.slice(0, 800));
console.log({
  tables: (xml.match(/<w:tbl/g) || []).length,
  pBdr: (xml.match(/pBdr/g) || []).length,
  framePr: (xml.match(/framePr/g) || []).length,
  pict: (xml.match(/<w:pict/g) || []).length,
  drawing: (xml.match(/<w:drawing/g) || []).length,
  form3: /FORM\s*3/i.test(xml),
  certify: /certify/i.test(xml),
});
