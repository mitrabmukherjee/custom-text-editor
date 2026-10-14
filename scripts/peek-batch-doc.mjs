import fs from "fs";
import path from "path";
import JSZip from "jszip";

const dir =
  process.argv[2] ||
  "C:/Users/Steora/Downloads/MASTER-BATCH-2013/MASTER-BATCH-2013/TF-plus-data-v2-2013-FLAT-B2";
const pick =
  process.argv[3] ||
  "JONAT Henning - April 4 2013 - ID 94661 - 5 PGS - COMPLETED - AV - RELEASED.docx";

const file = path.join(dir, pick);
const xml = await (await JSZip.loadAsync(fs.readFileSync(file)))
  .file("word/document.xml")
  .async("string");

console.log("file:", pick);
console.log("tcBorders:", (xml.match(/<w:tcBorders/g) || []).length);
const block = xml.match(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/);
console.log(block?.[0]?.slice(0, 600) ?? "(none)");
const tag = process.argv[4] || "FORM 2";
const re = new RegExp(tag.replace(/\s+/g, "\\s*"), "i");
const i = xml.search(re);
console.log(`\n--- ${tag} region (index ${i}) ---\n`);
console.log(xml.slice(i, i + 2000).replace(/></g, ">\n<"));
