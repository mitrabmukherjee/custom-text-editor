/**
 * Compare Word layout primitives in a DOCX with what import preserves.
 * Usage: node scripts/layout-audit.mjs path/to/file.docx
 */
import fs from "fs";
import JSZip from "jszip";

const path = process.argv[2] || "public/test.docx";
const buffer = fs.readFileSync(path);
const zip = await JSZip.loadAsync(buffer);
const xml = await zip.file("word/document.xml")?.async("string");
if (!xml) throw new Error("missing word/document.xml");

const count = (re) => (xml.match(re) || []).length;
const snippets = (re, label, max = 3) => {
  const m = xml.match(re);
  if (!m?.length) return;
  console.log(`  sample ${label}:`, m.slice(0, max).map((s) => s.slice(0, 120)).join(" | "));
};

console.log("DOCX layout audit:", path);
console.log("— Word-only constructs (not fully rendered in editor) —");
console.log("  w:framePr (body):", count(/<w:framePr\b/g));
console.log("  w:pBdr:", count(/<w:pBdr\b/g));
console.log("  w:tblInd:", count(/<w:tblInd\b/g));
console.log("  w:tblW:", count(/<w:tblW\b/g));
console.log("  w:tbl jc:", count(/<w:jc w:val="/g));
console.log("  w:tcBorders:", count(/<w:tcBorders\b/g));
console.log("  w:drawing:", count(/<w:drawing\b/g));
console.log("  w:pict (VML):", count(/<w:pict\b/g));
console.log("  tables:", count(/<w:tbl\b/g));

snippets(/<w:framePr[^>]*\/>/g, "framePr");
snippets(/<w:tblW[^>]*\/>/g, "tblW");
snippets(/<w:pBdr>[\s\S]{0,200}/g, "pBdr");

const form = xml.match(/FORM\s*\d+/g);
if (form) console.log("  FORM markers:", [...new Set(form)].join(", "));

import { spawnSync } from "node:child_process";
spawnSync("npx", ["tsx", "scripts/layout-audit-import.ts", path], {
  stdio: "inherit",
  shell: true,
});
