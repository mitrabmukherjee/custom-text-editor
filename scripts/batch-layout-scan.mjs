/**
 * Scan a folder of DOCX for layout primitives (certification / batch QA).
 * Usage: node scripts/batch-layout-scan.mjs "C:/path/to/FLAT-B2"
 */
import fs from "fs";
import path from "path";
import JSZip from "jszip";

const dir = process.argv[2];
if (!dir) {
  console.error("Usage: node scripts/batch-layout-scan.mjs <folder>");
  process.exit(1);
}

const files = fs.readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".docx"));
const totals = {
  files: 0,
  skipped: 0,
  form2: 0,
  form3: 0,
  pBdr: 0,
  framePr: 0,
  tcBorders: 0,
  tblW: 0,
  sectPr: 0,
  pageBreakBefore: 0,
};
const hot = [];

for (const name of files) {
  let xml;
  try {
    const buf = fs.readFileSync(path.join(dir, name));
    const zip = await JSZip.loadAsync(buf);
    xml = await zip.file("word/document.xml")?.async("string");
  } catch {
    totals.skipped += 1;
    continue;
  }
  if (!xml) {
    totals.skipped += 1;
    continue;
  }
  totals.files += 1;
  const form2 = /FORM\s*2/i.test(xml);
  const form3 = /FORM\s*3/i.test(xml);
  if (form2) totals.form2 += 1;
  if (form3) totals.form3 += 1;
  const pBdr = (xml.match(/<w:pBdr\b/g) || []).length;
  const framePr = (xml.match(/<w:framePr\b/g) || []).length;
  const tcBorders = (xml.match(/<w:tcBorders\b/g) || []).length;
  const tblW = (xml.match(/<w:tblW\b/g) || []).length;
  const sectPr = (xml.match(/<w:sectPr\b/g) || []).length;
  const pageBreakBefore = (xml.match(/<w:pageBreakBefore\b/g) || []).length;
  totals.pBdr += pBdr;
  totals.framePr += framePr;
  totals.tcBorders += tcBorders;
  totals.tblW += tblW;
  totals.sectPr += sectPr;
  totals.pageBreakBefore += pageBreakBefore;
  const score = framePr + tcBorders + (form2 || form3 ? 5 : 0);
  if (score >= 8 || (form3 && framePr > 0)) {
    hot.push({ name, form2, form3, pBdr, framePr, tcBorders, tblW, sectPr, pageBreakBefore });
  }
}

hot.sort((a, b) => b.framePr + b.tcBorders - (a.framePr + a.tcBorders));
console.log(JSON.stringify({ totals, hotSample: hot.slice(0, 12) }, null, 2));
