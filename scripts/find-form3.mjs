import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";

const dir = process.argv[2] ?? "C:/Users/Steora/Downloads";
for (const name of readdirSync(dir)) {
  if (!name.toLowerCase().endsWith(".docx")) continue;
  const path = join(dir, name);
  try {
    const zip = await JSZip.loadAsync(readFileSync(path));
    const xml = await zip.file("word/document.xml")?.async("string");
    if (!xml) continue;
    if (/FORM\s*3/i.test(xml) && /certify/i.test(xml)) {
      console.log(path);
    }
  } catch {
    /* skip */
  }
}
