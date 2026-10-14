import { readFileSync } from "node:fs";
import { importDocx } from "../lib/docx/import.ts";
import { schema } from "../lib/pm/schema.ts";

const file = process.argv[2] ?? "public/fixtures/zimmerman-form3.docx";
const { doc } = await importDocx(readFileSync(file));
const node = schema.nodeFromJSON(doc);
let form3Table = false;
let courier = false;
let anchoredSig = false;
node.descendants((n) => {
  if (n.type.name === "table" && n.textContent.includes("FORM 3")) form3Table = true;
  if (n.type.name === "paragraph" && n.attrs.floating && n.attrs.frame) {
    const img = n.firstChild;
    if (img?.type.name === "image") anchoredSig = true;
  }
  if (n.isText && n.marks.some((m) => m.attrs?.fontFamily === "Courier New")) courier = true;
});
console.log({ form3Table, courier, anchoredSig });
