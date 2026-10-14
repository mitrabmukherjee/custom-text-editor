import fs from "node:fs";

const { importDocx } = await import("../lib/docx/import.ts");

const buffer = fs.readFileSync("public/test.docx");
const { doc } = await importDocx(buffer);

function walk(nodes, fn) {
  for (const node of nodes ?? []) {
    fn(node);
    walk(node.content, fn);
  }
}

let pageBreaks = 0;
let sectionBreaks = 0;
let pageBreakBefore = 0;
walk(doc.content, (node) => {
  if (node.type === "page_break") pageBreaks += 1;
  if (node.type === "section_break") sectionBreaks += 1;
  if ((node.type === "paragraph" || node.type === "heading") && node.attrs?.format?.pageBreakBefore) {
    pageBreakBefore += 1;
  }
});

console.log({ pageBreaks, sectionBreaks, pageBreakBefore });
