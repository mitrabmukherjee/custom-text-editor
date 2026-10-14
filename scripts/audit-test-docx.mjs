import fs from "node:fs";
import { DOMParser } from "@xmldom/xmldom";
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function elements(el) {
  const out = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i];
    if (n.nodeType === 1) out.push(n);
  }
  return out;
}

function child(el, name) {
  return elements(el).find((c) => c.localName === name);
}

function visitBody(el, fn) {
  for (const c of elements(el)) {
    if (c.localName === "p") {
      const pPr = child(c, "pPr");
      const sectPr = pPr ? child(pPr, "sectPr") : null;
      if (sectPr) fn(sectPr, "p");
    } else if (c.localName === "sdt") {
      const content = child(c, "sdtContent");
      if (content) visitBody(content, fn);
    } else if (c.localName === "tbl") {
      for (const tr of elements(c).filter((r) => r.localName === "tr")) {
        for (const tc of elements(tr).filter((x) => x.localName === "tc")) {
          visitBody(tc, fn);
        }
      }
    } else if (c.localName === "sectPr") {
      fn(c, "body");
    }
  }
}

const buffer = fs.readFileSync("public/test.docx");
const zip = await JSZip.loadAsync(buffer);
const xml = await zip.file("word/document.xml").async("string");
const doc = new DOMParser().parseFromString(xml, "text/xml");
const body = elements(doc.documentElement).find((c) => c.localName === "body");
const inline = [];
visitBody(body, (sectPr, kind) => {
  const type = sectPr.getAttributeNS(W, "val") || sectPr.getAttribute("w:val");
  const typeEl = elements(sectPr).find((c) => c.localName === "type");
  const typeVal = typeEl?.getAttributeNS(W, "val") || typeEl?.getAttribute("w:val");
  inline.push({ kind, typeVal: typeVal ?? "(none)" });
});
console.log("sectPr nodes:", inline.length);
console.log(JSON.stringify(inline, null, 2));
const pageBr = (xml.match(/w:br[^>]*w:type="page"/g) || []).length;
const pbBefore = (xml.match(/w:pageBreakBefore/g) || []).length;
console.log("w:br page:", pageBr, "pageBreakBefore:", pbBefore);
