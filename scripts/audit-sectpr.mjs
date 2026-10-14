import fs from "node:fs";
import { DOMParser } from "@xmldom/xmldom";
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const path = process.argv[2] || "public/test.docx";

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
function attr(el, name) {
  if (!el) return null;
  return el.getAttributeNS(W, name) || el.getAttribute(`w:${name}`);
}

function pageKey(sectPr) {
  const pgSz = child(sectPr, "pgSz");
  const cols = child(sectPr, "cols");
  return [attr(pgSz, "w"), attr(pgSz, "h"), attr(cols, "num") || "1"].join("|");
}

function visitBody(el, fn) {
  for (const c of elements(el)) {
    if (c.localName === "p") {
      const sectPr = child(child(c, "pPr"), "sectPr");
      if (sectPr) fn(sectPr);
    } else if (c.localName === "sdt") {
      const content = child(c, "sdtContent");
      if (content) visitBody(content, fn);
    } else if (c.localName === "tbl") {
      for (const tr of elements(c).filter((r) => r.localName === "tr")) {
        for (const tc of elements(tr).filter((x) => x.localName === "tc")) visitBody(tc, fn);
      }
    }
  }
}

function inferStart(inline) {
  let prev = null;
  return inline.map((s) => {
    let start;
    if (s.type === "nextColumn" && (!s.cols || Number(s.cols) <= 1)) start = "continuous";
    else if (s.type) start = s.type;
    else if (prev === null || s.pageKey === prev) start = "continuous";
    else start = "nextPage";
    prev = s.pageKey;
    return start;
  });
}

const buffer = fs.readFileSync(path);
const zip = await JSZip.loadAsync(buffer);
const xml = await zip.file("word/document.xml").async("string");
const doc = new DOMParser().parseFromString(xml, "application/xml");
const body = elements(doc.documentElement).find((c) => c.localName === "body");
const inline = [];
visitBody(body, (sectPr) => {
  const typeEl = child(sectPr, "type");
  const cols = child(sectPr, "cols");
  inline.push({
    type: attr(typeEl, "val"),
    cols: attr(cols, "num"),
    pageKey: pageKey(sectPr),
  });
});
const starts = inferStart(inline);
const hist = {};
for (const s of starts) hist[s] = (hist[s] ?? 0) + 1;
console.log({ file: path, inline: inline.length, hist, noType: inline.filter((s) => !s.type).length });
