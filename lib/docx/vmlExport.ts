import JSZip from "jszip";
import type { DocJSON, NodeJSON } from "@/lib/pm/json";
import { parseVmlAnchorAttr, type VmlAnchorJSON } from "./vmlAnchor";

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const VML_NS = "urn:schemas-microsoft-com:vml";
const O_NS = "urn:schemas-microsoft-com:office:office";

type VmlSig = { anchor: VmlAnchorJSON; src: string };

function collectVmlSignatures(json: DocJSON): VmlSig[] {
  const out: VmlSig[] = [];
  const visit = (nodes: NodeJSON[] | undefined) => {
    for (const node of nodes ?? []) {
      if (node.type === "paragraph" || node.type === "heading") {
        for (const child of node.content ?? []) {
          if (child.type !== "image") continue;
          const anchor = parseVmlAnchorAttr(child.attrs?.vmlAnchor);
          const src = String(child.attrs?.src ?? "");
          if (anchor && src.startsWith("data:")) out.push({ anchor, src });
        }
      }
      visit(node.content);
    }
  };
  visit(json.content);
  return out;
}

function extFromMime(mime: string) {
  if (mime.includes("png")) return "png";
  if (mime.includes("jpeg") || mime.includes("jpg")) return "jpeg";
  if (mime.includes("gif")) return "gif";
  return "png";
}

function decodeDataUrl(src: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;]+);base64,(.+)$/.exec(src);
  if (!m) return null;
  const b64 = m[2];
  const bytes =
    typeof Buffer !== "undefined"
      ? new Uint8Array(Buffer.from(b64, "base64"))
      : Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  return { mime: m[1], bytes };
}

function pictParagraph(relId: string, anchor: VmlAnchorJSON): string {
  const { marginLeftPt, marginTopPt, widthPt, heightPt } = anchor;
  const ind =
    anchor.carrierIndentLeftPx && anchor.carrierIndentLeftPx > 0
      ? `<w:ind w:left="${Math.round(anchor.carrierIndentLeftPx * 15)}"/>`
      : "";
  return (
    `<w:p><w:pPr><w:widowControl w:val="0"/>${ind}</w:pPr>` +
    `<w:r><w:rPr><w:noProof/></w:rPr>` +
    `<w:pict>` +
    `<v:shape style="position:absolute;margin-left:${marginLeftPt}pt;margin-top:${marginTopPt}pt;width:${widthPt}pt;height:${heightPt}pt">` +
    `<v:imagedata r:id="${relId}" o:title=""/>` +
    `</v:shape></w:pict></w:r></w:p>`
  );
}

function paragraphText(block: string): string {
  return [...block.matchAll(/<w:t[^>]*>([^<]*)</g)].map((m) => m[1]).join("");
}

/** Start offsets of each `w:p` in document order. */
function paragraphStarts(xml: string): number[] {
  return [...xml.matchAll(/<w:p[\s>]/g)].map((m) => m.index!);
}

function findSignatureUnderlineStart(xml: string): number | null {
  for (const m of xml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)) {
    const block = m[0];
    if (!/_{8,}/.test(paragraphText(block))) continue;
    if (block.includes("w:drawing") || block.includes("w:pict")) return m.index ?? null;
    return m.index ?? null;
  }
  return null;
}

function insertCarrierBeforeUnderline(docXml: string, carrier: string, paragraphsBetween: number): string {
  const underlineAt = findSignatureUnderlineStart(docXml);
  if (underlineAt === null) return docXml;

  const starts = paragraphStarts(docXml);
  const underlineIdx = starts.findIndex((s) => s === underlineAt);
  const targetIdx = Math.max(0, underlineIdx - paragraphsBetween);
  const insertAt = starts[targetIdx] ?? underlineAt;

  let xml = docXml.slice(0, insertAt) + carrier + docXml.slice(insertAt);

  const afterCarrier = insertAt + carrier.length;
  const slice = xml.slice(afterCarrier, afterCarrier + 12000);
  const drawingRe = /<w:drawing>[\s\S]*?<\/w:drawing>/;
  const d = drawingRe.exec(slice);
  if (d) {
    const absStart = afterCarrier + d.index;
    xml = xml.slice(0, absStart) + xml.slice(absStart + d[0].length);
  }
  return xml;
}

/** Rebuild Zimmerman-style VML carrier paragraphs and strip inline drawings from signature lines. */
export async function patchVmlSignatures(buffer: ArrayBuffer, json: DocJSON): Promise<ArrayBuffer> {
  const sigs = collectVmlSignatures(json);
  if (!sigs.length) return buffer;

  const zip = await JSZip.loadAsync(buffer);
  let docXml = await zip.file("word/document.xml")!.async("string");
  let relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string");

  let nextId = 100;
  const existingIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  if (existingIds.length) nextId = Math.max(...existingIds) + 1;

  for (const sig of sigs) {
    const data = decodeDataUrl(sig.src);
    if (!data) continue;
    const relId = `rId${nextId++}`;
    const ext = extFromMime(data.mime);
    const mediaPath = `word/media/vml-sig-${relId}.${ext}`;
    zip.file(mediaPath, data.bytes);
    relsXml = relsXml.replace(
      "</Relationships>",
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/vml-sig-${relId}.${ext}"/></Relationships>`,
    );

    const carrier = pictParagraph(relId, sig.anchor);
    docXml = insertCarrierBeforeUnderline(docXml, carrier, sig.anchor.paragraphsBetween);
  }

  if (!docXml.includes(`xmlns:v="${VML_NS}"`)) {
    docXml = docXml.replace(
      /<w:document([^>]*)>/,
      `<w:document$1 xmlns:v="${VML_NS}" xmlns:o="${O_NS}" xmlns:r="${R_NS}">`,
    );
  }

  zip.file("word/document.xml", docXml);
  zip.file("word/_rels/document.xml.rels", relsXml);
  return zip.generateAsync({ type: "arraybuffer" });
}
