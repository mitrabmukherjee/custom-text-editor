import JSZip from "jszip";

/** Inject `w:docGrid w:linePitch` into each `w:sectPr` (docx library does not emit doc grid). */
export async function patchDocGridLinePitch(
  buffer: ArrayBuffer,
  pitches: (number | undefined)[],
): Promise<ArrayBuffer> {
  if (!pitches.some((p) => p !== undefined)) return buffer;

  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) return buffer;

  let docXml = await file.async("string");
  let idx = 0;
  docXml = docXml.replace(/<w:sectPr>([\s\S]*?)<\/w:sectPr>/g, (full, inner) => {
    const pitch = pitches[idx++];
    if (!pitch || inner.includes("<w:docGrid")) return full;
    return `<w:sectPr>${inner}<w:docGrid w:linePitch="${pitch}" w:charSpace="0" w:type="lines"/></w:sectPr>`;
  });

  zip.file("word/document.xml", docXml);
  return zip.generateAsync({ type: "arraybuffer" });
}
