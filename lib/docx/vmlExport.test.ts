import { describe, expect, it } from "vitest";
import type { DocJSON } from "@/lib/pm/json";
import { patchVmlSignatures } from "./vmlExport";
import JSZip from "jszip";

const MIN_DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Date line</w:t></w:r></w:p>
<w:p><w:r><w:t>________________________</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`;

describe("patchVmlSignatures", () => {
  it("inserts w:pict carrier before the signature underline paragraph", async () => {
    const zip = new JSZip();
    zip.file("word/document.xml", MIN_DOC);
    zip.file(
      "word/_rels/document.xml.rels",
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    );
    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const json: DocJSON = {
      type: "doc",
      content: [
        { type: "paragraph", content: [{ type: "text", text: "Date line" }] },
        {
          type: "paragraph",
          content: [
            {
              type: "image",
              attrs: {
                src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
                alt: "",
                title: null,
                width: 10,
                height: 10,
                anchorX: 0,
                anchorY: 0,
                vmlAnchor: {
                  placement: "underlineParagraph",
                  marginLeftPt: 100,
                  marginTopPt: 2,
                  widthPt: 50,
                  heightPt: 20,
                  paragraphsBetween: 1,
                },
              },
            },
            { type: "text", text: "________________________" },
          ],
        },
      ],
    };
    const out = await patchVmlSignatures(buffer, json);
    const xml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:pict[\s\S]*v:shape/i);
    expect(xml.indexOf("<w:pict")).toBeLessThan(xml.indexOf("________________________"));
  });
});
