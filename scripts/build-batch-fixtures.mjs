/**
 * Build minimal batch layout fixtures under public/fixtures/batch/
 */
import fs from "fs";
import path from "path";
import JSZip from "jszip";

const OUT = path.join("public", "fixtures", "batch");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const sect =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

function doc(body) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}">
  <w:body>${body}${sect}</w:body>
</w:document>`;
}

async function writeFixture(name, bodyXml) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("word/document.xml", doc(bodyXml));
  const buf = await zip.generateAsync({ type: "nodebuffer" });
  fs.mkdirSync(OUT, { recursive: true });
  fs.writeFileSync(path.join(OUT, name), buf);
}

const form2Paragraph = `
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>FORM 2</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>CERTIFICATE OF TRANSCRIPT</w:t></w:r></w:p>`;

const form2Table = `
<w:tbl>
  <w:tblPr><w:tblStyle w:val="TableGrid"/><w:jc w:val="center"/><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="nil"/></w:tcBorders></w:tcPr>
    <w:p><w:r><w:t>FORM 2</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>Cert body</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;

const form3Snippet = `
<w:tbl>
  <w:tblPr><w:jc w:val="center"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="10896"/></w:tblGrid>
  <w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr><w:tc><w:tcPr>
    <w:tcMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tcMar>
    <w:tcBorders><w:bottom w:val="dotted" w:sz="8" w:color="000000"/></w:tcBorders>
  </w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>FORM 3</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;

await writeFixture("form2-paragraph.docx", form2Paragraph);
await writeFixture("form2-table.docx", form2Table);
await writeFixture("form3-table.docx", form3Snippet);
console.log("Wrote fixtures to", OUT);
