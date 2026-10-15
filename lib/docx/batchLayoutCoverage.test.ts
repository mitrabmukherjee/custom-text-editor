import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import path from "path";
import JSZip from "jszip";
import { importDocx } from "./import";
import { schema } from "@/lib/pm/schema";

const dir = path.join(process.cwd(), "public", "fixtures", "batch");

describe("batch layout import coverage", () => {
  it("preserves tcBorders and pBdr on batch fixtures", async () => {
    const files = readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".docx"));
    expect(files.length).toBeGreaterThan(0);
    let mismatches = 0;
    for (const name of files) {
      const full = path.join(dir, name);
      const zip = await JSZip.loadAsync(readFileSync(full));
      const xml = await zip.file("word/document.xml")?.async("string");
      if (!xml) continue;
      const xmlTc = (xml.match(/<w:tcBorders\b/g) || []).length;
      const xmlPb = (xml.match(/<w:pBdr\b/g) || []).length;
      const node = schema.nodeFromJSON((await importDocx(readFileSync(full))).doc);
      let cellEdges = 0;
      let paraBorders = 0;
      node.descendants((n) => {
        if (n.type.name === "table_cell" && n.attrs.edges) cellEdges += 1;
        if (n.isTextblock && n.attrs.format?.boxBorder) paraBorders += 1;
      });
      if (xmlTc > 0 && cellEdges === 0) mismatches += 1;
      if (xmlPb > 0 && paraBorders === 0) mismatches += 1;
    }
    expect(mismatches).toBe(0);
  });
});
