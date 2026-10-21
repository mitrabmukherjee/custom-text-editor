import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { importDocx } from "./import";
import { lineBoxPx, lineMetricsContextFromFormat, paragraphCss } from "@/lib/pm/format";
import { sectionsOf } from "@/lib/pm/sections";
import { EditorState } from "prosemirror-state";
import { schema } from "@/lib/pm/schema";

const PAPASOTIRIOU = join(
  process.cwd(),
  "public",
  "fixtures",
  "papasotiriou-may-2013.docx",
);

describe("PAPASOTIRIOU grid line spacing", () => {
  it("maps 12pt Courier 1.5× in a 360-twips grid to 24px line boxes", async () => {
    const envPath = process.env.PAPASOTIRIOU_DOCX;
    const path = envPath && existsSync(envPath) ? envPath : PAPASOTIRIOU;
    if (!existsSync(path)) return;

    const { doc } = await importDocx(await readFile(path));
    const root = schema.nodeFromJSON(doc);
    let found = false;
    root.descendants((n) => {
      if (found || !n.isTextblock) return;
      const f = n.attrs.format as {
        fontSize?: number;
        lineRule?: string;
        lineValue?: number;
        docGridSectionPitchTwips?: number;
      } | null;
      if (f?.docGridSectionPitchTwips !== 360) return;
      if (f.lineRule !== "auto" || f.lineValue !== 1.5) return;
      if (f.fontSize !== 12) return;
      expect(lineBoxPx(f as import("@/lib/pm/format").ParagraphFormat, lineMetricsContextFromFormat(f as import("@/lib/pm/format").ParagraphFormat))).toBe(24);
      found = true;
    });
    expect(found).toBe(true);
  });

  it("section 0 has no grid stamp; grid sections use px line-height in paragraphCss", async () => {
    const envPath = process.env.PAPASOTIRIOU_DOCX;
    const path = envPath && existsSync(envPath) ? envPath : PAPASOTIRIOU;
    if (!existsSync(path)) return;

    const { doc } = await importDocx(await readFile(path));
    const root = schema.nodeFromJSON(doc);
    const sections = sectionsOf(EditorState.create({ doc: root }));
    expect(sections[0]?.docGridLinePitch).toBeUndefined();

    let beforeGrid = true;
    let gridCss: string | null = null;
    root.forEach((child, _offset, index) => {
      if (child.type.name === "section_break") beforeGrid = false;
      if (!child.isTextblock || gridCss) return;
      const f = child.attrs.format as { docGridSectionPitchTwips?: number } | null;
      if (beforeGrid) {
        expect(f?.docGridSectionPitchTwips).toBeUndefined();
        return;
      }
      if (f?.docGridSectionPitchTwips === 360) {
        gridCss = paragraphCss(f);
      }
    });
    expect(gridCss).toMatch(/line-height:\s*(16|24)px/);
  });
});
