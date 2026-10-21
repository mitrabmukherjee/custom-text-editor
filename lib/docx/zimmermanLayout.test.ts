import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { importDocx } from "./import";
import { schema } from "@/lib/pm/schema";
import { sectionsOf } from "@/lib/pm/sections";

describe("Zimmerman Form 3 layout metadata", () => {
  it("stores section pageNumberStart 20 and docGrid line pitch on the Form 3 section", async () => {
    const path = join(process.cwd(), "public", "fixtures", "zimmerman-form3.docx");
    if (!existsSync(path)) return;
    const { doc, layout } = await importDocx(await readFile(path));
    const sections = sectionsOf(EditorState.create({ doc: schema.nodeFromJSON(doc) }));
    expect(sections.some((s) => s.pageNumberStart === 20)).toBe(true);
    expect(sections.some((s) => s.docGridLinePitch === 360)).toBe(true);
    expect(layout?.docGridLinePitch).toBeDefined();
    expect(sections.some((s) => s.lineNumbers?.countBy === 5)).toBe(true);
  });
});
