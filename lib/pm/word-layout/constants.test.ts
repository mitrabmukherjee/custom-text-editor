import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { importDocx } from "@/lib/docx/import";
import { defaultStudioLayout } from "@/lib/pm/defaultLayout";
import { LETTER_PAGE_SETUP } from "@/lib/pm/pagination/pageSetup";
import { schema } from "@/lib/pm/schema";
import { documentUsesLayoutEngine } from "./constants";

describe("documentUsesLayoutEngine", () => {
  it("returns false for studio synthetic layout", () => {
    const doc = schema.node("doc", { layout: defaultStudioLayout() }, [
      schema.node("paragraph"),
    ]);
    expect(documentUsesLayoutEngine(doc)).toBe(false);
  });

  it("returns true when layoutEngine is true", () => {
    const doc = schema.node(
      "doc",
      { layout: { page: LETTER_PAGE_SETUP, layoutEngine: true } },
      [schema.node("paragraph")],
    );
    expect(documentUsesLayoutEngine(doc)).toBe(true);
  });

  it("returns false when layoutEngine is false even with imported-like metadata", () => {
    const doc = schema.node(
      "doc",
      {
        layout: {
          page: LETTER_PAGE_SETUP,
          layoutEngine: false,
          docGridLinePitch: 360,
        },
      },
      [schema.node("paragraph")],
    );
    expect(documentUsesLayoutEngine(doc)).toBe(false);
  });

  it("uses legacy heuristic when layoutEngine is absent", () => {
    const doc = schema.node(
      "doc",
      { layout: { page: LETTER_PAGE_SETUP, docGridLinePitch: 360 } },
      [schema.node("paragraph")],
    );
    expect(documentUsesLayoutEngine(doc)).toBe(true);
  });

  it("sets layoutEngine on DOCX import", async () => {
    const path = join(process.cwd(), "public", "fixtures", "zimmerman-form3.docx");
    if (!existsSync(path)) return;
    const { doc, layout } = await importDocx(await readFile(path));
    expect(layout?.layoutEngine).toBe(true);
    expect(documentUsesLayoutEngine(schema.nodeFromJSON(doc))).toBe(true);
  });
});
