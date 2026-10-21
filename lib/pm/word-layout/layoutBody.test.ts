import { describe, expect, it } from "vitest";
import { schema } from "@/lib/pm/schema";
import { layoutDocument } from "./layoutDocument";
import { LETTER_PAGE_SETUP } from "../pagination/pageSetup";

describe("layoutBody", () => {
  it("lays out a simple paragraph into at least one line", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(
        { format: { fontSize: 12, lineRule: "auto", lineValue: 1 } },
        [schema.text("Hello world from the word layout engine.")],
      ),
    ]);
    const { lines, pagination } = layoutDocument(doc, { defaultSetup: LETTER_PAGE_SETUP });
    expect(lines.length).toBeGreaterThan(0);
    expect(pagination.pageCount).toBeGreaterThanOrEqual(1);
    expect(lines[0].anchor?.pos).toBe(0);
  });
});
