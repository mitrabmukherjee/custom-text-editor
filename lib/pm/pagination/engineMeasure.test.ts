import { describe, expect, it } from "vitest";
import { schema } from "../schema";
import { measureWithEngine } from "./engineMeasure";
import { LETTER_PAGE_SETUP } from "./pageSetup";

describe("measureWithEngine", () => {
  it("lays out a simple Word paragraph", () => {
    const doc = schema.node("doc", { layout: { page: LETTER_PAGE_SETUP, styles: [] } }, [
      schema.node("paragraph", {
        format: { align: "left", fontFamily: "Times New Roman", fontSize: 12 },
      }),
    ]);
    const result = measureWithEngine({
      doc,
      setup: LETTER_PAGE_SETUP,
      defaultTabStop: 48,
    });
    expect(result.lines.length).toBeGreaterThan(0);
    expect(result.layout.pageCount).toBeGreaterThanOrEqual(1);
    expect(result.lines[0].ref.kind).toBe("engine");
  });
});
