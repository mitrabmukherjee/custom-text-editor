import { describe, expect, it } from "vitest";
import { schema } from "@/lib/pm/schema";
import { blockHasVisibleText } from "./visibleMarkup";
import { inlinePieces } from "@/lib/pm/word-layout/effectiveRun";

describe("visibleMarkup", () => {
  it("omits deletion mark text when includeDeletionText is false", () => {
    const { deletion } = schema.marks;
    const block = schema.nodes.paragraph.create({}, [
      schema.text("visible ", []),
      schema.text("gone", [deletion.create({ authorId: "a", authorName: "A", id: "1" })]),
    ]);
    expect(blockHasVisibleText(block, { includeDeletionText: true })).toBe(true);
    expect(blockHasVisibleText(block, { includeDeletionText: false })).toBe(true);
    const pieces = inlinePieces(block, 0, { includeDeletionText: false });
    expect(pieces.map((p) => (p.kind === "text" ? p.text : "")).join("")).toBe("visible ");
  });
});
