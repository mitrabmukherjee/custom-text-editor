import { describe, expect, it } from "vitest";
import { schema } from "@/lib/pm/schema";
import { inlineDocPos } from "./positions";

describe("inlineDocPos", () => {
  it("maps text offsets through tabs and text", () => {
    const block = schema.nodes.paragraph.create(null, [
      schema.text("ab"),
      schema.nodes.tab.create(),
      schema.text("cd"),
    ]);
    const pos = 0;
    expect(inlineDocPos(pos, block, 0)).toBe(pos + 1);
    expect(inlineDocPos(pos, block, 2)).toBe(pos + 3);
    expect(inlineDocPos(pos, block, 3)).toBe(pos + 4);
    expect(inlineDocPos(pos, block, 4)).toBe(pos + 5);
  });
});
