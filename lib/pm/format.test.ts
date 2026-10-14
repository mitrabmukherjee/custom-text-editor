import { describe, expect, it } from "vitest";
import { cellBorderCss } from "./format";

describe("cellBorderCss", () => {
  it("clears default cell borders when only some edges are set", () => {
    const css = cellBorderCss({
      bottom: { style: "dotted", widthPt: 1, color: "#000000" },
    });
    expect(css).toContain("border: none !important");
    expect(css).toContain("border-bottom:");
    expect(css).not.toMatch(/border-top:/);
  });
});
