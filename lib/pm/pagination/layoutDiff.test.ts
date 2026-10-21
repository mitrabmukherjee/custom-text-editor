import { describe, expect, it } from "vitest";
import { diffLayouts } from "./layoutDiff";

describe("diffLayouts", () => {
  it("reports page count and break index differences", () => {
    const diff = diffLayouts(
      { pageCount: 2, breaks: [{ lineIndex: 5, height: 100, forced: false }], lineCount: 10 },
      { pageCount: 3, breaks: [{ lineIndex: 6, height: 100, forced: false }], lineCount: 11 },
    );
    expect(diff.domPageCount).toBe(2);
    expect(diff.enginePageCount).toBe(3);
    expect(diff.lineCountDelta).toBe(1);
    expect(diff.firstBreakIndex).toBe(0);
  });
});
