import { describe, expect, it } from "vitest";
import {
  cellBorderCss,
  lineBoxPx,
  lineBoxSlack,
  lineHeightCss,
  resolveLineMetrics,
  singleLinePx,
} from "./format";

const GRID_360 = { docGridLinePitchTwips: 360 };
const GRID_272 = { docGridLinePitchTwips: 272 };

describe("line metrics", () => {
  it("uses 1.15× single line without doc grid", () => {
    expect(singleLinePx(12, {})).toBeCloseTo(12 * (96 / 72) * 1.15);
    expect(lineBoxPx({ fontSize: 12, lineRule: "auto", lineValue: 1.5 }, {})).toBeCloseTo(
      12 * (96 / 72) * 1.15 * 1.5,
    );
  });

  it("uses font em as single line in grid sections", () => {
    expect(singleLinePx(12, GRID_360)).toBeCloseTo(16);
    expect(
      lineBoxPx({ fontSize: 12, lineRule: "auto", lineValue: 1.5 }, GRID_360),
    ).toBe(24);
    expect(lineBoxPx({ fontSize: 12, lineRule: "auto", lineValue: 2 }, GRID_360)).toBe(32);
    expect(lineBoxPx({ fontSize: 12, lineRule: "auto", lineValue: 1 }, GRID_360)).toBe(24);
  });

  it("snaps 12pt @ 1.5× to 360 twips pitch only when close", () => {
    expect(
      lineBoxPx({ fontSize: 11.5, lineRule: "auto", lineValue: 1.5 }, GRID_360),
    ).not.toBe(24);
  });

  it("snaps MISSIONS-style 11.5pt @ 1.5× to two 272-twips rows", () => {
    const pitchPx = (272 / 15) * 2;
    expect(
      lineBoxPx({ fontSize: 11.5, lineRule: "auto", lineValue: 1.5 }, GRID_272),
    ).toBeCloseTo(pitchPx, 2);
  });

  it("handles exact and atLeast in px", () => {
    expect(lineBoxPx({ fontSize: 12, lineRule: "exact", lineValue: 25.2 }, {})).toBe(25.2);
    expect(
      lineBoxPx({ fontSize: 12, lineRule: "atLeast", lineValue: 20 }, {}),
    ).toBeGreaterThanOrEqual(20);
  });

  it("emits px line-height for grid auto spacing", () => {
    expect(
      lineHeightCss(
        { fontSize: 12, lineRule: "auto", lineValue: 1.5, docGridSectionPitchTwips: 360 },
      ),
    ).toBe("24px");
    expect(
      lineHeightCss(
        { fontSize: 12, lineRule: "auto", lineValue: 1, docGridSectionPitchTwips: 360 },
      ),
    ).toBe("24px");
  });

  it("resolveLineMetrics top-aligns grid and multiples", () => {
    expect(resolveLineMetrics({ lineRule: "auto", lineValue: 1.5 }, GRID_360).topAlign).toBe(
      true,
    );
    expect(resolveLineMetrics({ lineRule: "auto", lineValue: 1 }, {}).topAlign).toBe(false);
    expect(resolveLineMetrics({ lineRule: "auto", lineValue: 2 }, {}).topAlign).toBe(true);
  });

  it("lineBoxSlack matches Word leading rules", () => {
    const format = { lineRule: "auto" as const, lineValue: 1.5, fontSize: 12 };
    expect(lineBoxSlack(24, 16, format, GRID_360)).toBe(0);
    expect(lineBoxSlack(27.6, 16, format, {})).toBeCloseTo(27.6 - 27.6 / 1.5);
  });
});

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
