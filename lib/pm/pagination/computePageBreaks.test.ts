import { describe, expect, it } from "vitest";
import { computePageBreaks, type LayoutLine } from "./computePageBreaks";
import { LETTER_PAGE_SETUP, contentHeight, pageStride, stackHeight } from "./pageSetup";

const setup = LETTER_PAGE_SETUP;
const LINE = 32;

/** Lines in page coordinates, starting at the top margin. */
function lines(count: number, height = LINE, start = setup.marginTop): LayoutLine[] {
  return Array.from({ length: count }, (_, i) => ({
    top: start + i * height,
    bottom: start + (i + 1) * height,
  }));
}

/** Give consecutive lines block ids, e.g. blocks([3, 2]) → 3 lines of block 0, 2 of block 1. */
function inBlocks(input: LayoutLine[], sizes: number[], extra: Partial<LayoutLine> = {}) {
  let i = 0;
  sizes.forEach((size, block) => {
    for (let k = 0; k < size; k++, i++) input[i] = { ...input[i], block, ...extra };
  });
  return input;
}

describe("page setup", () => {
  it("derives Letter geometry", () => {
    expect(contentHeight(setup)).toBe(864);
    expect(pageStride(setup)).toBe(1080);
    expect(stackHeight(setup, 2)).toBe(2136);
  });
});

describe("computePageBreaks", () => {
  it("keeps short documents on one page", () => {
    const short = computePageBreaks(lines(10), setup);
    expect(short.breaks).toEqual([]);
    expect(short.pageCount).toBe(1);
    expect(computePageBreaks([], setup).pageCount).toBe(1);
  });

  it("fills a page exactly without breaking", () => {
    // 27 lines * 32px = 864px = the full content height.
    expect(computePageBreaks(lines(27), setup).pageCount).toBe(1);
  });

  it("moves the first overflowing line to the next page", () => {
    const { breaks, pageCount } = computePageBreaks(lines(28), setup);
    expect(pageCount).toBe(2);
    expect(breaks).toEqual([{ lineIndex: 27, height: 1080 - 864, forced: false }]);
  });

  it("breaks repeatedly for long documents", () => {
    const { breaks, pageCount } = computePageBreaks(lines(60), setup);
    expect(pageCount).toBe(3);
    expect(breaks.map((b) => b.lineIndex)).toEqual([27, 54]);
    for (const b of breaks) expect(b.height).toBeCloseTo(216);
  });

  it("puts each shifted line at the top of a page's content area", () => {
    // 28 lines of 30px fit in 864px; the 29th moves.
    const input = lines(30, 30);
    const { breaks, placements } = computePageBreaks(input, setup);
    const [b] = breaks;
    expect(b.lineIndex).toBe(28);
    expect(input[b.lineIndex].top + b.height).toBe(pageStride(setup) + setup.marginTop);
    expect(placements[28]).toEqual({ page: 1, top: 96, bottom: 126, section: 0 });
    expect(placements[27]).toEqual({ page: 0, top: 96 + 27 * 30, bottom: 96 + 28 * 30, section: 0 });
  });

  it("honours a hard page break", () => {
    const input = lines(5);
    input[1] = { ...input[1], breakAfter: true };
    const { breaks, pageCount } = computePageBreaks(input, setup);
    expect(pageCount).toBe(2);
    expect(breaks).toEqual([{ lineIndex: 2, height: 1080 - 64, forced: true }]);
  });

  it("does not add an empty page for a trailing hard break", () => {
    const input = lines(3);
    input[2] = { ...input[2], breakAfter: true };
    expect(computePageBreaks(input, setup).pageCount).toBe(1);
  });

  it("lets a line taller than a page overflow instead of looping", () => {
    const giant: LayoutLine[] = [
      { top: 96, bottom: 2096 },
      { top: 2096, bottom: 2128 },
    ];
    const { breaks, pageCount } = computePageBreaks(giant, setup);
    expect(breaks).toHaveLength(1);
    expect(pageCount).toBe(2);
  });

  it("lets the leading below the last line's text hang past the page bottom", () => {
    // 29 lines * 30px = 870px > 864px, but the 29th line's text ends at 860px.
    const loose = lines(30, 30).map((l) => ({ ...l, slack: 10 }));
    expect(computePageBreaks(loose, setup).breaks[0].lineIndex).toBe(29);
    expect(computePageBreaks(lines(30, 30), setup).breaks[0].lineIndex).toBe(28);
  });

  it("respects custom margins", () => {
    const tight = { ...setup, marginTop: 48, marginBottom: 48 };
    // 960px usable → 30 lines of 32px fit.
    expect(computePageBreaks(lines(30, LINE, 48), tight).pageCount).toBe(1);
    expect(computePageBreaks(lines(31, LINE, 48), tight).breaks[0].lineIndex).toBe(30);
  });

  it("uses per-page frames (taller footer from page 2)", () => {
    const frame = (page: { number: number }) => ({ top: 96, bottom: page.number === 1 ? 96 : 160 });
    const { breaks } = computePageBreaks(lines(60), setup, frame);
    // Page 1: 27 lines. Page 2: (1056 - 96 - 160) / 32 = 25 lines.
    expect(breaks.map((b) => b.lineIndex)).toEqual([27, 52]);
  });

  it("applies widow control: never leaves one line alone on the next page", () => {
    // Block 1 = lines 24..27; only line 27 overflows, so lines 26-27 move together.
    const input = inBlocks(lines(28), [24, 4], { widowControl: true });
    expect(computePageBreaks(input, setup).breaks[0].lineIndex).toBe(26);
  });

  it("moves a whole 3-line paragraph when a widow fix would leave an orphan", () => {
    const input = inBlocks(lines(28), [25, 3], { widowControl: true });
    expect(computePageBreaks(input, setup).breaks[0].lineIndex).toBe(25);
  });

  it("applies orphan control: never leaves one line alone at the page bottom", () => {
    // Block 1 starts on the last line of the page.
    const input = inBlocks(lines(30), [26, 4], { widowControl: true });
    expect(computePageBreaks(input, setup).breaks[0].lineIndex).toBe(26);
  });

  it("splits freely without widow control", () => {
    const input = inBlocks(lines(28), [25, 3]);
    expect(computePageBreaks(input, setup).breaks[0].lineIndex).toBe(27);
  });

  it("keeps a keep-with-next paragraph with the next one", () => {
    // Heading (block 1, line 26) is the last line on the page; block 2 starts the next page.
    const input = inBlocks(lines(30), [26, 1, 3]);
    input[26] = { ...input[26], keepNext: true };
    expect(computePageBreaks(input, setup).breaks[0].lineIndex).toBe(26);
  });

  it("starts a page-break-before paragraph on a new page", () => {
    const input = inBlocks(lines(6), [3, 3]);
    input[3] = { ...input[3], breakBefore: true };
    const { breaks, pageCount } = computePageBreaks(input, setup);
    expect(pageCount).toBe(2);
    expect(breaks).toMatchObject([{ lineIndex: 3, forced: true }]);
  });

  it("tracks sections per page and leaves a blank page for an odd-page break", () => {
    const input = lines(4).map((l, i) => ({ ...l, section: i < 2 ? 0 : 1 }));
    input[1] = { ...input[1], breakAfter: "odd" };
    const layout = computePageBreaks(input, setup, undefined, {
      numberStart: (s) => (s === 1 ? 1 : undefined),
    });
    // Page 2 would be even, so the section starts on page 3.
    expect(layout.pageCount).toBe(3);
    expect(layout.placements[2].page).toBe(2);
    expect(layout.pages).toEqual([
      { section: 0, first: true, number: 1 },
      { section: 0, first: false, number: 2 },
      { section: 1, first: true, number: 1 },
    ]);
  });
});
