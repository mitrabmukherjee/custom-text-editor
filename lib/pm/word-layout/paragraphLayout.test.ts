import { describe, expect, it } from "vitest";
import { schema } from "@/lib/pm/schema";
import { layoutParagraph } from "./paragraphLayout";
import { createProportionalMeasurer } from "./measure";

describe("layoutParagraph", () => {
  it("wraps at word boundaries, not mid-word", () => {
    const text =
      "DEMITRY PAPASOTIRIOU:  Good afternoon, Papasotiriou, Demitry.";
    const block = schema.nodes.paragraph.create(
      {
        format: {
          fontSize: 12,
          fontFamily: "Courier New",
          lineRule: "auto",
          lineValue: 1.5,
          indentLeft: 132,
          indentRight: 19,
          docGridSectionPitchTwips: 360,
        },
      },
      [schema.text(text)],
    );
    const measurer = createProportionalMeasurer();
    const contentWidth = 604;
    const { lines } = layoutParagraph(
      block,
      0,
      contentWidth,
      100,
      0,
      0,
      { defaultTabStop: 48, pitchTwips: 360, measurer },
      {},
    );
    expect(lines.length).toBeGreaterThan(1);
    const wordBoundary = (offset: number) => {
      if (offset <= 0 || offset >= text.length) return true;
      return !(/\w/.test(text[offset - 1]) && /\w/.test(text[offset]));
    };
    for (const line of lines) {
      const offset = line.anchor?.textOffset ?? 0;
      expect(wordBoundary(offset)).toBe(true);
    }
  });

  it("keeps trailing spaces on the same line as the paragraph text", () => {
    const text = "THE COURT:  Thank you.  ";
    const block = schema.nodes.paragraph.create(
      {
        format: {
          fontSize: 12,
          fontFamily: "Courier New",
          lineRule: "auto",
          lineValue: 1.5,
          indentLeft: 132,
          indentRight: 19,
          docGridSectionPitchTwips: 360,
        },
      },
      [schema.text(text)],
    );
    const { lines } = layoutParagraph(
      block,
      0,
      604,
      100,
      0,
      0,
      { defaultTabStop: 48, pitchTwips: 360, measurer: createProportionalMeasurer() },
      {},
    );
    expect(lines).toHaveLength(1);
    expect(lines[0].anchor?.textOffset).toBe(0);
  });
});
