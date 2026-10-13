import { describe, expect, it } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import { schema } from "./schema";
import { insertSectionBreak, sectionIndexAt, sectionsOf } from "./sections";
import { setHeaderFooterText } from "./header-footer";

function doc(...content: ReturnType<typeof schema.nodes.paragraph.create>[]) {
  return schema.nodes.doc.create({ layout: { page: { width: 816, height: 1056, marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96, gap: 24 } } }, content);
}

describe("sections", () => {
  it("insertSectionBreak adds a section_break node", () => {
    const d = doc(schema.nodes.paragraph.create(), schema.nodes.paragraph.create());
    const state = EditorState.create({
      doc: d,
      selection: TextSelection.create(d, 2),
    });
    let next = state;
    expect(insertSectionBreak("nextPage")(state, (tr) => {
      next = state.apply(tr);
    })).toBe(true);
    expect(next.doc.childCount).toBe(3);
    expect(next.doc.child(1).type).toBe(schema.nodes.section_break);
    expect(next.doc.child(1).attrs.start).toBe("nextPage");
  });

  it("sectionIndexAt counts breaks before position", () => {
    const d = doc(
      schema.nodes.paragraph.create(),
      schema.nodes.section_break.create({ start: "continuous", section: {} }),
      schema.nodes.paragraph.create(),
    );
    const state = EditorState.create({ doc: d });
    const breakPos = 2;
    expect(sectionIndexAt(state, 1)).toBe(0);
    expect(sectionIndexAt(state, breakPos + 1)).toBe(1);
  });

  it("setHeaderFooterText updates layout footers", () => {
    const state = EditorState.create({ doc: doc(schema.nodes.paragraph.create()) });
    const tr = setHeaderFooterText(state, { section: 0, part: "footer", variant: "default" }, "Page {PAGE}");
    expect(tr).toBeTruthy();
    const next = state.apply(tr!);
    const sections = sectionsOf(next);
    const runs = sections[0].footers?.default?.[0].runs ?? [];
    expect(runs.some((r) => r.kind === "field" && r.field === "PAGE")).toBe(true);
  });
});
