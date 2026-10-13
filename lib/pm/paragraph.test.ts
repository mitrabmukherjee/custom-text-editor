import { EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import type { ParagraphFormat } from "./format";
import {
  INDENT_STEP,
  activeList,
  hangingIndent,
  indent,
  paragraphFormatAt,
  setParagraphFormat,
  toggleList,
} from "./paragraph";
import { schema } from "./schema";

function makeState(paragraphs: string[], format: ParagraphFormat | null = null) {
  const doc = schema.node(
    "doc",
    null,
    paragraphs.map((t) => schema.node("paragraph", { format }, t ? [schema.text(t)] : [])),
  );
  return EditorState.create({ doc });
}

function selectAll(state: EditorState) {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)),
  );
}

function run(state: EditorState, command: Command) {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

const formats = (state: EditorState) => {
  const out: (ParagraphFormat | null)[] = [];
  state.doc.descendants((node) => {
    if (node.isTextblock) out.push(node.attrs.format as ParagraphFormat | null);
    return true;
  });
  return out;
};

describe("paragraph formatting", () => {
  it("merges a patch into every selected paragraph", () => {
    const s = run(
      selectAll(makeState(["one", "two"], { fontSize: 12 })),
      setParagraphFormat({ align: "center", lineRule: "auto", lineValue: 1.5 }),
    );
    expect(formats(s)).toEqual([
      { fontSize: 12, align: "center", lineRule: "auto", lineValue: 1.5 },
      { fontSize: 12, align: "center", lineRule: "auto", lineValue: 1.5 },
    ]);
  });

  it("drops cleared properties and stores null when nothing is left", () => {
    let s = run(makeState(["one"]), setParagraphFormat({ align: "right" }));
    expect(paragraphFormatAt(s)).toEqual({ align: "right" });
    s = run(s, setParagraphFormat({ align: undefined }));
    expect(formats(s)).toEqual([null]);
  });

  it("steps the left indent to the next and previous 0.5 inch stop", () => {
    let s = run(makeState(["one"], { indentLeft: 20 }), indent(1));
    expect(paragraphFormatAt(s).indentLeft).toBe(INDENT_STEP);
    s = run(s, indent(1));
    expect(paragraphFormatAt(s).indentLeft).toBe(2 * INDENT_STEP);
    s = run(run(s, indent(-1)), indent(-1));
    expect(paragraphFormatAt(s).indentLeft).toBeUndefined();
  });

  it("adds and removes a hanging indent like Ctrl+T", () => {
    let s = run(makeState(["one"]), hangingIndent(1));
    expect(paragraphFormatAt(s)).toMatchObject({ indentLeft: INDENT_STEP, indentFirstLine: -INDENT_STEP });
    s = run(s, hangingIndent(-1));
    expect(formats(s)).toEqual([null]);
  });

  it("toggles bullets, switches to numbering, and removes the list", () => {
    let s = run(makeState(["one"]), toggleList(schema.nodes.bullet_list));
    expect(activeList(s)).toBe(schema.nodes.bullet_list);
    s = run(s, toggleList(schema.nodes.ordered_list));
    expect(activeList(s)).toBe(schema.nodes.ordered_list);
    s = run(s, toggleList(schema.nodes.ordered_list));
    expect(activeList(s)).toBeNull();
    expect(s.doc.child(0).type).toBe(schema.nodes.paragraph);
  });
});
