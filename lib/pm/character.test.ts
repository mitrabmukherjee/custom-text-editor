import { Slice } from "prosemirror-model";
import { EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  applyCharacterFormat,
  characterFormatAt,
  clearFormatting,
  growFont,
  setFontSize,
  setTextColor,
  toggleSuperscript,
  toggleSubscript,
  toggleUnderline,
} from "./character";
import { cleanWordHtml, transformSlice } from "./plugins/paste";
import { schema } from "./schema";
import { collectChanges, trackChangesPlugin, trackTransaction } from "./track-changes";

function makeState(text: string, tracking = false) {
  const doc = schema.node("doc", null, [schema.node("paragraph", null, [schema.text(text)])]);
  const state = EditorState.create({
    doc,
    plugins: [trackChangesPlugin(() => ({ enabled: tracking, author: { id: "a", name: "Ann" } }))],
  });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 1 + text.length)));
}

function run(state: EditorState, command: Command) {
  let next = state;
  command(state, (tr) => {
    next = state.apply(trackTransaction(tr, state));
  });
  return next;
}

const marksOf = (state: EditorState) =>
  state.doc.firstChild!.firstChild!.marks.map((m) => [m.type.name, m.attrs]);

describe("character commands", () => {
  it("toggles underline with a style, and super/subscript exclusively", () => {
    let s = run(makeState("word"), toggleUnderline("double"));
    expect(marksOf(s)).toEqual([["underline", { style: "double" }]]);
    s = run(s, toggleSuperscript);
    s = run(s, toggleSubscript);
    expect(characterFormatAt(s)).toMatchObject({ underline: "double", vertAlign: "subscript" });
    s = run(s, toggleSubscript);
    expect(characterFormatAt(s).vertAlign).toBeNull();
  });

  it("grows through Word's size list and by one point", () => {
    let s = run(makeState("word"), growFont(1));
    expect(characterFormatAt(s).size).toBe(14);
    s = run(s, growFont(1, true));
    expect(characterFormatAt(s).size).toBe(15);
    s = run(s, growFont(-1));
    expect(characterFormatAt(s).size).toBe(14);
    s = run(s, setFontSize(12));
    expect(marksOf(s)).toEqual([]);
  });

  it("clears every character format", () => {
    let s = run(makeState("word"), setTextColor("#ff0000"));
    s = run(s, toggleUnderline());
    s = run(s, clearFormatting);
    expect(marksOf(s)).toEqual([]);
  });

  it("records toolbar formatting as one tracked revision", () => {
    const s = run(
      makeState("word", true),
      applyCharacterFormat({ underline: "single", color: "#ff0000", size: 14 }),
    );
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0].type).toBe("format");
    expect(changes[0].description).toMatch(/14 pt.*Underline.*Red|Underline.*Red.*14/);
  });
});

describe("paste", () => {
  it("cleans Word clipboard HTML", () => {
    const html =
      '<!--[if gte mso 9]><xml></xml><![endif]--><style>p{}</style><p class=MsoNormal style="mso-line-height-alt:12pt;color:red">Hi<o:p></o:p></p>';
    expect(cleanWordHtml(html)).toBe('<p class=MsoNormal style="color:red">Hi</p>');
  });

  it("merge formatting keeps emphasis but drops fonts and colours", () => {
    const para = schema.node("paragraph", { format: { align: "center" } }, [
      schema.text("x", [
        schema.marks.strong.create(),
        schema.marks.textColor.create({ color: "#eeeeee" }),
        schema.marks.textStyle.create({ fontFamily: "Inter", fontSize: 9 }),
      ]),
      schema.text("gone", [schema.marks.deletion.create({ id: "d", authorId: "a", authorName: "A", date: "" })]),
    ]);
    const slice = transformSlice(
      new Slice(schema.node("doc", null, [para]).content, 0, 0),
      (m) => !["textStyle", "textColor", "highlight"].includes(m.type.name),
      "clear",
    );
    const out = slice.content.firstChild!;
    expect(out.attrs.format).toBeNull();
    expect(out.textContent).toBe("x");
    expect(out.firstChild!.marks.map((m) => m.type.name)).toEqual(["strong"]);
  });
});
