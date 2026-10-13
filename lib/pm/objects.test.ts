import { EditorState, NodeSelection, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { fitImage, insertImage, linkAt, removeLink, setImageSize, setLink } from "./objects";
import { basePlugins } from "./plugins";
import { safeHref, schema } from "./schema";
import { insertTable, setTableBorders, tabInTable, tableBorders } from "./tables";

function makeState(text: string, from?: number, to = from) {
  const doc = schema.node("doc", null, [schema.node("paragraph", null, text ? [schema.text(text)] : [])]);
  const state = EditorState.create({ doc, plugins: basePlugins({ placeholder: "" }) });
  return from === undefined ? state : state.apply(state.tr.setSelection(TextSelection.create(doc, from, to)));
}

function run(state: EditorState, command: Command) {
  let next = state;
  const ok = command(state, (tr) => {
    next = state.apply(tr);
  });
  return { state: next, ok };
}

function linkedText(state: EditorState) {
  const out: string[] = [];
  state.doc.descendants((n) => {
    const link = n.marks.find((m) => m.type === schema.marks.link);
    if (n.isText && link) out.push(`${n.text}=${link.attrs.href}`);
  });
  return out;
}

describe("tables", () => {
  it("inserts a Word table after the paragraph with equal columns and a trailing paragraph", () => {
    const { state } = run(makeState("Intro", 3), insertTable(2, 3, 600));
    const table = state.doc.child(1);
    expect(table.type.name).toBe("table");
    expect(table.childCount).toBe(2);
    expect(table.child(0).childCount).toBe(3);
    expect(table.child(0).child(0).attrs.colwidth).toEqual([200]);
    expect(state.doc.lastChild!.type.name).toBe("paragraph");
    expect(state.selection.$from.parent.type.name).toBe("paragraph");
    expect(state.selection.$from.node(-1).type.name).toBe("table_cell");
  });

  it("Tab moves through cells and adds a row after the last one", () => {
    let { state } = run(makeState("", 1), insertTable(1, 2, 400));
    ({ state } = run(state, tabInTable));
    expect(state.selection.$from.index(-2)).toBe(1);
    ({ state } = run(state, tabInTable));
    expect(state.doc.child(1).childCount).toBe(2);
    expect(state.selection.$from.index(-3)).toBe(1);
  });

  it("toggles gridlines", () => {
    let { state } = run(makeState("", 1), insertTable(1, 1, 400));
    expect(tableBorders(state)).toBe(true);
    ({ state } = run(state, setTableBorders(false)));
    expect(tableBorders(state)).toBe(false);
    expect(tableBorders(makeState("x", 1))).toBeNull();
  });
});

describe("hyperlinks", () => {
  it("links the selection and edits the whole link from a cursor inside it", () => {
    let { state } = run(makeState("Visit the site now", 7, 15), setLink("https://a.example"));
    expect(linkedText(state)).toEqual(["the site=https://a.example"]);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 10)));
    expect(linkAt(state)).toMatchObject({ href: "https://a.example", text: "the site", from: 7, to: 15 });
    ({ state } = run(state, setLink("https://b.example", "our site")));
    expect(state.doc.textContent).toBe("Visit our site now");
    expect(linkedText(state)).toEqual(["our site=https://b.example"]);
    ({ state } = run(state.apply(state.tr.setSelection(TextSelection.create(state.doc, 9))), removeLink));
    expect(linkedText(state)).toEqual([]);
  });

  it("inserts the address as text when nothing is selected", () => {
    const { state } = run(makeState("Go ", 4), setLink("mailto:x@y.z"));
    expect(state.doc.textContent).toBe("Go mailto:x@y.z");
    expect(state.storedMarks ?? []).toEqual([]);
  });

  it("only renders safe link targets", () => {
    expect(safeHref("javascript:alert(1)")).toBeNull();
    expect(safeHref("https://ok.example")).toBe("https://ok.example");
    expect(safeHref("#bookmark")).toBe("#bookmark");
  });
});

describe("pictures", () => {
  it("inserts and resizes an inline picture", () => {
    let { state } = run(makeState("ab", 2), insertImage({ src: "data:image/png;base64,AA==", width: 50, height: 25 }));
    const pos = 2;
    expect(state.doc.nodeAt(pos)?.type.name).toBe("image");
    ({ state } = run(state, setImageSize(pos, 100.4, 50.2)));
    expect(state.doc.nodeAt(pos)?.attrs).toMatchObject({ width: 100, height: 50 });
    expect(state.selection).toBeInstanceOf(NodeSelection);
  });

  it("fits wide pictures to the text column", () => {
    expect(fitImage(1200, 600, 600)).toEqual({ width: 600, height: 300 });
    expect(fitImage(300, 200, 600)).toEqual({ width: 300, height: 200 });
  });
});
