import { EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  MULTILEVEL_PRESETS,
  applyMultilevel,
  continueNumbering,
  restartNumbering,
  setNumbering,
} from "./lists";
import { basePlugins } from "./plugins";
import { schema } from "./schema";
import { applyStyle, currentStyleId, galleryStyles, updateStyleToMatch } from "./styles";
import {
  collectChanges,
  rejectAllChanges,
  trackChangesPlugin,
  trackTransaction,
} from "./track-changes";

function makeState(paragraphs: string[], tracking = false) {
  const doc = schema.node(
    "doc",
    null,
    paragraphs.map((t) => schema.node("paragraph", null, t ? [schema.text(t)] : [])),
  );
  const state = EditorState.create({
    doc,
    plugins: [
      ...basePlugins({ placeholder: "" }),
      trackChangesPlugin(() => ({ enabled: tracking, author: { id: "a", name: "Ann" } })),
    ],
  });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
}

function run(state: EditorState, command: Command) {
  let next = state;
  command(state, (tr) => {
    next = state.apply(trackTransaction(tr, state));
  });
  return next;
}

const select = (state: EditorState, pos: number) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));

describe("styles", () => {
  it("lists Word's quick styles in gallery order", () => {
    const ids = galleryStyles(makeState(["x"]).doc).map((s) => s.id);
    expect(ids.slice(0, 5)).toEqual(["Normal", "NoSpacing", "Heading1", "Heading2", "Heading3"]);
  });

  it("applies a heading style: node type, format and bold", () => {
    const s = run(makeState(["Title text", "body"]), applyStyle("Heading1"));
    const h = s.doc.firstChild!;
    expect(h.type.name).toBe("heading");
    expect(h.attrs.level).toBe(1);
    expect(h.attrs.format).toMatchObject({ styleId: "Heading1", fontSize: 16, keepNext: true });
    expect(h.firstChild!.marks.map((m) => m.type.name)).toEqual(["strong"]);
    expect(currentStyleId(s)).toBe("Heading1");

    const back = run(s, applyStyle("Normal"));
    expect(back.doc.firstChild!.type.name).toBe("paragraph");
    expect(back.doc.firstChild!.attrs.format).toBeNull();
    expect(back.doc.firstChild!.firstChild!.marks).toEqual([]);
  });

  it("tracks a style change and rejecting restores the paragraph", () => {
    const s = run(makeState(["Title text"], true), applyStyle("Heading2"));
    const changes = collectChanges(s.doc);
    const para = changes.find((c) => c.type === "paragraph");
    expect(para?.description).toMatch(/^Style: Heading 2/);
    const rejected = run(s, rejectAllChanges);
    expect(rejected.doc.firstChild!.type.name).toBe("paragraph");
    expect(rejected.doc.firstChild!.attrs.format).toBeNull();
    expect(rejected.doc.firstChild!.firstChild!.marks).toEqual([]);
  });

  it("updates a style to match the selection and restyles its paragraphs", () => {
    let s = run(makeState(["one", "two"]), applyStyle("Quote"));
    s = run(select(s, s.doc.firstChild!.nodeSize + 1), applyStyle("Quote"));
    s = run(
      s,
      (state, dispatch) => {
        const tr = state.tr.setNodeAttribute(0, "format", {
          ...state.doc.firstChild!.attrs.format,
          align: "right",
        });
        dispatch?.(tr.setSelection(TextSelection.create(tr.doc, 1)));
        return true;
      },
    );
    s = run(s, updateStyleToMatch);
    expect(s.doc.child(1).attrs.format.align).toBe("right");
    expect(s.doc.attrs.styles).toHaveLength(1);
  });

  it("Enter after a heading starts a Normal paragraph", () => {
    let s = run(makeState(["Head"]), applyStyle("Heading1"));
    s = select(s, 5);
    const enter = (state: EditorState) => {
      let next = state;
      for (const plugin of state.plugins) {
        const handled = plugin.props.handleKeyDown?.call(
          plugin,
          { state, dispatch: (tr: import("prosemirror-state").Transaction) => (next = state.apply(tr)) } as never,
          { key: "Enter", keyCode: 13, shiftKey: false, altKey: false, ctrlKey: false, metaKey: false } as never,
        );
        if (handled) break;
      }
      return next;
    };
    s = enter(s);
    expect(s.doc.childCount).toBe(2);
    expect(s.doc.child(1).type.name).toBe("paragraph");
    expect(s.doc.child(1).attrs.format).toBeNull();
  });
});

describe("lists", () => {
  it("numbers with a format and applies a multilevel preset", () => {
    let s = run(makeState(["a", "b"]), setNumbering("upperRoman"));
    const list = s.doc.firstChild!;
    expect(list.type.name).toBe("ordered_list");
    expect(list.attrs.numFormat).toBe("upperRoman");
    s = run(s, applyMultilevel(MULTILEVEL_PRESETS.find((p) => p.id === "legal")!));
    expect(s.doc.firstChild!.attrs.numFormat).toBe("legal");
  });

  it("restarts at the selected item and continues from the previous list", () => {
    const item = (t: string) => schema.node("list_item", null, [schema.node("paragraph", null, [schema.text(t)])]);
    const doc = schema.node("doc", null, [
      schema.node("ordered_list", null, [item("one"), item("two"), item("three")]),
    ]);
    let s = EditorState.create({ doc, plugins: [trackChangesPlugin(() => ({ enabled: false, author: { id: "a", name: "A" } }))] });
    // Inside "two" (second item).
    const two = 1 + doc.firstChild!.child(0).nodeSize + 2;
    s = run(select(s, two + 1), restartNumbering);
    expect(s.doc.childCount).toBe(2);
    expect(s.doc.child(1).attrs.order).toBe(1);
    expect(s.doc.child(1).childCount).toBe(2);
    s = run(s, continueNumbering);
    expect(s.doc.child(1).attrs.order).toBe(2);
  });
});
