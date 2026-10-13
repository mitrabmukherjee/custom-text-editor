import { EditorState, TextSelection, type Command, type Transaction } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { docJsonToPlainText } from "@/lib/serializers/plainText";
import type { DocJSON } from "../json";
import { schema } from "../schema";
import {
  acceptAllChanges,
  acceptChange,
  collectChanges,
  rejectAllChanges,
  rejectChange,
} from "./changes";
import { SKIP_TRACKING, trackChangesPlugin, type Author } from "./plugin";
import { trackTransaction } from "./track";

const alice: Author = { id: "alice", name: "Alice" };
const bob: Author = { id: "bob", name: "Bob" };

function makeState(paragraphs: string[], author: Author = alice, enabled = true) {
  const doc = schema.node(
    "doc",
    null,
    paragraphs.map((t) => schema.node("paragraph", null, t ? [schema.text(t)] : [])),
  );
  return EditorState.create({
    doc,
    plugins: [trackChangesPlugin(() => ({ enabled, author }))],
  });
}

function withAuthor(state: EditorState, author: Author) {
  return EditorState.create({
    doc: state.doc,
    plugins: [trackChangesPlugin(() => ({ enabled: true, author }))],
  });
}

/** Apply a user transaction through the tracking pipeline. */
function edit(state: EditorState, build: (tr: Transaction) => Transaction) {
  return state.apply(trackTransaction(build(state.tr), state));
}

function run(state: EditorState, command: Command) {
  let next = state;
  command(state, (tr) => {
    next = state.apply(trackTransaction(tr, state));
  });
  return next;
}

function text(state: EditorState, mode: "accepted" | "original" = "accepted") {
  return docJsonToPlainText(state.doc.toJSON() as DocJSON, { mode });
}

/** Paragraph 1 positions: text starts at 1. */
const at = (offset: number) => 1 + offset;

describe("trackTransaction", () => {
  it("marks typed text as an insertion", () => {
    const s = edit(makeState(["Hello world"]), (tr) => tr.insertText(" big", at(5)));
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({
      type: "insertion",
      text: " big",
      authorId: "alice",
      authorName: "Alice",
    });
    expect(text(s)).toBe("Hello big world");
    expect(text(s, "original")).toBe("Hello world");
  });

  it("keeps consecutive typing in one change", () => {
    let s = makeState(["ab"]);
    s = edit(s, (tr) => tr.insertText("1", at(1)));
    s = edit(s, (tr) => tr.insertText("2", at(2)));
    s = edit(s, (tr) => tr.insertText("3", at(3)));
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0].text).toBe("123");
  });

  it("strikes deleted text instead of removing it", () => {
    const s = edit(makeState(["Hello cruel world"]), (tr) => tr.delete(at(5), at(11)));
    expect(s.doc.textContent).toBe("Hello cruel world");
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({ type: "deletion", text: " cruel" });
    expect(text(s)).toBe("Hello world");
    expect(text(s, "original")).toBe("Hello cruel world");
  });

  it("removes the author's own insertion for real", () => {
    let s = edit(makeState(["Hello"]), (tr) => tr.insertText(" there", at(5)));
    s = edit(s, (tr) => tr.delete(at(5), at(11)));
    expect(s.doc.textContent).toBe("Hello");
    expect(collectChanges(s.doc)).toHaveLength(0);
  });

  it("strikes another author's insertion", () => {
    let s = edit(makeState(["Hello"]), (tr) => tr.insertText(" there", at(5)));
    s = withAuthor(s, bob);
    s = edit(s, (tr) => tr.delete(at(5), at(11)));
    expect(s.doc.textContent).toBe("Hello there");
    const types = collectChanges(s.doc).map((c) => `${c.type}:${c.authorId}`);
    expect(types).toEqual(expect.arrayContaining(["insertion:alice", "deletion:bob"]));
    expect(text(s)).toBe("Hello");
  });

  it("tracks a replacement as deletion plus insertion", () => {
    const s = edit(makeState(["The cat sat"]), (tr) =>
      tr.replaceWith(at(4), at(7), schema.text("dog")),
    );
    expect(s.doc.textContent).toBe("The catdog sat");
    const changes = collectChanges(s.doc);
    expect(changes.map((c) => [c.type, c.text])).toEqual([
      ["deletion", "cat"],
      ["insertion", "dog"],
    ]);
    expect(text(s)).toBe("The dog sat");
  });

  it("places the cursor before struck text on Backspace", () => {
    const base = makeState(["abc"]);
    const state = base.apply(
      base.tr.setSelection(TextSelection.create(base.doc, at(3))),
    );
    const tr = state.tr.delete(at(2), at(3));
    const s = state.apply(trackTransaction(tr, state));
    expect(s.selection.head).toBe(at(2));
  });

  it("passes edits through when tracking is off", () => {
    const s = edit(makeState(["Hello world"], alice, false), (tr) => tr.delete(at(5), at(11)));
    expect(s.doc.textContent).toBe("Hello");
    expect(collectChanges(s.doc)).toHaveLength(0);
  });

  it("skips transactions tagged SKIP_TRACKING", () => {
    const s = edit(makeState(["Hello world"]), (tr) =>
      tr.delete(at(5), at(11)).setMeta(SKIP_TRACKING, true),
    );
    expect(s.doc.textContent).toBe("Hello");
  });

  it("does not track paragraph splits as text", () => {
    const s = edit(makeState(["Hello world"]), (tr) => tr.split(at(5)));
    expect(s.doc.childCount).toBe(2);
    expect(collectChanges(s.doc)).toHaveLength(0);
  });
});

describe("formatting revisions", () => {
  const italic = (s: EditorState, from: number, to: number) =>
    edit(s, (tr) => tr.addMark(from, to, schema.marks.em.create()));

  it("records character formatting as a formatting revision", () => {
    const s = italic(makeState(["Hello world"]), at(6), at(11));
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({
      type: "format",
      text: "world",
      description: "Font: Italic",
      authorName: "Alice",
    });
  });

  it("drops the revision when the formatting is toggled back", () => {
    let s = italic(makeState(["Hello world"]), at(6), at(11));
    s = edit(s, (tr) => tr.removeMark(at(6), at(11), schema.marks.em));
    expect(collectChanges(s.doc)).toHaveLength(0);
  });

  it("keeps the original formatting across successive changes", () => {
    let s = italic(makeState(["Hello world"]), at(6), at(11));
    s = edit(s, (tr) => tr.addMark(at(6), at(11), schema.marks.strong.create()));
    const changes = collectChanges(s.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0].description).toBe("Font: Bold, Italic");
  });

  it("does not track formatting of the author's own insertion", () => {
    let s = edit(makeState(["Hello"]), (tr) => tr.insertText(" there", at(5)));
    s = italic(s, at(6), at(11));
    expect(collectChanges(s.doc).map((c) => c.type)).toEqual(["insertion"]);
  });

  it("rejecting restores the old formatting; accepting keeps the new", () => {
    const s = italic(makeState(["Hello world"]), at(6), at(11));
    const rejected = run(s, rejectAllChanges);
    expect(collectChanges(rejected.doc)).toHaveLength(0);
    expect(rejected.doc.rangeHasMark(at(6), at(11), schema.marks.em)).toBe(false);

    const accepted = run(s, acceptAllChanges);
    expect(collectChanges(accepted.doc)).toHaveLength(0);
    expect(accepted.doc.rangeHasMark(at(6), at(11), schema.marks.em)).toBe(true);
  });

  it("tracks underline and font colour, and rejecting restores them", () => {
    let s = edit(makeState(["Hello world"]), (tr) =>
      tr.addMark(at(6), at(11), schema.marks.underline.create()),
    );
    s = edit(s, (tr) =>
      tr.addMark(at(6), at(11), schema.marks.textColor.create({ color: "#ff0000" })),
    );
    const [change] = collectChanges(s.doc);
    expect(change.description).toBe("Font: Underline, Font color: Red");
    const rejected = run(s, rejectAllChanges);
    expect(rejected.doc.rangeHasMark(at(6), at(11), schema.marks.underline)).toBe(false);
    expect(rejected.doc.rangeHasMark(at(6), at(11), schema.marks.textColor)).toBe(false);
  });
});

describe("paragraph formatting revisions", () => {
  const center = (s: EditorState) =>
    edit(s, (tr) => tr.setNodeAttribute(0, "format", { align: "center" }));

  it("records a paragraph format change as a revision", () => {
    const s = center(makeState(["Hello", "World"]));
    const changes = collectChanges(s.doc);
    expect(changes).toMatchObject([
      { type: "paragraph", description: "Centered", authorName: "Alice", text: "Hello" },
    ]);
  });

  it("drops the revision when the format is set back", () => {
    let s = center(makeState(["Hello"]));
    s = edit(s, (tr) => tr.setNodeAttribute(0, "format", null));
    expect(collectChanges(s.doc)).toHaveLength(0);
    expect(s.doc.child(0).attrs.paraChange).toBeNull();
  });

  it("rejecting restores the old format; accepting keeps the new", () => {
    const s = center(makeState(["Hello"]));
    const rejected = run(s, rejectAllChanges);
    expect(rejected.doc.child(0).attrs.format).toBeNull();
    expect(collectChanges(rejected.doc)).toHaveLength(0);
    const accepted = run(s, acceptAllChanges);
    expect(accepted.doc.child(0).attrs.format).toEqual({ align: "center" });
    expect(accepted.doc.child(0).attrs.paraChange).toBeNull();
  });
});

describe("accept / reject", () => {
  function reviewed() {
    let s = makeState(["The cat sat"]);
    s = edit(s, (tr) => tr.delete(at(4), at(8)));
    s = edit(s, (tr) => tr.insertText(" down", s.doc.content.size - 1));
    return s;
  }

  it("accepts a single change", () => {
    const s = reviewed();
    const ins = collectChanges(s.doc).find((c) => c.type === "insertion")!;
    const next = run(s, acceptChange(ins.id));
    expect(collectChanges(next.doc).map((c) => c.type)).toEqual(["deletion"]);
    expect(next.doc.textContent).toBe("The cat sat down");
  });

  it("rejects a single change", () => {
    const s = reviewed();
    const del = collectChanges(s.doc).find((c) => c.type === "deletion")!;
    const next = run(s, rejectChange(del.id));
    expect(collectChanges(next.doc).map((c) => c.type)).toEqual(["insertion"]);
    expect(text(next)).toBe("The cat sat down");
  });

  it("accepts all changes", () => {
    const next = run(reviewed(), acceptAllChanges);
    expect(collectChanges(next.doc)).toHaveLength(0);
    expect(next.doc.textContent).toBe("The sat down");
  });

  it("rejects all changes", () => {
    const next = run(reviewed(), rejectAllChanges);
    expect(collectChanges(next.doc)).toHaveLength(0);
    expect(next.doc.textContent).toBe("The cat sat");
  });
});
