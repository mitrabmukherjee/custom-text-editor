import { EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  addComment,
  collectComments,
  commentThreads,
  commentsInSelection,
  deleteComment,
  editComment,
  replyToComment,
  resolveComment,
} from "./comments";
import { basePlugins } from "./plugins";
import { schema } from "./schema";
import {
  acceptChanges,
  changesInSelection,
  collectChanges,
  setTrackChanges,
  setTrackingLocked,
  trackChangesKey,
  trackChangesPlugin,
  trackTransaction,
} from "./track-changes";

const ANN = { id: "ann", name: "Ann" };
const BOB = { id: "bob", name: "Bob" };

function makeState(paragraphs: string[], tracking = false) {
  const doc = schema.node(
    "doc",
    null,
    paragraphs.map((t) => schema.node("paragraph", null, [schema.text(t)])),
  );
  return EditorState.create({
    doc,
    plugins: [
      ...basePlugins({ placeholder: "" }),
      trackChangesPlugin(() => ({ enabled: tracking, author: ANN })),
    ],
  });
}

function run(state: EditorState, command: Command) {
  let next = state;
  const ok = command(state, (tr) => {
    next = state.apply(trackTransaction(tr, state));
  });
  return { state: next, ok };
}

const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

describe("comments", () => {
  it("anchors to the selection, or the word at the cursor", () => {
    let state = select(makeState(["The witness said yes."]), 5, 12);
    state = run(state, addComment("Who?", ANN)).state;
    state = run(select(state, 19), addComment("Really?", BOB)).state;

    const comments = collectComments(state.doc);
    expect(comments.map((c) => [c.quote, c.authorName, c.text])).toEqual([
      ["witness", "Ann", "Who?"],
      ["yes", "Bob", "Really?"],
    ]);
    expect(commentsInSelection(select(state, 7)).map((c) => c.text)).toEqual(["Who?"]);
    expect(run(select(state, 13), addComment("", ANN)).ok).toBe(false);
  });

  it("replies, edits, resolves and deletes threads", () => {
    let state = run(select(makeState(["alpha beta"]), 2), addComment("first", ANN)).state;
    const [{ id }] = collectComments(state.doc);
    state = run(state, replyToComment(id, "agreed", BOB)).state;
    state = run(state, editComment(id, "edited")).state;
    state = run(state, resolveComment(id, true)).state;
    expect(commentThreads(state.doc)[id]).toMatchObject({
      text: "edited",
      resolved: true,
      replies: [{ authorName: "Bob", text: "agreed" }],
    });

    state = run(state, deleteComment(id)).state;
    expect(state.doc.attrs.comments).toBeNull();
    let marked = false;
    state.doc.descendants((n) => {
      marked ||= n.marks.some((m) => m.type === schema.marks.comment);
    });
    expect(marked).toBe(false);
  });

  it("is not recorded as a revision while tracking", () => {
    let state = select(makeState(["alpha beta"], true), 1, 6);
    state = run(state, addComment("note", ANN)).state;
    expect(collectChanges(state.doc)).toHaveLength(0);
    expect(collectComments(state.doc)).toHaveLength(1);
  });
});

describe("review", () => {
  it("locks tracking on until unlocked", () => {
    let state = makeState(["text"]);
    state = run(state, setTrackingLocked(true)).state;
    expect(trackChangesKey.getState(state)).toMatchObject({ enabled: true, locked: true });
    expect(run(state, setTrackChanges(false)).ok).toBe(false);
    state = run(state, setTrackingLocked(false)).state;
    state = run(state, setTrackChanges(false)).state;
    expect(trackChangesKey.getState(state)?.enabled).toBe(false);
  });

  it("accepts only the changes in the selection", () => {
    let state = makeState(["one two"], true);
    state = run(select(state, 4), (s, d) => {
      d?.(s.tr.insertText(" A"));
      return true;
    }).state;
    state = run(select(state, state.doc.content.size - 1), (s, d) => {
      d?.(s.tr.insertText(" B"));
      return true;
    }).state;
    expect(collectChanges(state.doc)).toHaveLength(2);

    const inSelection = changesInSelection(select(state, 5));
    expect(inSelection.map((c) => c.text)).toEqual([" A"]);
    state = run(state, acceptChanges(inSelection.map((c) => c.id))).state;
    expect(collectChanges(state.doc).map((c) => c.text)).toEqual([" B"]);
    expect(state.doc.textContent).toBe("one A two B");
  });
});
