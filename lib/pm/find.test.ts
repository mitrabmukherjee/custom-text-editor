import { EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { findMatches, findNext, findState, replaceAll, replaceMatch, setFindQuery } from "./find";
import { documentOutline, outlineIndexAt } from "./outline";
import { basePlugins } from "./plugins";
import { schema } from "./schema";
import { textStats } from "./stats";
import { collectChanges, trackChangesPlugin, trackTransaction } from "./track-changes";

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

const texts = (state: EditorState) => {
  const out: string[] = [];
  state.doc.forEach((p) => out.push(p.textContent));
  return out;
};

describe("find and replace", () => {
  it("matches case-insensitively by default, with match case and whole word options", () => {
    const doc = makeState(["The cat scattered.", "CAT cat"]).doc;
    expect(findMatches(doc, { text: "cat" })).toHaveLength(4);
    expect(findMatches(doc, { text: "cat", matchCase: true })).toHaveLength(3);
    expect(findMatches(doc, { text: "cat", wholeWord: true })).toHaveLength(3);
    expect(findMatches(doc, { text: "cat", wholeWord: true, matchCase: true })).toHaveLength(2);
    const [first] = findMatches(doc, { text: "cat" });
    expect(doc.textBetween(first.from, first.to)).toBe("cat");
  });

  it("highlights matches and steps through them, wrapping around", () => {
    let s = run(makeState(["one two one", "one"]), setFindQuery({ text: "one" }));
    expect(findState(s).matches).toHaveLength(3);
    expect(findState(s).current).toBe(0);
    s = run(s, findNext(1));
    expect(findState(s).current).toBe(0);
    expect(s.selection.from).toBe(1);
    s = run(s, findNext(1));
    expect(findState(s).current).toBe(1);
    s = run(s, findNext(1));
    s = run(s, findNext(1));
    expect(findState(s).current).toBe(0);
    s = run(s, findNext(-1));
    expect(findState(s).current).toBe(2);
  });

  it("replaces the selected match, then all matches in one step", () => {
    let s = run(makeState(["a cat, a cat", "cat"]), setFindQuery({ text: "cat" }));
    s = run(s, replaceMatch("dog"));
    expect(texts(s)).toEqual(["a cat, a cat", "cat"]);
    s = run(s, replaceMatch("dog"));
    expect(texts(s)).toEqual(["a dog, a cat", "cat"]);
    expect(findState(s).matches).toHaveLength(2);
    s = run(s, replaceAll("bird"));
    expect(texts(s)).toEqual(["a dog, a bird", "bird"]);
    expect(findState(s).matches).toHaveLength(0);
  });

  it("tracks replacements and skips deleted text", () => {
    let s = run(makeState(["cat cat"], true), setFindQuery({ text: "cat" }));
    s = run(s, replaceAll("dog"));
    const types = collectChanges(s.doc).map((c) => c.type);
    expect(types).toContain("insertion");
    expect(types).toContain("deletion");
    expect(findState(s).matches).toHaveLength(0);
    expect(findMatches(s.doc, { text: "dog" })).toHaveLength(2);
  });
});

describe("document info", () => {
  it("counts words and characters, excluding deletions", () => {
    let s = makeState(["Hello there world", "", "Two words"], true);
    expect(textStats(s.doc)).toEqual({
      words: 5,
      characters: 26,
      charactersNoSpaces: 23,
      paragraphs: 2,
    });
    s = run(s, (state, dispatch) => {
      dispatch?.(state.tr.delete(7, 13));
      return true;
    });
    expect(textStats(s.doc).words).toBe(4);
    expect(textStats(s.doc, 1, 6).words).toBe(1);
  });

  it("outlines headings and finds the one containing a position", () => {
    const { heading, paragraph } = schema.nodes;
    const doc = schema.node("doc", null, [
      heading.create({ level: 1 }, schema.text("Intro")),
      paragraph.create(null, schema.text("body")),
      heading.create({ level: 2 }, schema.text("Details")),
    ]);
    const outline = documentOutline(doc);
    expect(outline.map((e) => [e.level, e.text])).toEqual([
      [1, "Intro"],
      [2, "Details"],
    ]);
    expect(outlineIndexAt(outline, 10)).toBe(0);
    expect(outlineIndexAt(outline, outline[1].pos + 2)).toBe(1);
  });
});
