import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { schema } from "./schema";
import { SKIP_TRACKING } from "./track-changes/plugin";

export type FindQuery = { text: string; matchCase?: boolean; wholeWord?: boolean };
export type FindMatch = { from: number; to: number };

type FindState = {
  query: FindQuery;
  matches: FindMatch[];
  /** Index of the highlighted match, or -1. */
  current: number;
  decorations: DecorationSet;
};

type FindMeta = { query?: FindQuery; current?: number };

export const findKey = new PluginKey<FindState>("find");

const EMPTY: FindState = {
  query: { text: "" },
  matches: [],
  current: -1,
  decorations: DecorationSet.empty,
};

const WORD_CHAR = /[\p{L}\p{N}_]/u;
/** Stands in for anything that isn't searchable text, so matches can't span it. */
const GAP = "\u0000";

/** Lower-case per UTF-16 unit so indexes keep matching document positions. */
function fold(text: string) {
  let out = "";
  for (const ch of text) {
    const lower = ch.toLowerCase();
    out += lower.length === ch.length ? lower : ch;
  }
  return out;
}

/** Every match of `query`, like Word: within one paragraph, skipping tracked deletions. */
export function findMatches(doc: PMNode, query: FindQuery): FindMatch[] {
  if (!query.text) return [];
  const needle = query.matchCase ? query.text : fold(query.text);
  const { deletion } = schema.marks;
  const out: FindMatch[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    let text = "";
    const positions: number[] = [];
    node.forEach((child, offset) => {
      const start = pos + 1 + offset;
      if (child.isText && !deletion.isInSet(child.marks)) {
        text += child.text;
        for (let i = 0; i < child.text!.length; i++) positions.push(start + i);
      } else {
        text += GAP;
        positions.push(start);
      }
    });
    const hay = query.matchCase ? text : fold(text);
    for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + needle.length)) {
      const end = i + needle.length;
      if (query.wholeWord && (WORD_CHAR.test(text[i - 1] ?? "") || WORD_CHAR.test(text[end] ?? ""))) {
        continue;
      }
      out.push({ from: positions[i], to: positions[end - 1] + 1 });
    }
    return false;
  });
  return out;
}

function decorate(doc: PMNode, matches: FindMatch[], current: number) {
  if (!matches.length) return DecorationSet.empty;
  return DecorationSet.create(
    doc,
    matches.map((m, i) =>
      Decoration.inline(m.from, m.to, {
        class: i === current ? "pm-find-match pm-find-match--current" : "pm-find-match",
      }),
    ),
  );
}

/** First match starting at or after `pos` (wrapping to the first), or -1. */
function matchFrom(matches: FindMatch[], pos: number) {
  if (!matches.length) return -1;
  const i = matches.findIndex((m) => m.from >= pos);
  return i < 0 ? 0 : i;
}

export function findPlugin() {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: () => EMPTY,
      apply(tr, prev, _old, state) {
        const meta = tr.getMeta(findKey) as FindMeta | undefined;
        if (!meta && (!tr.docChanged || !prev.query.text)) return prev;
        const query = meta?.query ?? prev.query;
        const matches =
          meta?.query || tr.docChanged ? findMatches(state.doc, query) : prev.matches;
        const current =
          meta?.current ?? (matches.length ? matchFrom(matches, state.selection.from) : -1);
        return { query, matches, current, decorations: decorate(state.doc, matches, current) };
      },
    },
    props: {
      decorations: (state) => findKey.getState(state)?.decorations,
    },
  });
}

export function findState(state: EditorState): FindState {
  return findKey.getState(state) ?? EMPTY;
}

function findTr(state: EditorState, meta: FindMeta) {
  return state.tr.setMeta(findKey, meta).setMeta("addToHistory", false).setMeta(SKIP_TRACKING, true);
}

/** Search for `query` (an empty text clears the highlights). */
export function setFindQuery(query: FindQuery): Command {
  return (state, dispatch) => {
    dispatch?.(findTr(state, { query }));
    return true;
  };
}

/** Select the next (or previous) match after the selection, wrapping around the document. */
export function findNext(direction: 1 | -1 = 1): Command {
  return (state, dispatch) => {
    const { matches } = findState(state);
    if (!matches.length) return false;
    const sel = state.selection;
    let i: number;
    if (direction > 0) {
      i = matchFrom(matches, sel.empty ? sel.from : sel.from + 1);
    } else {
      i = -1;
      for (let k = matches.length - 1; k >= 0; k--) {
        if (matches[k].from < sel.from) {
          i = k;
          break;
        }
      }
      if (i < 0) i = matches.length - 1;
    }
    if (dispatch) {
      const m = matches[i];
      dispatch(
        findTr(state, { current: i })
          .setSelection(TextSelection.create(state.doc, m.from, m.to))
          .scrollIntoView(),
      );
    }
    return true;
  };
}

/** The match the selection covers exactly, if any. */
function selectedMatch(state: EditorState) {
  const { from, to } = state.selection;
  return findState(state).matches.find((m) => m.from === from && m.to === to);
}

/**
 * Word's Replace: replace the selected match, or find the next one first.
 * Returns true when text was replaced; select the following match with `findNext`.
 */
export function replaceMatch(replacement: string): Command {
  return (state, dispatch) => {
    const match = selectedMatch(state);
    if (!match) {
      findNext(1)(state, dispatch);
      return false;
    }
    if (dispatch) {
      const tr = state.tr.insertText(replacement, match.from, match.to);
      tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(match.to)));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Replace every match in one undoable step. */
export function replaceAll(replacement: string): Command {
  return (state, dispatch) => {
    const { matches } = findState(state);
    if (!matches.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (let i = matches.length - 1; i >= 0; i--) {
        tr.insertText(replacement, matches[i].from, matches[i].to);
      }
      dispatch(tr);
    }
    return true;
  };
}
