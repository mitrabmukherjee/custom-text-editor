import type { Mark, Node as PMNode } from "prosemirror-model";
import {
  TextSelection,
  type Command,
  type EditorState,
  type Transaction,
} from "prosemirror-state";
import type { ParagraphChange } from "../format";
import { schema } from "../schema";
import { FORMAT_MARK_NAMES, runFormatMarks, type RunFormat } from "./formatting";
import { SKIP_TRACKING } from "./plugin";

export type ChangeType = "insertion" | "deletion" | "format" | "paragraph";

export type TrackedChange = {
  id: string;
  type: ChangeType;
  authorId: string;
  authorName: string;
  date: string;
  text: string;
  from: number;
  to: number;
  /** Formatting revisions only, e.g. "Font: Italic". */
  description?: string;
};

type MarkRange = { from: number; to: number; mark: Mark };

const CHANGE_TYPE: Record<string, ChangeType> = {
  insertion: "insertion",
  deletion: "deletion",
  formatChange: "format",
};

function changeMarks(node: PMNode): Mark[] {
  return node.marks.filter((m) => m.type.name in CHANGE_TYPE);
}

/** Contiguous ranges carrying a tracking mark, merged per mark. */
function findChangeRanges(
  doc: PMNode,
  include: (mark: Mark) => boolean,
): MarkRange[] {
  const ranges: MarkRange[] = [];
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const mark of changeMarks(node)) {
      if (!include(mark)) continue;
      const end = pos + node.nodeSize;
      const prev = ranges.findLast((r) => r.mark.eq(mark));
      if (prev && prev.to === pos) prev.to = end;
      else ranges.push({ from: pos, to: end, mark });
    }
    return false;
  });
  return ranges;
}

/** Paragraphs carrying a paragraph formatting revision. */
function paragraphChanges(doc: PMNode) {
  const out: { pos: number; node: PMNode; change: ParagraphChange }[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const change = node.attrs.paraChange as ParagraphChange | null;
    if (change) out.push({ pos, node, change });
    return false;
  });
  return out;
}

/** All tracked changes in document order, one entry per change id and type. */
export function collectChanges(doc: PMNode): TrackedChange[] {
  const byKey = new Map<string, TrackedChange>();
  for (const { pos, node, change } of paragraphChanges(doc)) {
    const key = `paragraph:${change.id}`;
    const existing = byKey.get(key);
    if (existing) {
      existing.to = pos + node.nodeSize - 1;
      existing.text += `\n${node.textContent}`;
      continue;
    }
    byKey.set(key, {
      id: change.id,
      type: "paragraph",
      authorId: change.authorId,
      authorName: change.authorName,
      date: change.date,
      text: node.textContent,
      from: pos + 1,
      to: pos + node.nodeSize - 1,
      description: change.description,
    });
  }
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const mark of changeMarks(node)) {
      const type = CHANGE_TYPE[mark.type.name];
      const key = `${type}:${mark.attrs.id}`;
      const text = node.isText
        ? (node.text ?? "")
        : node.type.name === "tab"
          ? "\t"
          : "\u21b5";
      const existing = byKey.get(key);
      if (existing) {
        existing.text += text;
        existing.to = pos + node.nodeSize;
      } else {
        byKey.set(key, {
          id: String(mark.attrs.id),
          type,
          authorId: String(mark.attrs.authorId),
          authorName: String(mark.attrs.authorName),
          date: String(mark.attrs.date),
          text,
          from: pos,
          to: pos + node.nodeSize,
          ...(type === "format" ? { description: String(mark.attrs.description) } : {}),
        });
      }
    }
    return false;
  });
  return [...byKey.values()].sort((a, b) => a.from - b.from);
}

/** Put back the character formatting a formatting revision recorded. */
function restoreFormat(tr: Transaction, from: number, to: number, old: RunFormat | null) {
  if (!old) return;
  for (const name of FORMAT_MARK_NAMES) tr.removeMark(from, to, schema.marks[name]);
  for (const mark of runFormatMarks(old)) tr.addMark(from, to, schema.markFromJSON(mark));
}

function resolveChanges(
  ids: Set<string> | null,
  action: "accept" | "reject",
): Command {
  return (state: EditorState, dispatch) => {
    const ranges = findChangeRanges(
      state.doc,
      (mark) => !ids || ids.has(String(mark.attrs.id)),
    );
    const paragraphs = paragraphChanges(state.doc).filter((p) => !ids || ids.has(p.change.id));
    if (!ranges.length && !paragraphs.length) return false;
    if (!dispatch) return true;

    const tr = state.tr.setMeta(SKIP_TRACKING, true);
    for (const { pos, node, change } of paragraphs) {
      const block = change.block;
      if (action === "reject" && block && schema.nodes[block.type]) {
        tr.setNodeMarkup(pos, schema.nodes[block.type], {
          ...node.attrs,
          level: block.level ?? node.attrs.level,
          format: change.old,
          paraChange: null,
        });
        continue;
      }
      if (action === "reject") tr.setNodeAttribute(pos, "format", change.old);
      tr.setNodeAttribute(pos, "paraChange", null);
    }
    const removeText: MarkRange[] = [];
    for (const range of ranges) {
      if (range.mark.type === schema.marks.formatChange) {
        tr.removeMark(range.from, range.to, range.mark);
        if (action === "reject") {
          restoreFormat(tr, range.from, range.to, range.mark.attrs.old as RunFormat | null);
        }
        continue;
      }
      const isInsertion = range.mark.type === schema.marks.insertion;
      const dropText = action === "accept" ? !isInsertion : isInsertion;
      if (dropText) removeText.push(range);
      else tr.removeMark(range.from, range.to, range.mark);
    }
    removeText.sort((a, b) => b.from - a.from);
    for (const range of removeText) tr.delete(range.from, range.to);
    dispatch(tr);
    return true;
  };
}

export const acceptChange = (id: string) => resolveChanges(new Set([id]), "accept");
export const rejectChange = (id: string) => resolveChanges(new Set([id]), "reject");
export const acceptChanges = (ids: Iterable<string>) => resolveChanges(new Set(ids), "accept");
export const rejectChanges = (ids: Iterable<string>) => resolveChanges(new Set(ids), "reject");

/** Changes overlapping the selection, or containing the cursor (Word's Accept/Reject buttons). */
export function changesInSelection(state: EditorState): TrackedChange[] {
  const { from, to } = state.selection;
  return collectChanges(state.doc).filter((c) => c.from <= to && c.to >= from);
}
export const acceptAllChanges = resolveChanges(null, "accept");
export const rejectAllChanges = resolveChanges(null, "reject");

/** Select a change and scroll it into view. */
export function goToChange(id: string): Command {
  return (state, dispatch) => {
    const change = collectChanges(state.doc).find((c) => c.id === id);
    if (!change) return false;
    dispatch?.(
      state.tr
        .setSelection(TextSelection.create(state.doc, change.from, change.to))
        .scrollIntoView(),
    );
    return true;
  };
}
