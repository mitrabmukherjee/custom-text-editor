import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type Command, type EditorState } from "prosemirror-state";
import { schema } from "./schema";
import type { Author } from "./track-changes/plugin";

export type CommentReply = {
  id: string;
  authorId: string;
  authorName: string;
  date: string;
  text: string;
};

/** A comment and its replies, stored in the doc's `comments` attribute. */
export type CommentThread = CommentReply & {
  resolved?: boolean;
  replies?: CommentReply[];
};

/** A thread with the range of text it is anchored to. */
export type CommentEntry = CommentThread & { from: number; to: number; quote: string };

let counter = 0;

function newId(prefix: string) {
  counter += 1;
  return `${prefix}${Date.now().toString(36)}${counter.toString(36)}`;
}

export function commentThreads(doc: PMNode): Record<string, CommentThread> {
  return (doc.attrs.comments as Record<string, CommentThread> | null) ?? {};
}

const cache = new WeakMap<PMNode, CommentEntry[]>();

/** Threads that still have anchored text, in document order. */
export function collectComments(doc: PMNode): CommentEntry[] {
  const cached = cache.get(doc);
  if (cached) return cached;
  const threads = commentThreads(doc);
  const ranges = new Map<string, { from: number; to: number; quote: string }>();
  const { comment } = schema.marks;
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const mark of node.marks) {
      if (mark.type !== comment) continue;
      const id = String(mark.attrs.id);
      const text = node.isText ? node.text! : " ";
      const range = ranges.get(id);
      if (range) {
        range.to = pos + node.nodeSize;
        range.quote += text;
      } else {
        ranges.set(id, { from: pos, to: pos + node.nodeSize, quote: text });
      }
    }
    return false;
  });
  const out: CommentEntry[] = [];
  for (const [id, range] of ranges) {
    const thread = threads[id];
    if (thread) out.push({ ...thread, ...range });
  }
  out.sort((a, b) => a.from - b.from);
  cache.set(doc, out);
  return out;
}

/** The selection, or the word at the cursor when nothing is selected (as Word does). */
function commentRange(state: EditorState): { from: number; to: number } | null {
  const { from, to, empty, $from } = state.selection;
  if (!empty) return { from, to };
  if (!$from.parent.isTextblock) return null;
  const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, "\ufffc");
  const offset = $from.parentOffset;
  const isWord = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_'’-]/u.test(ch);
  let start = offset;
  let end = offset;
  while (isWord(text[start - 1])) start--;
  while (isWord(text[end])) end++;
  if (start === end) return null;
  return { from: $from.start() + start, to: $from.start() + end };
}

export function canAddComment(state: EditorState) {
  return !!commentRange(state);
}

function setThreads(state: EditorState, threads: Record<string, CommentThread>) {
  return state.tr.setDocAttribute("comments", Object.keys(threads).length ? threads : null);
}

/** Word's New Comment: anchor `text` to the selection (or the word at the cursor). */
export function addComment(text: string, author: Author, date = new Date().toISOString()): Command {
  return (state, dispatch) => {
    const range = commentRange(state);
    if (!range || !text.trim()) return false;
    if (dispatch) {
      const id = newId("c");
      const thread: CommentThread = { id, authorId: author.id, authorName: author.name, date, text };
      const tr = setThreads(state, { ...commentThreads(state.doc), [id]: thread });
      tr.addMark(range.from, range.to, schema.marks.comment.create({ id, authorId: author.id }));
      tr.setSelection(TextSelection.create(tr.doc, range.to));
      dispatch(tr);
    }
    return true;
  };
}

function updateThread(id: string, update: (thread: CommentThread) => CommentThread): Command {
  return (state, dispatch) => {
    const threads = commentThreads(state.doc);
    const thread = threads[id];
    if (!thread) return false;
    dispatch?.(setThreads(state, { ...threads, [id]: update(thread) }));
    return true;
  };
}

export const editComment = (id: string, text: string) =>
  updateThread(id, (thread) => ({ ...thread, text }));

export const resolveComment = (id: string, resolved: boolean) =>
  updateThread(id, (thread) => ({ ...thread, resolved: resolved || undefined }));

export function replyToComment(
  id: string,
  text: string,
  author: Author,
  date = new Date().toISOString(),
): Command {
  return updateThread(id, (thread) => ({
    ...thread,
    replies: [
      ...(thread.replies ?? []),
      { id: newId("r"), authorId: author.id, authorName: author.name, date, text },
    ],
  }));
}

/** Delete a thread and its anchor. */
export function deleteComment(id: string): Command {
  return (state, dispatch) => {
    const threads = commentThreads(state.doc);
    if (!threads[id]) return false;
    if (dispatch) {
      const rest = { ...threads };
      delete rest[id];
      const tr = setThreads(state, rest);
      const { comment } = schema.marks;
      state.doc.descendants((node, pos) => {
        const mark = node.marks.find((m) => m.type === comment && m.attrs.id === id);
        if (mark) tr.removeMark(pos, pos + node.nodeSize, mark);
        return !node.isInline;
      });
      dispatch(tr);
    }
    return true;
  };
}

/** Comments whose text overlaps the selection (or contains the cursor). */
export function commentsInSelection(state: EditorState): CommentEntry[] {
  const { from, to } = state.selection;
  return collectComments(state.doc).filter((c) => c.from <= to && c.to >= from);
}
