import type { Mark, MarkType, Node as PMNode, Slice } from "prosemirror-model";
import { isHistoryTransaction } from "prosemirror-history";
import {
  TextSelection,
  type EditorState,
  type Transaction,
} from "prosemirror-state";
import {
  AddMarkStep,
  AttrStep,
  Mapping,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
} from "prosemirror-transform";
import type { ParagraphChange, ParagraphFormat } from "../format";
import { schema } from "../schema";
import { displayName, findStyle } from "../styles";
import { describeParagraphChange } from "./paragraphFormat";
import {
  FORMAT_MARK_NAMES,
  describeFormatChange,
  runFormatOf,
  type RunFormat,
} from "./formatting";
import {
  SKIP_TRACKING,
  trackChangesKey,
  type Author,
} from "./plugin";

let idCounter = 0;

export function newChangeId(): string {
  idCounter += 1;
  return `${Date.now().toString(36)}${idCounter.toString(36)}`;
}

const COPIED_META = ["addToHistory", "uiEvent", "paste", "pointer", "inputType"];

type Ctx = {
  author: Author;
  date: string;
};

type Segment = { from: number; to: number };

function isTrackingMark(mark: Mark) {
  const { insertion, deletion, formatChange } = schema.marks;
  return mark.type === insertion || mark.type === deletion || mark.type === formatChange;
}

function sameAuthorMark(
  node: PMNode | null | undefined,
  type: MarkType,
  authorId: string,
): Mark | null {
  const mark = node?.marks.find((m) => m.type === type);
  return mark && mark.attrs.authorId === authorId ? mark : null;
}

/** Attrs for a new change, reusing an adjacent same-author change so a typing run stays one change. */
function changeAttrs(
  doc: PMNode,
  from: number,
  to: number,
  type: MarkType,
  ctx: Ctx,
) {
  const before = doc.resolve(from).nodeBefore;
  const after = doc.resolve(to).nodeAfter;
  const existing =
    sameAuthorMark(before, type, ctx.author.id) ??
    sameAuthorMark(after, type, ctx.author.id);
  if (existing) return existing.attrs;
  return {
    id: newChangeId(),
    authorId: ctx.author.id,
    authorName: ctx.author.name,
    date: ctx.date,
  };
}

function hasInlineContent(doc: PMNode, from: number, to: number) {
  let found = false;
  doc.nodesBetween(from, to, (node) => {
    if (found) return false;
    if (node.isInline) found = true;
    return !found;
  });
  return found;
}

/**
 * Mark [from, to) as deleted instead of removing it. Text the current author
 * inserted is removed for real. Returns the (mapped) range that now holds the
 * struck-out text.
 */
function trackDeletion(tr: Transaction, from: number, to: number, ctx: Ctx) {
  const { insertion, deletion } = schema.marks;
  const toMark: Segment[] = [];
  const toRemove: Segment[] = [];

  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start >= end) return false;
    if (sameAuthorMark(node, insertion, ctx.author.id)) {
      toRemove.push({ from: start, to: end });
    } else if (!deletion.isInSet(node.marks)) {
      toMark.push({ from: start, to: end });
    }
    return false;
  });

  if (toMark.length) {
    const first = toMark[0];
    const last = toMark[toMark.length - 1];
    const attrs = changeAttrs(tr.doc, first.from, last.to, deletion, ctx);
    const mark = deletion.create(attrs);
    for (const seg of toMark) tr.addMark(seg.from, seg.to, mark);
  }

  const startMap = tr.mapping.maps.length;
  for (const seg of [...toRemove].reverse()) tr.delete(seg.from, seg.to);
  const local = tr.mapping.slice(startMap);
  return { from: local.map(from, -1), to: local.map(to, 1) };
}

/** Insert `slice` at `pos` and mark its inline content as an insertion. */
function trackInsertion(tr: Transaction, pos: number, slice: Slice, ctx: Ctx) {
  const { insertion, deletion, formatChange } = schema.marks;
  const sizeBefore = tr.doc.content.size;
  tr.replace(pos, pos, slice);
  const end = pos + (tr.doc.content.size - sizeBefore);
  if (end <= pos || !hasInlineContent(tr.doc, pos, end)) return end;
  tr.removeMark(pos, end, deletion);
  tr.removeMark(pos, end, formatChange);
  tr.removeMark(pos, end, insertion);
  tr.addMark(pos, end, insertion.create(changeAttrs(tr.doc, pos, end, insertion, ctx)));
  return end;
}

/**
 * Apply a bold/italic/font step and record it as a formatting revision. Formatting the
 * author's own insertion is not a separate revision, as in Word.
 */
function trackFormatting(tr: Transaction, step: AddMarkStep | RemoveMarkStep, ctx: Ctx) {
  const { insertion, formatChange } = schema.marks;
  const { from, to } = step;
  const segments: (Segment & { old: RunFormat; id?: string })[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    if (sameAuthorMark(node, insertion, ctx.author.id)) return false;
    const existing = formatChange.isInSet(node.marks);
    segments.push({
      from: Math.max(pos, from),
      to: Math.min(pos + node.nodeSize, to),
      old: (existing?.attrs.old as RunFormat | null) ?? runFormatOf(node.marks),
      id: existing?.attrs.id as string | undefined,
    });
    return false;
  });
  if (tr.maybeStep(step).failed) return;

  const ids = new Map<string, string>();
  for (const seg of segments) {
    const node = tr.doc.nodeAt(seg.from);
    if (!node) continue;
    const description = describeFormatChange(seg.old, runFormatOf(node.marks));
    tr.removeMark(seg.from, seg.to, formatChange);
    if (!description) continue;
    const key = `${description}|${JSON.stringify(seg.old)}`;
    const id = seg.id ?? ids.get(key) ?? newChangeId();
    ids.set(key, id);
    tr.addMark(
      seg.from,
      seg.to,
      formatChange.create({
        id,
        authorId: ctx.author.id,
        authorName: ctx.author.name,
        date: ctx.date,
        description,
        old: seg.old,
      }),
    );
  }
}

function blockOf(node: PMNode): { type: string; level?: number } {
  return node.type === schema.nodes.heading
    ? { type: "heading", level: node.attrs.level as number }
    : { type: node.type.name };
}

const blockName = (block: { type: string; level?: number }) =>
  block.type === "heading" ? `Heading ${block.level}` : "Normal";

/**
 * Apply a step that changes one paragraph's `format` or block type (a style change) and
 * record it as a paragraph formatting revision against the paragraph before any revision.
 */
function trackParagraphStep(tr: Transaction, step: AttrStep | ReplaceAroundStep, pos: number, ctx: Ctx) {
  const node = tr.doc.nodeAt(pos);
  if (tr.maybeStep(step).failed || !node?.isTextblock) return;
  const after = tr.doc.nodeAt(pos);
  if (!after?.isTextblock) return;
  const existing = node.attrs.paraChange as ParagraphChange | null;
  const old = existing ? existing.old : ((node.attrs.format as ParagraphFormat | null) ?? null);
  const oldBlock = existing ? (existing.block ?? blockOf(node)) : blockOf(node);
  const nowBlock = blockOf(after);
  const now = (after.attrs.format as ParagraphFormat | null) ?? null;
  const styleName = (id: string) => {
    const style = findStyle(tr.doc, id);
    return style ? displayName(style) : id;
  };
  let description = describeParagraphChange(old, now, styleName);
  const blockChanged = JSON.stringify(oldBlock) !== JSON.stringify(nowBlock);
  if (blockChanged && !description.startsWith("Style:")) {
    description = [`Style: ${blockName(nowBlock)}`, description].filter(Boolean).join(", ");
  }
  const change: ParagraphChange | null = description
    ? {
        id: existing?.id ?? newChangeId(),
        authorId: ctx.author.id,
        authorName: ctx.author.name,
        date: ctx.date,
        old,
        ...(blockChanged ? { block: oldBlock } : {}),
        description,
      }
    : null;
  tr.setNodeAttribute(pos, "paraChange", change);
}

/** A `setNodeMarkup` on a textblock: one ReplaceAroundStep wrapping the whole node. */
function textblockMarkupStep(step: ReplaceAroundStep, doc: PMNode): boolean {
  const node = doc.nodeAt(step.from);
  const replacement = step.slice.content.firstChild;
  return (
    !!node?.isTextblock &&
    step.to === step.from + node.nodeSize &&
    step.gapFrom === step.from + 1 &&
    step.gapTo === step.to - 1 &&
    step.insert === 1 &&
    step.slice.content.childCount === 1 &&
    !!replacement?.isTextblock
  );
}

/**
 * Rebuild a user transaction so its edits become tracked changes. Character formatting
 * becomes a formatting revision; structural edits (splits, joins, wrapping) and other
 * steps pass through untracked.
 */
export function trackTransaction(
  tr: Transaction,
  state: EditorState,
  now: () => Date = () => new Date(),
): Transaction {
  if (
    !tr.docChanged ||
    tr.getMeta(SKIP_TRACKING) ||
    tr.getMeta("composition") !== undefined ||
    isHistoryTransaction(tr)
  ) {
    return tr;
  }
  const ts = trackChangesKey.getState(state);
  if (!ts?.enabled) return tr;

  const ctx: Ctx = { author: ts.author, date: now().toISOString() };
  const out = state.tr;
  const steps = tr.steps;
  let pureDeletion: Segment | null = null;

  for (let i = 0; i < steps.length; i++) {
    const back = new Mapping();
    for (let j = i - 1; j >= 0; j--) back.appendMap(steps[j].getMap().invert());
    back.appendMapping(out.mapping);
    const step = steps[i].map(back);
    if (!step) continue;

    if (
      (step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
      FORMAT_MARK_NAMES.includes(step.mark.type.name)
    ) {
      trackFormatting(out, step, ctx);
      continue;
    }
    if (step instanceof AttrStep && step.attr === "format") {
      trackParagraphStep(out, step, step.pos, ctx);
      continue;
    }
    if (step instanceof ReplaceAroundStep && textblockMarkupStep(step, out.doc)) {
      trackParagraphStep(out, step, step.from, ctx);
      continue;
    }
    if (!(step instanceof ReplaceStep)) {
      out.maybeStep(step);
      continue;
    }

    const { from, to, slice } = step;
    let insertAt = from;
    if (from < to) {
      if (hasInlineContent(out.doc, from, to)) {
        const struck = trackDeletion(out, from, to, ctx);
        insertAt = struck.to;
        if (slice.size === 0 && steps.length === 1) pureDeletion = struck;
      } else {
        // Structure-only change (e.g. joining paragraphs): apply untracked,
        // but still mark any inserted text.
        const sizeBefore = out.doc.content.size;
        out.replace(from, to, slice);
        const end = to + (out.doc.content.size - sizeBefore);
        if (end > from && hasInlineContent(out.doc, from, end)) {
          const { insertion } = schema.marks;
          out.addMark(from, end, insertion.create(changeAttrs(out.doc, from, end, insertion, ctx)));
        }
        continue;
      }
    }
    if (slice.size > 0) trackInsertion(out, insertAt, slice, ctx);
  }

  if (tr.selectionSet || pureDeletion) {
    const forward = new Mapping();
    for (let j = steps.length - 1; j >= 0; j--) {
      forward.appendMap(steps[j].getMap().invert());
    }
    forward.appendMapping(out.mapping);
    let anchor = forward.map(tr.selection.anchor, 1);
    let head = forward.map(tr.selection.head, 1);
    if (pureDeletion) {
      const sel = state.selection;
      const backward = sel.empty && sel.head === (steps[0] as ReplaceStep).to;
      anchor = head = backward ? pureDeletion.from : pureDeletion.to;
    }
    const size = out.doc.content.size;
    anchor = Math.max(0, Math.min(anchor, size));
    head = Math.max(0, Math.min(head, size));
    out.setSelection(
      TextSelection.between(out.doc.resolve(anchor), out.doc.resolve(head)),
    );
  }

  if (tr.storedMarksSet) {
    out.setStoredMarks(tr.storedMarks?.filter((m) => !isTrackingMark(m)) ?? null);
  }
  if (tr.scrolledIntoView) out.scrollIntoView();
  for (const key of COPIED_META) {
    const value = tr.getMeta(key);
    if (value !== undefined) out.setMeta(key, value);
  }
  out.setTime(tr.time);
  return out;
}
