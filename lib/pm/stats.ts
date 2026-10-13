import type { Node as PMNode } from "prosemirror-model";
import { schema } from "./schema";

/** Word's Word Count figures; tracked deletions don't count. */
export type TextStats = {
  words: number;
  characters: number;
  charactersNoSpaces: number;
  paragraphs: number;
};

const cache = new WeakMap<PMNode, TextStats>();

function blockText(block: PMNode, from: number, to: number) {
  let text = "";
  block.forEach((child, offset) => {
    const start = offset;
    const end = offset + child.nodeSize;
    if (end <= from || start >= to) return;
    if (child.isText) {
      if (schema.marks.deletion.isInSet(child.marks)) return;
      text += child.text!.slice(Math.max(0, from - start), to - start);
    } else if (child.type === schema.nodes.tab) {
      text += "\t";
    } else if (child.type === schema.nodes.hard_break) {
      text += "\n";
    }
  });
  return text;
}

/** Counts for the whole document, or for the text between `from` and `to`. */
export function textStats(doc: PMNode, from?: number, to?: number): TextStats {
  const whole = from === undefined || to === undefined;
  if (whole) {
    const cached = cache.get(doc);
    if (cached) return cached;
  }
  const start = from ?? 0;
  const end = to ?? doc.content.size;
  const stats: TextStats = { words: 0, characters: 0, charactersNoSpaces: 0, paragraphs: 0 };
  doc.nodesBetween(start, end, (node, pos) => {
    if (!node.isTextblock) return true;
    const text = blockText(node, start - pos - 1, end - pos - 1);
    if (text.trim()) stats.paragraphs += 1;
    stats.words += text.split(/\s+/).filter(Boolean).length;
    stats.characters += text.replace(/\n/g, "").length;
    stats.charactersNoSpaces += text.replace(/\s/g, "").length;
    return false;
  });
  if (whole) cache.set(doc, stats);
  return stats;
}
