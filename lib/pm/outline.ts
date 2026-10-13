import type { Node as PMNode } from "prosemirror-model";
import { schema } from "./schema";

export type OutlineEntry = { pos: number; level: number; text: string };

const cache = new WeakMap<PMNode, OutlineEntry[]>();

/** Headings in document order, for the navigation pane. */
export function documentOutline(doc: PMNode): OutlineEntry[] {
  const cached = cache.get(doc);
  if (cached) return cached;
  const out: OutlineEntry[] = [];
  const { heading } = schema.nodes;
  const { deletion } = schema.marks;
  doc.descendants((node, pos) => {
    if (node.type === heading) {
      let text = "";
      node.forEach((child) => {
        if (child.isText && !deletion.isInSet(child.marks)) text += child.text;
        else if (!child.isText) text += " ";
      });
      text = text.replace(/\s+/g, " ").trim();
      if (text) out.push({ pos, level: node.attrs.level as number, text });
      return false;
    }
    return !node.isTextblock;
  });
  cache.set(doc, out);
  return out;
}

/** Index of the heading the position falls under, or -1. */
export function outlineIndexAt(outline: OutlineEntry[], pos: number) {
  let found = -1;
  for (let i = 0; i < outline.length && outline[i].pos <= pos; i++) found = i;
  return found;
}
