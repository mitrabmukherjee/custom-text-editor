import type { Node as PMNode } from "prosemirror-model";

/** Map a text offset inside a textblock to an absolute document position. */
export function inlineDocPos(blockPos: number, block: PMNode, textOffset: number): number {
  let cur = 0;
  let pos = blockPos + 1;
  for (let i = 0; i < block.childCount; i++) {
    const child = block.child(i);
    const size = child.isText ? child.text!.length : 1;
    if (textOffset < cur + size) {
      return pos + (textOffset - cur);
    }
    cur += size;
    pos += child.nodeSize;
  }
  return pos - 1;
}

export function collectLineWrapPositions(
  doc: PMNode,
  lines: { anchor?: { pos: number; textOffset: number }; block?: number }[],
): number[] {
  const out: number[] = [];
  const seen = new Set<string>();
  for (const line of lines) {
    const a = line.anchor;
    if (!a || a.textOffset <= 0) continue;
    const block = doc.nodeAt(a.pos);
    if (!block?.isTextblock) continue;
    const key = `${a.pos}:${a.textOffset}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(inlineDocPos(a.pos, block, a.textOffset));
  }
  return out.sort((a, b) => a - b);
}
