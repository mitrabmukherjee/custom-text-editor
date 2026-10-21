import type { Node as PMNode } from "prosemirror-model";

/** Text that participates in layout when revision markup matches the editor view. */
export function blockHasVisibleText(
  block: PMNode,
  options: { includeDeletionText: boolean },
): boolean {
  if (!block.isTextblock) return false;
  let text = "";
  block.forEach((child) => {
    if (!child.isText) return;
    if (!options.includeDeletionText && child.marks.some((m) => m.type.name === "deletion")) {
      return;
    }
    text += child.text ?? "";
  });
  return text.trim().length > 0;
}

export function textNodeHiddenForLayout(node: Text): boolean {
  const del = node.parentElement?.closest("del.tc-deletion");
  if (!del) return false;
  return getComputedStyle(del).display === "none";
}
