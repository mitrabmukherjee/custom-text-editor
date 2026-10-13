import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { ParagraphFormat } from "../format";

function formatOf(node: PMNode | null | undefined): ParagraphFormat | null {
  return node?.isTextblock ? ((node.attrs.format as ParagraphFormat | null) ?? null) : null;
}

/** Same style for contextual spacing: the Word style id, or both unstyled. */
function sameStyle(a: PMNode, b: PMNode) {
  return a.type === b.type && (formatOf(a)?.styleId ?? null) === (formatOf(b)?.styleId ?? null);
}

function build(doc: PMNode) {
  const decos: Decoration[] = [];
  doc.descendants((node, pos, parent) => {
    if (!node.isTextblock || !parent) return true;
    if (!formatOf(node)?.contextualSpacing) return false;
    const $pos = doc.resolve(pos);
    const index = $pos.index();
    const prev = index > 0 ? parent.child(index - 1) : null;
    const next = index < parent.childCount - 1 ? parent.child(index + 1) : null;
    const classes = [
      prev && sameStyle(prev, node) ? "pm-ctx-before" : "",
      next && sameStyle(next, node) ? "pm-ctx-after" : "",
    ].filter(Boolean);
    if (classes.length) {
      decos.push(Decoration.node(pos, pos + node.nodeSize, { class: classes.join(" ") }));
    }
    return false;
  });
  return DecorationSet.create(doc, decos);
}

/**
 * Word's "Don't add space between paragraphs of the same style": drop a paragraph's
 * space before/after next to a paragraph of the same style.
 */
export function contextualSpacingPlugin() {
  return new Plugin<DecorationSet>({
    state: {
      init: (_, state) => build(state.doc),
      apply: (tr, prev) => (tr.docChanged ? build(tr.doc) : prev),
    },
    props: {
      decorations(state) {
        return this.getState(state);
      },
    },
  });
}
