import { Plugin } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

export function placeholderPlugin(text: string) {
  return new Plugin({
    props: {
      decorations(state) {
        const { doc } = state;
        const first = doc.firstChild;
        if (
          doc.childCount !== 1 ||
          !first ||
          !first.isTextblock ||
          first.content.size !== 0
        ) {
          return null;
        }
        return DecorationSet.create(doc, [
          Decoration.node(0, first.nodeSize, {
            class: "is-editor-empty",
            "data-placeholder": text,
          }),
        ]);
      },
    },
  });
}
