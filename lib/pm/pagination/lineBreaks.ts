import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

/** Word may break after a hyphen before a digit ("25-127116"); Chrome's UAX #14 breaker won't. */
const HYPHEN_BEFORE_DIGIT = /-(?=\d)/g;

const OBJECT_REPLACEMENT = "\ufffc";

function createWbr() {
  return document.createElement("wbr");
}

function breakOpportunities(doc: PMNode) {
  const decos: Decoration[] = [];
  doc.descendants((block, blockPos) => {
    if (!block.isTextblock) return true;
    let text = "";
    block.forEach((child) => {
      text += child.isText ? child.text! : OBJECT_REPLACEMENT;
    });
    for (const m of text.matchAll(HYPHEN_BEFORE_DIGIT)) {
      decos.push(
        Decoration.widget(blockPos + 1 + m.index + 1, createWbr, {
          side: -1,
          ignoreSelection: true,
          key: "wbr",
        }),
      );
    }
    return false;
  });
  return DecorationSet.create(doc, decos);
}

const lineBreaksKey = new PluginKey<DecorationSet>("wordLineBreaks");

/** Adds the line-break opportunities Word has but the browser lacks. */
export function wordLineBreaksPlugin() {
  return new Plugin<DecorationSet>({
    key: lineBreaksKey,
    state: {
      init: (_config, state) => breakOpportunities(state.doc),
      apply: (tr, prev) => (tr.docChanged ? breakOpportunities(tr.doc) : prev),
    },
    props: {
      decorations: (state) => lineBreaksKey.getState(state),
    },
  });
}
