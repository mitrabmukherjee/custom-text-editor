import {
  baseKeymap,
  chainCommands,
  createParagraphNear,
  exitCode,
  liftEmptyBlock,
  newlineInCode,
  splitBlockAs,
  toggleMark,
} from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { gapCursor } from "prosemirror-gapcursor";
import { history, redo, undo } from "prosemirror-history";
import {
  inputRules,
  textblockTypeInputRule,
  wrappingInputRule,
} from "prosemirror-inputrules";
import { keymap } from "prosemirror-keymap";
import { liftListItem, sinkListItem, splitListItem } from "prosemirror-schema-list";
import type { Command, Plugin } from "prosemirror-state";
import {
  clearFormatting,
  growFont,
  toggleCaps,
  toggleSubscript,
  toggleSuperscript,
  toggleUnderline,
} from "../character";
import { insertPageBreak } from "../commands";
import { findPlugin } from "../find";
import { tabInList } from "../lists";
import { indent, setParagraphFormat, toggleList } from "../paragraph";
import { schema } from "../schema";
import { nextParagraph } from "../styles";
import { shiftTabInTable, tabInTable, tablePlugins } from "../tables";
import { toggleTrackChanges } from "../track-changes/plugin";
import { contextualSpacingPlugin } from "./contextualSpacing";
import { framedParagraphPlugin } from "./framedParagraph";
import { imagePlugin, linkClickPlugin } from "./objects";
import { pastePlugin } from "./paste";
import { placeholderPlugin } from "./placeholder";

const insertHardBreak: Command = (state, dispatch) => {
  dispatch?.(
    state.tr.replaceSelectionWith(schema.nodes.hard_break.create()).scrollIntoView(),
  );
  return true;
};

const insertTab: Command = (state, dispatch) => {
  if (!state.selection.$from.parent.isTextblock) return false;
  dispatch?.(state.tr.replaceSelectionWith(schema.nodes.tab.create(), true).scrollIntoView());
  return true;
};

/**
 * Split like Word: at the end of a paragraph whose style names a following style
 * (headings → Normal) the new paragraph gets that style; otherwise it keeps the formatting.
 */
const splitFormatted = splitBlockAs((node, atEnd, $from) => {
  if (!atEnd) return null;
  const next = nextParagraph($from.doc, node);
  if (next) return next;
  const format = node.attrs.format as Record<string, unknown> | null;
  if (!format) return null;
  return {
    type: schema.nodes.paragraph,
    attrs: { format: { ...format, keepNext: undefined, pageBreakBefore: undefined } },
  };
});

function buildKeymap() {
  const { strong, em } = schema.marks;
  const { list_item } = schema.nodes;
  return keymap({
    "Mod-z": undo,
    "Shift-Mod-z": redo,
    "Mod-y": redo,
    "Mod-b": toggleMark(strong),
    "Mod-i": toggleMark(em),
    "Mod-Enter": insertPageBreak,
    "Shift-Enter": chainCommands(exitCode, insertHardBreak),
    Enter: chainCommands(
      splitListItem(list_item),
      newlineInCode,
      createParagraphNear,
      liftEmptyBlock,
      splitFormatted,
    ),
    Tab: chainCommands(tabInTable, tabInList, insertTab),
    "Shift-Tab": chainCommands(shiftTabInTable, liftListItem(list_item)),
    "Mod-Tab": sinkListItem(list_item),
    // Word's character shortcuts.
    "Mod-u": toggleUnderline(),
    "Shift-Mod-d": toggleUnderline("double"),
    "Shift-Mod-w": toggleUnderline("words"),
    "Mod-=": toggleSubscript,
    "Shift-Mod-=": toggleSuperscript,
    "Shift-Mod-a": toggleCaps("all"),
    "Shift-Mod-k": toggleCaps("small"),
    "Shift-Mod-.": growFont(1),
    "Shift-Mod-,": growFont(-1),
    "Mod-]": growFont(1, true),
    "Mod-[": growFont(-1, true),
    "Mod-Space": clearFormatting,
    // Word's paragraph shortcuts.
    "Mod-l": setParagraphFormat({ align: "left" }),
    "Mod-e": setParagraphFormat({ align: "center" }),
    "Mod-r": setParagraphFormat({ align: "right" }),
    "Mod-j": setParagraphFormat({ align: "justify" }),
    "Mod-m": indent(1),
    "Shift-Mod-m": indent(-1),
    "Mod-1": setParagraphFormat({ lineRule: "auto", lineValue: 1 }),
    "Mod-5": setParagraphFormat({ lineRule: "auto", lineValue: 1.5 }),
    "Mod-2": setParagraphFormat({ lineRule: "auto", lineValue: 2 }),
    "Mod-0": setParagraphFormat((f) => ({ spaceBefore: f.spaceBefore ? undefined : 16 })),
    "Shift-Mod-l": toggleList(schema.nodes.bullet_list),
    "Shift-Mod-e": toggleTrackChanges,
  });
}

function buildInputRules() {
  return inputRules({
    rules: [
      wrappingInputRule(/^\s*([-+*])\s$/, schema.nodes.bullet_list),
      wrappingInputRule(
        /^(\d+)\.\s$/,
        schema.nodes.ordered_list,
        (match) => ({ order: +match[1] }),
        (match, node) => node.childCount + node.attrs.order === +match[1],
      ),
      textblockTypeInputRule(/^(#{1,3})\s$/, schema.nodes.heading, (match) => ({
        level: match[1].length,
      })),
    ],
  });
}

export function basePlugins(options: { placeholder: string }): Plugin[] {
  return [
    buildInputRules(),
    buildKeymap(),
    keymap(baseKeymap),
    history(),
    dropCursor(),
    gapCursor(),
    placeholderPlugin(options.placeholder),
    contextualSpacingPlugin(),
    pastePlugin(),
    findPlugin(),
    ...tablePlugins(),
    imagePlugin(),
    linkClickPlugin(),
    framedParagraphPlugin(),
  ];
}
