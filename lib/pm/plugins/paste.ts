import { Fragment, Slice, type Mark, type Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { schema } from "../schema";

/**
 * Word's paste options. "keep": keep source formatting. "merge": keep bold, italic,
 * underline and the like but take fonts, colours and paragraph formatting from the
 * destination. "text": plain text only. "auto": keep for this editor's own content and
 * Word, merge for other HTML (web pages bring theme colours that would be invisible in print).
 */
export type PasteMode = "auto" | "keep" | "merge" | "text";

type Source = "internal" | "word" | "html" | "text";

export const pasteKey = new PluginKey<{ mode: PasteMode }>("paste");

const MERGE_DROPPED = new Set(["textStyle", "textColor", "highlight"]);
/** Revision marks and comment anchors, which must not be duplicated by pasting. */
const TRACKING = new Set(["insertion", "deletion", "formatChange", "comment"]);

/** Strip what Word and Office put on the clipboard besides the content. */
export function cleanWordHtml(html: string): string {
  return html
    .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<\/?o:p[^>]*>/gi, "")
    .replace(/<\/?(?:v|w|m):[^>]*>/gi, "")
    .replace(/\s*mso-[a-z-]+\s*:\s*[^;"']+;?/gi, "")
    .replace(/\scolor\s*:\s*windowtext;?/gi, "");
}

function sourceOf(html: string): Source {
  if (/data-pm-slice/.test(html)) return "internal";
  if (/urn:schemas-microsoft-com:office|class="?Mso|mso-/i.test(html)) return "word";
  return "html";
}

type Blocks = "keep" | "clear" | "plain";

/**
 * Rebuild a slice: drop tracked deletions, filter marks, and keep, clear (`format`) or
 * flatten (`plain`: headings become paragraphs) the blocks.
 */
export function transformSlice(slice: Slice, keepMark: (mark: Mark) => boolean, blocks: Blocks): Slice {
  const { deletion } = schema.marks;
  const visit = (fragment: Fragment): Fragment => {
    const nodes: PMNode[] = [];
    fragment.forEach((node) => {
      if (node.isInline && deletion.isInSet(node.marks)) return;
      if (node.isInline) {
        nodes.push(node.mark(node.marks.filter((m) => !TRACKING.has(m.type.name) && keepMark(m))));
        return;
      }
      const content = visit(node.content);
      if (node.isTextblock && blocks === "plain") {
        nodes.push(schema.nodes.paragraph.create(null, content));
      } else if (node.isTextblock) {
        const attrs = { ...node.attrs, paraChange: null };
        if (blocks === "clear" && "format" in attrs) attrs.format = null;
        nodes.push(node.type.create(attrs, content, node.marks));
      } else {
        nodes.push(node.copy(content));
      }
    });
    return Fragment.fromArray(nodes);
  };
  return new Slice(visit(slice.content), slice.openStart, slice.openEnd);
}

/** Paste handling: clean Word HTML and apply the paste mode. Tracking marks never survive a paste. */
export function pastePlugin(initial: PasteMode = "auto") {
  let source: Source = "text";
  return new Plugin<{ mode: PasteMode }>({
    key: pasteKey,
    state: {
      init: () => ({ mode: initial }),
      apply: (tr, prev) => (tr.getMeta(pasteKey) as { mode: PasteMode } | undefined) ?? prev,
    },
    props: {
      transformPastedHTML(html) {
        source = sourceOf(html);
        return source === "word" ? cleanWordHtml(html) : html;
      },
      transformPastedText(text) {
        source = "text";
        return text;
      },
      transformPasted(slice, view) {
        const setting = pasteKey.getState(view.state)?.mode ?? "auto";
        const mode =
          setting === "auto" ? (source === "html" ? "merge" : "keep") : setting;
        source = "text";
        if (mode === "text") return transformSlice(slice, () => false, "plain");
        if (mode === "merge") return transformSlice(slice, (m) => !MERGE_DROPPED.has(m.type.name), "clear");
        return transformSlice(slice, () => true, "keep");
      },
    },
  });
}
