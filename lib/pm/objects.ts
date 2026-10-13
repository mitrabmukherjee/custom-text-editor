import type { Mark, ResolvedPos } from "prosemirror-model";
import { NodeSelection, TextSelection, type Command, type EditorState } from "prosemirror-state";
import { schema } from "./schema";

export type LinkInfo = { href: string; from: number; to: number; text: string };

/** The run of text around `$pos` carrying the same link as the text at it. */
function linkRange($pos: ResolvedPos): LinkInfo | null {
  const { link } = schema.marks;
  const parent = $pos.parent;
  const start = $pos.start();
  const find = (marks: readonly Mark[]) => marks.find((m) => m.type === link);
  const at =
    find($pos.marks()) ?? find($pos.nodeAfter?.marks ?? []) ?? find($pos.nodeBefore?.marks ?? []);
  if (!at) return null;
  const href = String(at.attrs.href);
  let found: LinkInfo | null = null;
  let from = -1;
  let to = -1;
  const endRun = () => {
    if (!found && from >= 0 && $pos.pos >= from && $pos.pos <= to) {
      found = { href, from, to, text: parent.textBetween(from - start, to - start) };
    }
    from = -1;
  };
  parent.forEach((child, offset) => {
    if (find(child.marks)?.attrs.href !== href) return endRun();
    if (from < 0) from = start + offset;
    to = start + offset + child.nodeSize;
  });
  endRun();
  return found;
}

/** The link at the cursor or covering the selection. */
export function linkAt(state: EditorState): LinkInfo | null {
  const { from, to, $from } = state.selection;
  const range = linkRange($from);
  if (!range) return null;
  return to <= range.to && from >= range.from ? range : null;
}

/**
 * Word's Insert Hyperlink: link the selection (or the link at the cursor) to `href`. `text`
 * replaces the linked text when it differs; with nothing selected it is inserted.
 */
export function setLink(href: string, text?: string): Command {
  return (state, dispatch) => {
    const url = href.trim();
    if (!url) return false;
    const mark = schema.marks.link.create({ href: url });
    const existing = state.selection.empty ? linkAt(state) : null;
    const from = existing?.from ?? state.selection.from;
    const to = existing?.to ?? state.selection.to;
    const current = state.doc.textBetween(from, to);
    const label = text?.trim() ? text : current || url;
    if (dispatch) {
      const tr = state.tr;
      if (label !== current) {
        const marks = from < to ? state.doc.resolve(from + 1).marks() : state.selection.$from.marks();
        tr.replaceWith(from, to, schema.text(label, mark.addToSet(marks.filter((m) => m.type !== mark.type))));
      } else {
        tr.removeMark(from, to, schema.marks.link).addMark(from, to, mark);
      }
      const end = from + label.length;
      dispatch(tr.setSelection(TextSelection.create(tr.doc, end)).removeStoredMark(schema.marks.link));
    }
    return true;
  };
}

export const removeLink: Command = (state, dispatch) => {
  const range = state.selection.empty ? linkAt(state) : state.selection;
  if (!range || range.from === range.to) return false;
  if (!state.doc.rangeHasMark(range.from, range.to, schema.marks.link)) return false;
  dispatch?.(state.tr.removeMark(range.from, range.to, schema.marks.link));
  return true;
};

export type ImageAttrs = { src: string; alt?: string; width?: number | null; height?: number | null };

/** Word fits a newly inserted picture to the text column, keeping its aspect ratio. */
export function fitImage(width: number, height: number, max: number) {
  return width > max && max > 0
    ? { width: Math.round(max), height: Math.round((height * max) / width) }
    : { width, height };
}

/** Read an image file as a data URL plus its pixel size. */
export function readImageFile(file: File): Promise<ImageAttrs & { width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => {
      const src = String(reader.result);
      const img = new Image();
      img.onerror = () => reject(new Error(`Unreadable image: ${file.name}`));
      img.onload = () => resolve({ src, alt: file.name, width: img.naturalWidth, height: img.naturalHeight });
      img.src = src;
    };
    reader.readAsDataURL(file);
  });
}

export function insertImage(attrs: ImageAttrs): Command {
  return (state, dispatch) => {
    if (!attrs.src) return false;
    const node = schema.nodes.image.create(attrs);
    dispatch?.(state.tr.replaceSelectionWith(node).scrollIntoView());
    return true;
  };
}

/** Resize the image at `pos` (px), keeping it selected. */
export function setImageSize(pos: number, width: number, height: number): Command {
  return (state, dispatch) => {
    const node = state.doc.nodeAt(pos);
    if (node?.type !== schema.nodes.image) return false;
    if (dispatch) {
      const tr = state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        width: Math.round(width),
        height: Math.round(height),
      });
      dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
    }
    return true;
  };
}
