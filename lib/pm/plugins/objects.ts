import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorView, NodeView } from "prosemirror-view";
import { fitImage, readImageFile, setImageSize } from "../objects";
import { safeHref, schema } from "../schema";
import { SKIP_TRACKING } from "../track-changes";

const MIN_IMAGE = 16;

/** An image with Word's corner resize handle while it is selected. */
class ImageView implements NodeView {
  readonly dom: HTMLSpanElement;
  private readonly img: HTMLImageElement;
  private readonly handle: HTMLSpanElement;

  constructor(
    private node: PMNode,
    private readonly view: EditorView,
    private readonly getPos: () => number | undefined,
  ) {
    this.dom = document.createElement("span");
    this.dom.className = "pm-image-wrap";
    this.img = document.createElement("img");
    this.img.className = "pm-image";
    this.handle = document.createElement("span");
    this.handle.className = "pm-image-handle";
    this.handle.addEventListener("pointerdown", (e) => this.startResize(e));
    this.dom.append(this.img, this.handle);
    this.render();
  }

  private render() {
    const { src, alt, title, width, height, anchorX, anchorY } = this.node.attrs;
    this.img.src = src;
    this.img.alt = alt ?? "";
    if (title) this.img.title = title;
    this.img.style.width = width ? `${width}px` : "";
    this.img.style.height = height ? `${height}px` : "";
    const anchored = anchorX != null || anchorY != null;
    this.dom.classList.toggle("pm-image-anchored", anchored);
    if (anchored) {
      this.dom.style.cssText = "";
      this.img.style.position = "absolute";
      this.img.style.left = anchorX != null ? `${anchorX}px` : "0";
      this.img.style.top = anchorY != null ? `${anchorY}px` : "0";
    } else {
      this.dom.style.cssText = "";
      this.img.style.position = "";
      this.img.style.left = "";
      this.img.style.top = "";
    }
  }

  private startResize(e: PointerEvent) {
    e.preventDefault();
    e.stopPropagation();
    const scale = this.dom.offsetWidth ? this.dom.getBoundingClientRect().width / this.dom.offsetWidth : 1;
    const startX = e.clientX;
    const startWidth = this.img.offsetWidth;
    const ratio = startWidth ? this.img.offsetHeight / startWidth : 1;
    let width = startWidth;
    const move = (ev: PointerEvent) => {
      width = Math.max(MIN_IMAGE, startWidth + (ev.clientX - startX) / scale);
      this.img.style.width = `${width}px`;
      this.img.style.height = `${width * ratio}px`;
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      const pos = this.getPos();
      if (pos !== undefined && width !== startWidth) {
        // Resizing is formatting, not an insertion and deletion of the picture.
        setImageSize(pos, width, width * ratio)(this.view.state, (tr) =>
          this.view.dispatch(tr.setMeta(SKIP_TRACKING, true)),
        );
      }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add("ProseMirror-selectednode");
  }

  deselectNode() {
    this.dom.classList.remove("ProseMirror-selectednode");
  }

  stopEvent(e: Event) {
    return e.target === this.handle;
  }

  ignoreMutation() {
    return true;
  }
}

const PICTURE_TYPES = /^image\/(png|jpeg|gif|bmp)$/;

function pictureFiles(data: DataTransfer | null) {
  return Array.from(data?.files ?? []).filter((file) => PICTURE_TYPES.test(file.type));
}

/** Width of the text column at `pos`, so dropped pictures fit like Word's. */
function columnWidth(view: EditorView, pos: number) {
  try {
    const { node } = view.domAtPos(pos);
    const el = node instanceof HTMLElement ? node : node.parentElement;
    return el?.closest<HTMLElement>("p, h1, h2, h3, h4, h5, h6, li, td, th")?.clientWidth ?? 0;
  } catch {
    return 0;
  }
}

async function insertPictures(view: EditorView, files: File[], pos?: number) {
  const images = await Promise.all(files.map(readImageFile));
  const { state } = view;
  const at = pos ?? state.selection.from;
  const max = columnWidth(view, at);
  const nodes = images.map((image) =>
    schema.nodes.image.create({ ...image, ...fitImage(image.width, image.height, max) }),
  );
  const tr = pos === undefined ? state.tr.deleteSelection() : state.tr;
  const insertAt = pos === undefined ? tr.selection.from : tr.mapping.map(pos);
  view.dispatch(tr.insert(insertAt, nodes).scrollIntoView());
}

export function imagePlugin() {
  return new Plugin({
    props: {
      nodeViews: {
        image: (node, view, getPos) => new ImageView(node, view, getPos),
      },
      handlePaste(view, event) {
        const files = pictureFiles(event.clipboardData);
        // Pasted HTML (with its own <img> tags) or text wins over a bitmap copy of it.
        if (!files.length || event.clipboardData?.getData("text/html")) return false;
        void insertPictures(view, files);
        return true;
      },
      handleDrop(view, event) {
        const files = pictureFiles(event.dataTransfer);
        if (!files.length) return false;
        const pos = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos;
        void insertPictures(view, files, pos);
        event.preventDefault();
        return true;
      },
    },
  });
}

/** Ctrl+Click (Cmd+Click) follows a link in a new tab, as in Word. */
export function linkClickPlugin() {
  return new Plugin({
    props: {
      handleClick(_view, _pos, event) {
        if (!(event.ctrlKey || event.metaKey)) return false;
        const link = (event.target as HTMLElement | null)?.closest<HTMLElement>("a.pm-link");
        const href = safeHref(link?.dataset.href ?? null);
        if (!href || href.startsWith("#")) return false;
        window.open(href, "_blank", "noopener,noreferrer");
        return true;
      },
    },
  });
}
