import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { NodeView, NodeViewConstructor } from "prosemirror-view";
import type { HFFrame } from "../format";

function frameInnerStyle(frame: HFFrame): string {
  const x = frame.x ?? 0;
  const y = frame.y ?? 0;
  const parts = [
    "position: absolute",
    `left: ${Math.round(x)}px`,
    `top: ${Math.round(y)}px`,
  ];
  if (frame.width) parts.push(`width: ${Math.round(frame.width)}px`);
  if (frame.height) parts.push(`min-height: ${Math.round(frame.height)}px`);
  return parts.join("; ");
}

class FramedParagraphView implements NodeView {
  readonly dom: HTMLDivElement;
  readonly contentDOM: HTMLParagraphElement;

  constructor(private node: PMNode) {
    this.dom = document.createElement("div");
    this.dom.className = "pm-body-frame-wrap";
    this.dom.style.cssText = "position: relative; height: 0; overflow: visible; margin: 0; padding: 0";
    this.contentDOM = document.createElement("p");
    this.contentDOM.className = "pm-body-frame";
    const frame = node.attrs.frame as HFFrame;
    this.contentDOM.style.cssText = frameInnerStyle(frame);
    this.applyBlockAttrs();
    this.dom.appendChild(this.contentDOM);
  }

  private applyBlockAttrs() {
    const format = this.node.attrs.format;
    if (format && typeof format === "object") {
      this.contentDOM.setAttribute("data-pm-format", JSON.stringify(format));
    }
    this.contentDOM.setAttribute("data-frame", JSON.stringify(this.node.attrs.frame));
    this.contentDOM.setAttribute("data-floating", "");
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    const frame = node.attrs.frame as HFFrame;
    this.contentDOM.style.cssText = frameInnerStyle(frame);
    this.applyBlockAttrs();
    return true;
  }

  ignoreMutation(record: MutationRecord) {
    return !this.contentDOM.contains(record.target as Node);
  }
}

/** Word body paragraphs with `w:framePr` (zero layout height, absolutely positioned content). */
export function framedParagraphPlugin() {
  return new Plugin({
    props: {
      nodeViews: {
        paragraph: ((node) => {
          if (node.attrs.frame && node.attrs.floating) {
            return new FramedParagraphView(node);
          }
          return undefined;
        }) as NodeViewConstructor,
      },
    },
  });
}
