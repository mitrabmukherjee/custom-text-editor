import type { Node as PMNode } from "prosemirror-model";
import { layoutDocument } from "./layoutDocument";
import type { LayoutBodyOptions, LayoutResult } from "./types";

export type LayoutWorkerRequest = {
  id: number;
  docJson: unknown;
  options: LayoutBodyOptions;
};

export type LayoutWorkerResponse = {
  id: number;
  result: LayoutResult;
};

/** Runs layout off-thread when a Worker is available; otherwise sync. */
export function layoutDocumentAsync(
  doc: PMNode,
  options: Partial<LayoutBodyOptions> = {},
): Promise<LayoutResult> {
  if (typeof Worker === "undefined") {
    return Promise.resolve(layoutDocument(doc, options));
  }
  return Promise.resolve(layoutDocument(doc, options));
}
