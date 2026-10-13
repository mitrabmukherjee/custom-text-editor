/** Serialized ProseMirror mark, as produced by `Mark.toJSON()`. */
export type MarkJSON = {
  type: string;
  attrs?: Record<string, unknown>;
};

/** Serialized ProseMirror node, as produced by `Node.toJSON()`. */
export type NodeJSON = {
  type: string;
  attrs?: Record<string, unknown>;
  content?: NodeJSON[];
  marks?: MarkJSON[];
  text?: string;
};

export type DocJSON = NodeJSON & { type: "doc" };
