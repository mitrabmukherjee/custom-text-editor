import type { DocJSON, NodeJSON } from "@/lib/pm/json";

/** Which side of tracked changes to export. */
export type PlainTextMode = "accepted" | "original";

/**
 * Split plain transcript text into paragraph blocks (blank-line separated).
 */
export function splitPlainParagraphs(plain: string): string[] {
  return String(plain ?? "")
    .split(/\n{2,}/)
    .map((p) => p.trim())
    .filter(Boolean);
}

function blockToInline(text: string): NodeJSON[] {
  const out: NodeJSON[] = [];
  text.split("\n").forEach((line, i) => {
    if (i > 0) out.push({ type: "hard_break" });
    if (line) out.push({ type: "text", text: line });
  });
  return out;
}

/**
 * Build a ProseMirror JSON document from plain text. Blank lines separate
 * paragraphs; single newlines become hard breaks.
 */
export function plainTextToDoc(plain: string): DocJSON {
  const blocks = splitPlainParagraphs(plain);
  if (!blocks.length) {
    return { type: "doc", content: [{ type: "paragraph" }] };
  }
  return {
    type: "doc",
    content: blocks.map((text) => {
      const content = blockToInline(text);
      return content.length
        ? { type: "paragraph", content }
        : { type: "paragraph" };
    }),
  };
}

/**
 * Walk ProseMirror JSON and export plain text with paragraph breaks.
 * "accepted" drops deleted text; "original" drops inserted text.
 */
export function docJsonToPlainText(
  doc: DocJSON,
  options: { mode?: PlainTextMode } = {},
): string {
  const mode = options.mode ?? "accepted";
  const parts: string[] = [];
  for (const node of doc.content ?? []) {
    collectBlock(node, mode, parts);
  }
  return parts.join("\n\n").replace(/\n{3,}/g, "\n\n").trim();
}

function collectBlock(node: NodeJSON, mode: PlainTextMode, parts: string[]) {
  if (node.type === "speaker_line") {
    const label = (node.attrs?.label as string) ?? "";
    const body = inlineText(node, mode);
    const line = `${label}${body ? ` ${body}` : ""}`.trim();
    if (line) parts.push(line);
  } else if (node.type === "qa_line") {
    const role = node.attrs?.role === "a" ? "A." : "Q.";
    const body = inlineText(node, mode);
    const line = body ? `${role} ${body}` : role;
    if (line.trim()) parts.push(line);
  } else if (node.type === "paragraph" || node.type === "heading") {
    const line = inlineText(node, mode);
    if (line.trim()) parts.push(line);
    else if (parts.length) parts.push("");
  } else if (node.type === "bullet_list" || node.type === "ordered_list") {
    for (const item of node.content ?? []) {
      if (item.type === "list_item") {
        const line = (item.content ?? [])
          .map((p) => inlineText(p, mode))
          .filter(Boolean)
          .join(" ");
        if (line.trim()) parts.push(line);
      }
    }
  } else if (node.type === "table") {
    for (const row of node.content ?? []) {
      const cells = (row.content ?? []).map((cell) => {
        const inner: string[] = [];
        for (const block of cell.content ?? []) collectBlock(block, mode, inner);
        return inner.filter(Boolean).join(" ");
      });
      if (cells.some((c) => c.trim())) parts.push(cells.join("\t"));
    }
  }
}

function hidden(node: NodeJSON, mode: PlainTextMode): boolean {
  const hiddenMark = mode === "accepted" ? "deletion" : "insertion";
  return !!node.marks?.some((m) => m.type === hiddenMark);
}

function inlineText(node: NodeJSON, mode: PlainTextMode): string {
  if (!node.content?.length) return "";
  return node.content.map((child) => textFromInline(child, mode)).join("");
}

function textFromInline(node: NodeJSON, mode: PlainTextMode): string {
  if (hidden(node, mode)) return "";
  if (node.type === "text") return node.text ?? "";
  if (node.type === "hard_break") return "\n";
  if (node.type === "tab") return "\t";
  if (node.type === "timestamp") return (node.attrs?.time as string) ?? "";
  if (node.content?.length) {
    return node.content.map((child) => textFromInline(child, mode)).join("");
  }
  return "";
}
