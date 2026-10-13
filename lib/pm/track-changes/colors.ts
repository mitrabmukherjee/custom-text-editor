/**
 * Word's default "By author" revision colours (w:color), tuned for contrast on dark pages.
 * Order matches Word: red, blue, green, purple, dark red, teal, olive, gray.
 */
const PALETTE = [
  "#ff6b6b",
  "#6eb5ff",
  "#6bd46b",
  "#c490ff",
  "#ff9f7a",
  "#5ed4d4",
  "#d4a5ff",
  "#a8b0c8",
];

/** Margin change bars and connectors use Word's primary revision red. */
export const REVISION_BAR_COLOR = "#e81123";

const assigned = new Map<string, string>();

/** Colour per author, handed out in order of first appearance like Word. */
export function authorColor(authorId: string): string {
  let color = assigned.get(authorId);
  if (!color) {
    color = PALETTE[assigned.size % PALETTE.length];
    assigned.set(authorId, color);
  }
  return color;
}
