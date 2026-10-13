/** Word's "By author" sequence (red, blue, green, …), lightened to read on dark pages. */
const PALETTE = [
  "#f05454",
  "#5b9cf5",
  "#4cc46b",
  "#c27ef0",
  "#e8a33c",
  "#3cc4c4",
  "#f06ab0",
  "#9aa7ff",
];

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
