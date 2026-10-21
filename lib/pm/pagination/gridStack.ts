/** Place the next doc-grid row on the stacked slot (Word pitch), not DOM sub-pixel drift. */
export function gridStackTop(naturalTop: number, nextSlotTop: number | undefined) {
  return nextSlotTop ?? naturalTop;
}
