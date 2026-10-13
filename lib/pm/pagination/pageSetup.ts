/** Page geometry in CSS px (96 px per inch). */
export type PageSetup = {
  width: number;
  height: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  /** Visual gap between pages on screen. */
  gap: number;
};

/** US Letter, 1 inch margins (OOXML 12240x15840 twips, 1440 twip margins). */
export const LETTER_PAGE_SETUP: PageSetup = {
  width: 816,
  height: 1056,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
  gap: 24,
};

/** A4 at 96 dpi (OOXML 11906×16838 twips). */
export const A4_PAGE_SETUP: PageSetup = {
  width: 794,
  height: 1123,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
  gap: 24,
};

/** US Legal 8.5×14 in. */
export const LEGAL_PAGE_SETUP: PageSetup = {
  width: 816,
  height: 1344,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
  gap: 24,
};

export const PAGE_SETUP_PRESETS: { id: string; label: string; setup: PageSetup }[] = [
  { id: "letter", label: "Letter (8.5\" × 11\")", setup: LETTER_PAGE_SETUP },
  { id: "a4", label: "A4", setup: A4_PAGE_SETUP },
  { id: "legal", label: "Legal (8.5\" × 14\")", setup: LEGAL_PAGE_SETUP },
];

export const TWIPS_PER_PX = 15;

export function contentHeight(setup: PageSetup): number {
  return setup.height - setup.marginTop - setup.marginBottom;
}

/** Distance from one page's top edge to the next page's top edge. */
export function pageStride(setup: PageSetup): number {
  return setup.height + setup.gap;
}

export function stackHeight(setup: PageSetup, pageCount: number): number {
  return pageCount * pageStride(setup) - setup.gap;
}
