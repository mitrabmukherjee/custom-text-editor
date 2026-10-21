import type { DocumentLayout, PageInfo, SectionLayout } from "../format";
import { pageFrames, sectionLayouts } from "../format";
import { LETTER_PAGE_SETUP, type PageSetup } from "../pagination/pageSetup";
import type { ContentFrame, PageGeometry, PaginationContext } from "./types";

export function geometryForSection(
  sectionIndex: number,
  sections: SectionLayout[],
  defaultSetup: PageSetup,
): PageGeometry {
  const sec = sections[Math.min(sectionIndex, sections.length - 1)] ?? {};
  const page = sec.page ?? defaultSetup;
  const gap = page.gap ?? defaultSetup.gap;
  const cols = Math.max(1, sec.columns?.count ?? 1);
  const colSpace = sec.columns?.space ?? 0;
  const contentWidth = page.width - page.marginLeft - page.marginRight;
  const colWidth =
    cols > 1 ? (contentWidth - colSpace * (cols - 1)) / cols : contentWidth;
  const frames: ContentFrame[] = [];
  for (let c = 0; c < cols; c++) {
    frames.push({
      left: page.marginLeft + c * (colWidth + colSpace),
      width: colWidth,
    });
  }
  return {
    width: page.width,
    height: page.height,
    gap,
    sectionIndex,
    marginTop: page.marginTop,
    marginRight: page.marginRight,
    marginBottom: page.marginBottom,
    marginLeft: page.marginLeft,
    frames,
  };
}

export function buildPaginationContext(
  layout: DocumentLayout | null | undefined,
  sections: SectionLayout[],
  defaultSetup: PageSetup = LETTER_PAGE_SETUP,
): PaginationContext {
  const frameFn = pageFrames(defaultSetup, layout ?? null, sections);
  return {
    getPageGeometry: (pageIndex, pageInfo) => {
      const section = pageInfo?.section ?? 0;
      return geometryForSection(section, sections, defaultSetup);
    },
    getFrame: (page) => frameFn(page),
  };
}

/** Natural Y offset to the top of `pageIndex` in the stacked page list. */
export function pageNaturalTop(geometries: PageGeometry[], pageIndex: number): number {
  let y = 0;
  for (let p = 0; p < pageIndex; p++) {
    y += geometries[p]?.height ?? LETTER_PAGE_SETUP.height;
    y += geometries[p]?.gap ?? LETTER_PAGE_SETUP.gap;
  }
  return y;
}

export function stackHeightForGeometries(geometries: PageGeometry[], pageCount: number): number {
  if (pageCount <= 0) return 0;
  let h = 0;
  for (let p = 0; p < pageCount; p++) {
    const g = geometries[p] ?? geometries[0];
    h += g.height;
    if (p < pageCount - 1) h += g.gap;
  }
  return h;
}

/** Expand per-page geometries using section from pagination `pages` (filled during layout). */
export function geometriesFromPages(
  pages: PageInfo[],
  sections: SectionLayout[],
  defaultSetup: PageSetup,
): PageGeometry[] {
  return pages.map((p) => geometryForSection(p.section, sections, defaultSetup));
}

export function pageSetupForPage(
  pageInfo: PageInfo,
  sections: SectionLayout[],
  defaultSetup: PageSetup,
): PageSetup {
  const g = geometryForSection(pageInfo.section, sections, defaultSetup);
  return {
    width: g.width,
    height: g.height,
    marginTop: g.marginTop,
    marginRight: g.marginRight,
    marginBottom: g.marginBottom,
    marginLeft: g.marginLeft,
    gap: g.gap,
  };
}

export function activeTextWidth(
  state: { section: number },
  sections: SectionLayout[],
  defaultSetup: PageSetup,
): number {
  const g = geometryForSection(state.section, sections, defaultSetup);
  return g.frames[0]?.width ?? defaultSetup.width - defaultSetup.marginLeft - defaultSetup.marginRight;
}
