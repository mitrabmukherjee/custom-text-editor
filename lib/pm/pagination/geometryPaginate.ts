import type { PageFrame, PageInfo } from "../format";
import type { PageGeometry, PaginationContext } from "../word-layout/types";
import { pageNaturalTop, stackHeightForGeometries } from "../word-layout/pageGeometry";
import { LETTER_PAGE_SETUP, type PageSetup } from "./pageSetup";
import { geometryForSection } from "../word-layout/pageGeometry";
import type { SectionLayout } from "../format";
import {
  type ComputedBreak,
  type ComputedLayout,
  type LayoutLine,
  type LinePlacement,
  type PaginateOptions,
  type BreakKind,
} from "./computePageBreaks";

function blockBounds(lines: LayoutLine[]) {
  const first = new Map<number, number>();
  const last = new Map<number, number>();
  lines.forEach((line, i) => {
    if (line.block === undefined) return;
    if (!first.has(line.block)) first.set(line.block, i);
    last.set(line.block, i);
  });
  const firstOf = (i: number) =>
    lines[i].block === undefined ? i : first.get(lines[i].block!)!;
  const lastOf = (i: number) =>
    lines[i].block === undefined ? i : last.get(lines[i].block!)!;
  return { firstOf, lastOf };
}

function pageBottomNatural(
  ctx: PaginationContext,
  pages: PageInfo[],
  geometries: PageGeometry[],
  pageIndex: number,
  line: LayoutLine,
): number {
  const info = pages[pageIndex];
  const geo = geometries[pageIndex] ?? ctx.getPageGeometry(pageIndex, info);
  const frame = ctx.getFrame(
    info ?? {
      section: line.section ?? 0,
      first: pageIndex === 0,
      number: pageIndex + 1,
    },
  );
  return pageNaturalTop(geometries, pageIndex) + geo.height - frame.bottom;
}

/**
 * Word-style pagination with per-page width/height and section-aware geometry.
 */
export function computePageBreaksWithContext(
  lines: LayoutLine[],
  ctx: PaginationContext,
  options: PaginateOptions = {},
  defaultSetup: PageSetup = LETTER_PAGE_SETUP,
): ComputedLayout {
  const { firstOf, lastOf } = blockBounds(lines);
  const breaks: ComputedBreak[] = [];
  const placements: LinePlacement[] = new Array(lines.length);
  const pages: PageInfo[] = [];
  const geometries: PageGeometry[] = [];
  let page = 0;
  let shift = 0;
  let pageStart = 0;
  let forceNext: BreakKind = false;

  const ensureGeometry = (pageIndex: number) => {
    while (geometries.length <= pageIndex) {
      const section = pages[pageIndex]?.section ?? lines[0]?.section ?? 0;
      geometries.push(
        ctx.getPageGeometry(pageIndex, pages[pageIndex] ?? { section, first: pageIndex === 0, number: pageIndex + 1 }),
      );
    }
  };

  const startPage = (index: number, section: number) => {
    const prev = pages[pages.length - 1];
    const first = !prev || prev.section !== section;
    const restart = first ? options.numberStart?.(section) : undefined;
    while (pages.length < index) {
      const last = pages[pages.length - 1];
      pages.push({
        section: last?.section ?? section,
        first: false,
        number: (last?.number ?? 0) + 1,
      });
      ensureGeometry(pages.length - 1);
    }
    const before = pages[pages.length - 1];
    pages[index] = { section, first, number: restart ?? (before ? before.number + 1 : 1) };
    ensureGeometry(index);
  };
  startPage(0, lines[0]?.section ?? 0);

  const lineFrame = (line: LayoutLine, pageIndex: number): PageFrame => {
    const section = line.section ?? 0;
    const base = pages[pageIndex];
    return ctx.getFrame({
      section,
      number: base?.number ?? 1,
      first: base?.section === section ? (base?.first ?? false) : false,
    });
  };

  const keepTarget = (i: number): number => {
    let t = i;
    const line = lines[t];
    const first = firstOf(t);
    const last = lastOf(t);
    if (line.keepLines) t = first;
    else if (line.widowControl && last > first) {
      if (last === t && last - first >= 2) t -= 1;
      if (t - first === 1) t = first;
    }
    while (t === firstOf(t) && t > 0 && lines[t - 1].keepNext && !lines[t - 1].breakAfter) {
      t = firstOf(t - 1);
    }
    return t > pageStart ? t : i;
  };

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    ensureGeometry(page);
    const pageBottom = pageBottomNatural(ctx, pages, geometries, page, line);
    const overflows = line.bottom - (line.slack ?? 0) + shift > pageBottom + 0.5;
    const breakBefore = !!line.breakBefore && i > pageStart && firstOf(i) === i;

    if (forceNext || breakBefore || (overflows && i > pageStart)) {
      const forced = !!forceNext || breakBefore;
      const target = forced ? i : keepTarget(i);
      page += 1;
      if (forceNext === "odd" && (page + 1) % 2 === 0) page += 1;
      if (forceNext === "even" && (page + 1) % 2 === 1) page += 1;
      startPage(page, lines[target].section ?? 0);
      ensureGeometry(page);
      const geo = geometries[page]!;
      const frame = lineFrame(lines[target], page);
      const nextTop = pageNaturalTop(geometries, page) + frame.top;
      const height = nextTop - (lines[target].top + shift);
      breaks.push({ lineIndex: target, height, forced });
      shift += height;
      pageStart = target;
      forceNext = false;
      i = target;
    }

    const placed = lines[i];
    const geo = geometries[page] ?? geometryForSection(placed.section ?? 0, [], defaultSetup);
    const pageTop = pageNaturalTop(geometries, page);
    placements[i] = {
      page,
      top: placed.top + shift - pageTop,
      bottom: placed.bottom + shift - pageTop,
      section: placed.section ?? 0,
      rowHeight: placed.rowHeight,
    };
    if (placed.breakAfter) forceNext = placed.breakAfter;
    i += 1;
  }

  return { breaks, pageCount: page + 1, placements, pages };
}

export function stackHeightFromContext(
  ctx: PaginationContext,
  pages: PageInfo[],
  pageCount: number,
  defaultSetup: PageSetup,
): number {
  const geometries: PageGeometry[] = [];
  for (let p = 0; p < pageCount; p++) {
    geometries.push(ctx.getPageGeometry(p, pages[p]));
  }
  return stackHeightForGeometries(geometries, pageCount);
}
