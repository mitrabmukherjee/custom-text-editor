import { createHash } from "node:crypto";
import type { EngineLayoutLine, LayoutGolden } from "./types";
import type { ComputedLayout } from "../pagination/computePageBreaks";

export function hashLineAnchors(lines: EngineLayoutLine[]): string {
  const parts: string[] = [];
  for (const line of lines) {
    const a = line.anchor;
    if (a) parts.push(`${a.pos}:${a.textOffset}`);
    else parts.push(`?:${Math.round(line.top)}`);
  }
  return createHash("sha256").update(parts.join("|")).digest("hex").slice(0, 16);
}

export function layoutGoldenFrom(
  fixture: string,
  lines: EngineLayoutLine[],
  pagination: ComputedLayout,
): LayoutGolden {
  const perPageLines = new Array(pagination.pageCount).fill(0);
  for (const p of pagination.placements) {
    perPageLines[p.page] = (perPageLines[p.page] ?? 0) + 1;
  }
  return {
    fixture,
    pageCount: pagination.pageCount,
    lineCount: lines.length,
    lineAnchorHash: hashLineAnchors(lines),
    perPageLines,
    breakLineIndices: pagination.breaks.map((b) => b.lineIndex),
  };
}

export function assertGolden(
  actual: LayoutGolden,
  expected: LayoutGolden,
): { ok: boolean; diff: string } {
  const keys = ["pageCount", "lineCount", "lineAnchorHash"] as const;
  for (const k of keys) {
    if (actual[k] !== expected[k]) {
      return { ok: false, diff: `${k}: expected ${expected[k]}, got ${actual[k]}` };
    }
  }
  if (JSON.stringify(actual.perPageLines) !== JSON.stringify(expected.perPageLines)) {
    return {
      ok: false,
      diff: `perPageLines: expected ${JSON.stringify(expected.perPageLines)}, got ${JSON.stringify(actual.perPageLines)}`,
    };
  }
  if (expected.breakLineIndices) {
    const a = actual.breakLineIndices ?? [];
    const e = expected.breakLineIndices;
    if (JSON.stringify(a) !== JSON.stringify(e)) {
      return {
        ok: false,
        diff: `breakLineIndices: expected ${JSON.stringify(e)}, got ${JSON.stringify(a)}`,
      };
    }
  }
  return { ok: true, diff: "" };
}
