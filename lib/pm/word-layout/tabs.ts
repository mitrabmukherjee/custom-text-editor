import type { ParagraphFormat, TabStop } from "../format";
import { nextTabStop } from "../pagination/plugin";
import type { TextMeasurer } from "./measure";
import type { LayoutRun } from "./types";

export type TabAlign = TabStop["align"];

export function tabWidthToStop(
  x: number,
  stop: TabStop,
  segmentWidth: number,
  measurer: TextMeasurer,
  run: LayoutRun,
): number {
  const target = stop.pos;
  switch (stop.align) {
    case "center":
      return Math.max(0, target - x - segmentWidth / 2);
    case "right":
      return Math.max(0, target - x - segmentWidth);
    case "decimal": {
      const dot = segmentWidth * 0.35;
      return Math.max(0, target - x - dot);
    }
    default:
      return Math.max(0, target - x);
  }
}

export function resolveTabWidth(
  x: number,
  format: ParagraphFormat | null,
  defaultTab: number,
  measurer: TextMeasurer,
  run: LayoutRun,
  pendingText: string,
): number {
  const stops = (format?.tabs ?? []).map((t) => t.pos);
  if (format?.indentLeft && (format.indentFirstLine ?? 0) < 0) stops.push(format.indentLeft);
  const custom = (format?.tabs ?? [])
    .filter((t) => t.pos > x + 0.5)
    .sort((a, b) => a.pos - b.pos)[0];
  if (custom) {
    const segW = measurer.measureRun(pendingText, run);
    return tabWidthToStop(x, custom, segW, measurer, run);
  }
  const next = nextTabStop(x, format, defaultTab);
  return Math.max(0, next - x);
}
