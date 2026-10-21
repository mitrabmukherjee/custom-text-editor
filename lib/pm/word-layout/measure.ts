import { PT_TO_PX } from "../format";
import type { LayoutRun } from "./types";

const SERIF = /times|georgia|garamond|cambria|book|serif|palatino|century/i;
const MONO = /courier|consolas|mono|lucida console/i;
const SANS = /arial|helvetica|calibri|verdana|sans/i;

export type TextMeasurer = {
  measureRun(text: string, run: LayoutRun): number;
  measureChar(ch: string, run: LayoutRun): number;
};

function familyClass(family: string): "mono" | "serif" | "sans" {
  if (MONO.test(family)) return "mono";
  if (SANS.test(family)) return "sans";
  if (SERIF.test(family)) return "serif";
  return "serif";
}

/** Proportional widths tuned for Word-like line breaks in Node/tests. */
export function createProportionalMeasurer(): TextMeasurer {
  const charWidth = (ch: string, run: LayoutRun) => {
    const px = run.fontSizePt * PT_TO_PX;
    const kind = familyClass(run.fontFamily);
    let w = px * (kind === "mono" ? 0.6 : kind === "sans" ? 0.48 : 0.5);
    if (run.bold) w *= 1.03;
    if (ch === " ") w = px * (kind === "mono" ? 0.6 : 0.25);
    if (ch === "\t") w = px * 4;
    if (ch === "\ufffc") w = px * 3;
    return w;
  };
  return {
    measureChar: charWidth,
    measureRun: (text, run) => {
      let sum = 0;
      for (const ch of text) sum += charWidth(ch, run);
      return sum;
    },
  };
}

let canvasMeasurer: TextMeasurer | null = null;

function canvasFont(run: LayoutRun): string {
  const style = run.italic ? "italic" : "normal";
  const weight = run.bold ? "bold" : "normal";
  const family = run.fontFamily.includes(" ") ? `"${run.fontFamily}"` : run.fontFamily;
  return `${style} ${weight} ${run.fontSizePt}pt ${family}, serif`;
}

/** Canvas `measureText` in the browser (embedded @font-face fonts). */
export function createCanvasMeasurer(): TextMeasurer | null {
  if (typeof document === "undefined") return null;
  if (canvasMeasurer) return canvasMeasurer;
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  canvasMeasurer = {
    measureChar(ch, run) {
      ctx.font = canvasFont(run);
      return ctx.measureText(ch).width;
    },
    measureRun(text, run) {
      ctx.font = canvasFont(run);
      return ctx.measureText(text).width;
    },
  };
  return canvasMeasurer;
}

export function defaultMeasurer(): TextMeasurer {
  return createCanvasMeasurer() ?? createProportionalMeasurer();
}

/** HarfBuzz hook: swap measurer when WASM shaper is loaded. */
let harfbuzzMeasurer: TextMeasurer | null = null;

export function setHarfBuzzMeasurer(m: TextMeasurer | null) {
  harfbuzzMeasurer = m;
}

export function activeMeasurer(): TextMeasurer {
  return harfbuzzMeasurer ?? defaultMeasurer();
}
