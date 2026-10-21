import type { PageFrame } from "../format";
import type { PageGeometry } from "./types";
import type { EngineLayoutLine } from "./types";

/** Tracks natural Y per column while laying out multi-column sections (Word reading order). */
export class ColumnFlow {
  private pageTop = 0;
  private col = 0;
  private colYs: number[] = [];
  private readonly bodyHeight: number;
  private readonly stride: number;
  private readonly frameTop: number;

  constructor(
    private readonly geo: PageGeometry,
    frame: PageFrame,
    startY: number,
  ) {
    this.stride = geo.height + geo.gap;
    this.frameTop = frame.top;
    this.bodyHeight = geo.height - frame.top - frame.bottom;
    this.pageTop = startY;
    this.colYs = geo.frames.map(() => startY);
  }

  /** Returns Y for the next block top in the current column. */
  blockTop(): number {
    return this.colYs[this.col] ?? this.pageTop;
  }

  /** Advance after placing a block of height `h`. Inserts synthetic page breaks when needed. */
  advance(
    h: number,
    onPageBreak: () => EngineLayoutLine,
  ): { column: number; deltaTop: number } {
    let top = this.blockTop();
    const pageBottom = this.pageTop + this.bodyHeight;
    if (top + h > pageBottom + 0.5 && h <= this.bodyHeight + 0.5) {
      if (this.col < this.geo.frames.length - 1) {
        this.col += 1;
        top = this.pageTop;
        this.colYs[this.col] = top;
      } else {
        onPageBreak();
        this.pageTop += this.stride;
        this.col = 0;
        this.colYs = this.geo.frames.map(() => this.pageTop);
        top = this.pageTop;
      }
    }
    const deltaTop = top - (this.colYs[this.col] ?? top);
    this.colYs[this.col] = top + h;
    return { column: this.col, deltaTop };
  }

  syncNaturalY(): number {
    return Math.max(...this.colYs, this.pageTop);
  }
}
