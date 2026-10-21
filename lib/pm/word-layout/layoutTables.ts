import type { Node as PMNode } from "prosemirror-model";
import type { CellMarginPx, ParagraphFormat } from "../format";
import { layoutParagraph } from "./paragraphLayout";
import type { EngineLayoutLine } from "./types";

function cellContentWidth(cell: PMNode, tableWidth: number, row: PMNode): number {
  const colwidth = cell.attrs.colwidth as number[] | null | undefined;
  if (colwidth?.[0]) return colwidth[0];
  return tableWidth / Math.max(1, row.childCount);
}

function cellVerticalMargin(margin: CellMarginPx | null | undefined): number {
  if (!margin) return 0;
  return (margin.top ?? 0) + (margin.bottom ?? 0);
}

export function layoutTableRows(
  table: PMNode,
  tablePos: number,
  section: number,
  startBlock: number,
  naturalTop: number,
  tableWidth: number,
  gridRowPx?: number,
): { lines: EngineLayoutLine[]; bottom: number; height: number } {
  const lines: EngineLayoutLine[] = [];
  let y = naturalTop;
  let block = startBlock;
  table.forEach((row, rowOffset, rowIndex) => {
    const rowPos = tablePos + 1 + rowOffset;
    const cantSplit = !!row.attrs.cantSplit;
    const rowBlock = block++;
    let rowH = gridRowPx ?? 0;
    let keepNext = false;
    const minRowPx = (row.attrs.rowHeight as number | null) ?? 0;
    row.forEach((cell, cellOffset) => {
      let cellH = cellVerticalMargin(cell.attrs.cellMargin as CellMarginPx | null);
      const cellPos = rowPos + 1 + cellOffset;
      let innerY = 0;
      cell.forEach((blockNode, blockOffset) => {
        if (!blockNode.isTextblock) return;
        const blockPos = cellPos + 1 + blockOffset;
        const f = blockNode.attrs.format as ParagraphFormat | null;
        if (f?.keepNext) keepNext = true;
        const cellW = cellContentWidth(cell, tableWidth, row);
        const { height } = layoutParagraph(
          blockNode,
          blockPos,
          cellW,
          innerY,
          block,
          section,
          { defaultTabStop: 48 },
          { keepNext, numbered: false },
        );
        innerY += height;
      });
      cellH += innerY;
      rowH = Math.max(rowH, cellH);
    });

    if (minRowPx) rowH = Math.max(rowH, minRowPx);
    if (!rowH) rowH = gridRowPx ?? 20;

    if (gridRowPx) {
      const steps = Math.max(1, Math.ceil(rowH / gridRowPx));
      for (let s = 0; s < steps; s++) {
        const top = y + s * gridRowPx;
        lines.push({
          top,
          bottom: top + gridRowPx,
          block: rowBlock,
          section,
          keepNext: s === 0 && keepNext,
          keepLines: cantSplit || undefined,
          numbered: false,
          rowHeight: gridRowPx,
          anchor: { pos: rowIndex === 0 ? tablePos : rowPos, textOffset: 0 },
        });
      }
      if (steps > 1) block = rowBlock + 1;
    } else {
      lines.push({
        top: y,
        bottom: y + rowH,
        block: rowBlock,
        section,
        keepNext,
        keepLines: cantSplit || undefined,
        numbered: false,
        anchor: { pos: rowIndex === 0 ? tablePos : rowPos, textOffset: 0 },
      });
    }
    y += rowH;
  });
  return { lines, bottom: y, height: y - naturalTop };
}
