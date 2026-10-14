import { BorderStyle, type IBorderOptions, type IBordersOptions, type ITableCellBorders } from "docx";
import { PT_TO_PX, type CellBorderEdges, type ParagraphBorderSide, type ParagraphFormat } from "@/lib/pm/format";

const BORDER_STYLE: Record<ParagraphBorderSide["style"], (typeof BorderStyle)[keyof typeof BorderStyle]> =
  {
    solid: BorderStyle.SINGLE,
    dashed: BorderStyle.DASHED,
    dotted: BorderStyle.DOTTED,
    double: BorderStyle.DOUBLE,
  };

export function exportBorderSide(side: ParagraphBorderSide): IBorderOptions {
  return {
    style: BORDER_STYLE[side.style] ?? BorderStyle.SINGLE,
    size: Math.max(2, Math.round(side.widthPt * 8)),
    color: side.color.replace("#", "").toUpperCase(),
    space: side.paddingPx !== undefined ? Math.round((side.paddingPx / PT_TO_PX) * 8) : undefined,
  };
}

export function exportParagraphBorders(
  box: ParagraphFormat["boxBorder"],
): IBordersOptions | undefined {
  if (!box) return undefined;
  const borders: Record<string, IBorderOptions> = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const edge = box[side];
    if (edge) borders[side] = exportBorderSide(edge);
  }
  return Object.keys(borders).length ? (borders as IBordersOptions) : undefined;
}

export function exportCellBorders(edges: CellBorderEdges | null | undefined): ITableCellBorders | undefined {
  if (!edges) return undefined;
  const borders: Record<string, IBorderOptions> = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const edge = edges[side];
    if (edge === "none") borders[side] = { style: BorderStyle.NONE, size: 0, color: "auto" };
    else if (edge) borders[side] = exportBorderSide(edge);
  }
  return Object.keys(borders).length ? (borders as ITableCellBorders) : undefined;
}
