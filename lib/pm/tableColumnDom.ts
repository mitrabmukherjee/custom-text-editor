import type { Node as PMNode } from "prosemirror-model";
import { TableMap, updateColumnsOnResize } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";

export const DEFAULT_CELL_MIN_WIDTH = 24;

/** CSS px per screen px when the page is zoomed with `transform: scale`. */
export function editorScale(view: EditorView): number {
  const paper = view.dom.closest(".paper") as HTMLElement | null;
  if (!paper?.offsetWidth) return 1;
  return paper.getBoundingClientRect().width / paper.offsetWidth;
}

export function tableElementAt(view: EditorView, tablePos: number): HTMLTableElement | null {
  const raw = view.nodeDOM(tablePos) as HTMLElement | null;
  if (!raw) return null;
  if (raw.classList.contains("tableWrapper")) return raw.querySelector("table");
  if (raw.tagName === "TABLE") return raw as HTMLTableElement;
  return raw.querySelector("table");
}

/** Drop stray `<col>` nodes and extra `<colgroup>` columns after a resize drag. */
export function trimTableColumnDom(tableEl: HTMLTableElement, columnCount: number) {
  tableEl.querySelectorAll("tbody col, tbody > colgroup").forEach((el) => el.remove());
  const colgroup = tableEl.querySelector("colgroup");
  if (!colgroup) return;
  if (colgroup !== tableEl.firstElementChild) {
    tableEl.insertBefore(colgroup, tableEl.firstElementChild);
  }
  while (colgroup.children.length > columnCount) colgroup.lastElementChild?.remove();
}

/** Rebuild `<colgroup>` from the document table node (Word column widths). */
export function syncTableColumnDom(
  view: EditorView,
  tablePos: number,
  defaultCellMinWidth = DEFAULT_CELL_MIN_WIDTH,
) {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table || table.type.spec.tableRole !== "table") return;
  const tableEl = tableElementAt(view, tablePos);
  if (!tableEl) return;
  const width = TableMap.get(table).width;
  trimTableColumnDom(tableEl, width);
  const colgroup = tableEl.querySelector("colgroup");
  if (!colgroup) return;
  updateColumnsOnResize(table, colgroup, tableEl, defaultCellMinWidth);
}

export function tablePosFromCell(view: EditorView, cellPos: number): number | null {
  if (cellPos < 0) return null;
  const cell = view.state.doc.nodeAt(cellPos);
  if (!cell || cell.type.spec.tableRole !== "cell") return null;
  return view.state.doc.resolve(cellPos).before(-1);
}

export function columnCount(table: PMNode) {
  return TableMap.get(table).width;
}
