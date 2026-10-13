import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import { TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import {
  TableView,
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  columnResizing,
  columnResizingPluginKey,
  deleteColumn,
  deleteRow,
  deleteTable,
  findTable,
  goToNextCell,
  isInTable,
  mergeCells,
  splitCell,
  tableEditing,
} from "prosemirror-tables";
import { schema } from "./schema";
import {
  DEFAULT_CELL_MIN_WIDTH,
  columnCount,
  syncTableColumnDom,
  tableElementAt,
  tablePosFromCell,
  trimTableColumnDom,
} from "./tableColumnDom";

/** Word's Table Layout commands, in menu order. */
export const TABLE_COMMANDS: { label: string; command: Command }[] = [
  { label: "Insert row above", command: addRowBefore },
  { label: "Insert row below", command: addRowAfter },
  { label: "Insert column left", command: addColumnBefore },
  { label: "Insert column right", command: addColumnAfter },
  { label: "Delete row", command: deleteRow },
  { label: "Delete column", command: deleteColumn },
  { label: "Delete table", command: deleteTable },
  { label: "Merge cells", command: mergeCells },
  { label: "Split cell", command: splitCell },
];

export {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  splitCell,
};

/** Prosemirror-tables' view, plus the table's own attributes on the `<table>`. */
class WordTableView extends TableView {
  constructor(node: PMNode, cellMinWidth: number) {
    super(node, cellMinWidth);
    this.applyAttrs(node);
  }

  update(node: PMNode) {
    trimTableColumnDom(this.table, columnCount(node));
    if (!super.update(node)) return false;
    trimTableColumnDom(this.table, columnCount(node));
    this.applyAttrs(node);
    return true;
  }

  private applyAttrs(node: PMNode) {
    this.table.classList.add("pm-table");
    this.table.toggleAttribute("data-borderless", !node.attrs.borders);
  }
}

/**
 * While a column is dragged, prosemirror-tables can leave spare `<col>` nodes in `<tbody>`; trim
 * them every frame so gridlines do not trail past the table when the column is pulled back.
 */
function tableColumnDomCleanupPlugin() {
  let raf = 0;
  return new Plugin({
    view(view) {
      const tick = () => {
        raf = 0;
        const pluginState = columnResizingPluginKey.getState(view.state);
        if (!pluginState?.dragging) return;
        const cell = pluginState.activeHandle;
        const tablePos = tablePosFromCell(view, cell);
        if (tablePos == null) return;
        const table = view.state.doc.nodeAt(tablePos);
        const tableEl = tableElementAt(view, tablePos);
        if (table && tableEl) trimTableColumnDom(tableEl, columnCount(table));
        raf = requestAnimationFrame(tick);
      };
      return {
        update(view, prevState) {
          const prev = columnResizingPluginKey.getState(prevState);
          const next = columnResizingPluginKey.getState(view.state);
          if (next?.dragging && !raf) raf = requestAnimationFrame(tick);
          // Resync after the built-in plugin commits colwidth on mouseup.
          if (prev?.dragging && !next?.dragging) {
            const tablePos = tablePosFromCell(view, next?.activeHandle ?? prev.activeHandle);
            if (tablePos != null) {
              requestAnimationFrame(() => syncTableColumnDom(view, tablePos));
              requestAnimationFrame(() => syncTableColumnDom(view, tablePos));
            }
          }
        },
        destroy() {
          if (raf) cancelAnimationFrame(raf);
        },
      };
    },
  });
}

export function tablePlugins() {
  return [
    columnResizing({ cellMinWidth: DEFAULT_CELL_MIN_WIDTH, View: WordTableView }),
    tableColumnDomCleanupPlugin(),
    tableEditing(),
  ];
}

/** Word's Insert Table: `rows` × `cols` empty cells, equal columns across `width` px. */
export function insertTable(rows: number, cols: number, width: number): Command {
  return (state, dispatch) => {
    if (rows < 1 || cols < 1) return false;
    const { table, table_row, table_cell, paragraph } = schema.nodes;
    const colwidth = [Math.max(24, Math.floor(width / cols))];
    const row = () =>
      table_row.create(
        null,
        Array.from({ length: cols }, () => table_cell.create({ colwidth }, paragraph.create())),
      );
    const node = table.create(null, Array.from({ length: rows }, row));
    if (dispatch) {
      const { $from } = state.selection;
      const at = $from.depth ? $from.after(1) : $from.pos;
      const tr = state.tr.insert(at, node);
      // Word always keeps a paragraph after a table.
      if (at + node.nodeSize === tr.doc.content.size) tr.insert(tr.doc.content.size, paragraph.create());
      // table > row > cell > paragraph
      tr.setSelection(TextSelection.create(tr.doc, at + 4));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Tab: next cell, adding a row after the last cell as Word does. */
export const tabInTable: Command = (state, dispatch) => {
  if (!isInTable(state)) return false;
  if (goToNextCell(1)(state, dispatch)) return true;
  if (!dispatch) return true;
  const added: { tr?: Transaction } = {};
  addRowAfter(state, (tr) => (added.tr = tr));
  const tr = added.tr;
  if (!tr) return false;
  goToNextCell(1)(state.apply(tr), ({ selection }) =>
    tr.setSelection(TextSelection.create(tr.doc, selection.anchor, selection.head)),
  );
  dispatch(tr.scrollIntoView());
  return true;
};

export const shiftTabInTable: Command = (state, dispatch) =>
  isInTable(state) ? goToNextCell(-1)(state, dispatch) || true : false;

/** Show or hide the gridlines of the table holding the selection. */
export function setTableBorders(borders: boolean): Command {
  return (state, dispatch) => {
    const found = findTable(state.selection.$from);
    if (!found) return false;
    dispatch?.(state.tr.setNodeMarkup(found.pos, undefined, { ...found.node.attrs, borders }));
    return true;
  };
}

export function tableBorders(state: EditorState): boolean | null {
  const found = findTable(state.selection.$from);
  return found ? !!found.node.attrs.borders : null;
}
