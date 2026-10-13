"use client";

import { useEffect, useRef, useState } from "react";
import {
  applyCharacterFormat,
  clearFormatting,
  growFont,
  toggleCaps,
  toggleStrike,
  toggleSubscript,
  toggleSuperscript,
  toggleUnderline,
  type CharacterPatch,
} from "@/lib/pm/character";
import type { EditorController } from "@/lib/pm/controller";
import type { TextAlign } from "@/lib/pm/format";
import {
  BULLET_STYLES,
  MULTILEVEL_PRESETS,
  NUMBER_FORMATS,
  applyMultilevel,
  continueNumbering,
  listAttrsAt,
  restartNumbering,
  setBullet,
  setNumbering,
  setNumberingValue,
} from "@/lib/pm/lists";
import { activeList, hangingIndent, indent, toggleList } from "@/lib/pm/paragraph";
import type { MarkupFilter } from "@/lib/pm/pagination";
import type { PasteMode } from "@/lib/pm/plugins/paste";
import { schema } from "@/lib/pm/schema";
import { fitImage, readImageFile } from "@/lib/pm/objects";
import { TABLE_COMMANDS } from "@/lib/pm/tables";
import { InsertTableDialog } from "./InsertDialogs";
import { CharacterToolbar } from "./CharacterToolbar";
import { FindReplaceDialog, type FindMode } from "./FindReplaceDialog";
import { FontDialog, type FontDialogValue } from "./FontDialog";
import { NumberingValueDialog } from "./NumberingValueDialog";
import { StylesGallery } from "./StylesGallery";
import { ParagraphDialog } from "./ParagraphDialog";
import { PageSetupDialog } from "./PageSetupDialog";
import { SectionColumnsDialog } from "./SectionColumnsDialog";
import { insertSectionBreak, sectionIndexAt, sectionsOf } from "@/lib/pm/sections";
import { Bars, MenuList, Separator, ToolButton, iconBtn, type MenuItem } from "./toolbarParts";
import { insertQaLine, insertSpeakerLine, insertTimestamp } from "@/lib/pm/transcript";

type StudioActions = {
  onPlayPause?: () => void;
  onPrint?: () => void;
};

/** View options owned by the editor shell (navigation pane, ruler, zoom). */
export type ViewControls = {
  zoom: number;
  setZoom: (zoom: number) => void;
  navigation: boolean;
  setNavigation: (open: boolean) => void;
  ruler: boolean;
  setRuler: (show: boolean) => void;
};

/** Review options owned by the editor shell (Show Markup filter, comment dialog). */
export type ReviewControls = {
  markup: MarkupFilter;
  setMarkup: (markup: MarkupFilter) => void;
  onNewComment: () => void;
  /** Open the hyperlink dialog (shared with the document's context menu). */
  onLink: () => void;
};

type Props = {
  controller: EditorController | null;
  readOnly?: boolean;
  actions?: StudioActions;
  simpleMarkup: boolean;
  onSimpleMarkupChange: (simple: boolean) => void;
  view: ViewControls;
  review: ReviewControls;
};

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2];

const MENUS = [
  "File",
  "Edit",
  "Format",
  "Insert",
  "Layout",
  "Table",
  "Review",
  "Audio",
  "View",
  "Help",
] as const;

const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];

const PASTE_MODES: { mode: PasteMode; label: string }[] = [
  { mode: "auto", label: "Paste: automatic" },
  { mode: "keep", label: "Paste: keep source formatting" },
  { mode: "merge", label: "Paste: merge formatting" },
  { mode: "text", label: "Paste: keep text only" },
];

function fontDialogValue(
  format: NonNullable<ReturnType<EditorController["getCharacterFormat"]>>,
): FontDialogValue {
  return {
    family: format.family,
    size: format.size,
    bold: format.bold,
    italic: format.italic,
    underline: format.underline ?? undefined,
    strike: format.strike ?? undefined,
    color: format.color ?? undefined,
    caps: format.caps ?? undefined,
    vertAlign: format.vertAlign ?? undefined,
  };
}

/** Space before/after toggled by Word's line-spacing menu: 12 pt before, 8 pt after. */
const SPACE_BEFORE_PX = 16;
const SPACE_AFTER_PX = 32 / 3;

const ALIGN_ICONS: Record<TextAlign, [number, number][]> = {
  left: [[1, 14], [1, 9], [1, 14], [1, 9], [1, 12]],
  center: [[1, 14], [3.5, 9], [1, 14], [3.5, 9], [2, 12]],
  right: [[1, 14], [6, 9], [1, 14], [6, 9], [3, 12]],
  justify: [[1, 14], [1, 14], [1, 14], [1, 14], [1, 9]],
};

const ALIGNMENTS: { value: TextAlign; label: string; shortcut: string }[] = [
  { value: "left", label: "Align left", shortcut: "Ctrl+L" },
  { value: "center", label: "Center", shortcut: "Ctrl+E" },
  { value: "right", label: "Align right", shortcut: "Ctrl+R" },
  { value: "justify", label: "Justify", shortcut: "Ctrl+J" },
];

function ListIcon({ numbered }: { numbered: boolean }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden>
      {[3, 8, 13].map((y, i) =>
        numbered ? (
          <text key={y} x="0" y={y + 2} fontSize="5" fill="currentColor">
            {i + 1}
          </text>
        ) : (
          <circle key={y} cx="2" cy={y} r="1.3" fill="currentColor" />
        ),
      )}
      {[3, 8, 13].map((y) => (
        <rect key={y} x="5" y={y - 0.75} width="10" height="1.5" fill="currentColor" />
      ))}
    </svg>
  );
}

function IndentIcon({ increase }: { increase: boolean }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden>
      {[2, 14].map((y) => (
        <rect key={y} x="1" y={y - 0.75} width="14" height="1.5" fill="currentColor" />
      ))}
      {[6, 10].map((y) => (
        <rect key={y} x="7" y={y - 0.75} width="8" height="1.5" fill="currentColor" />
      ))}
      <path d={increase ? "M1 5.5 L4.5 8 L1 10.5 Z" : "M4.5 5.5 L1 8 L4.5 10.5 Z"} fill="currentColor" />
    </svg>
  );
}

export function StudioEditorChrome({
  controller,
  readOnly,
  actions,
  simpleMarkup,
  onSimpleMarkupChange,
  view,
  review,
}: Props) {
  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const [paragraphDialog, setParagraphDialog] = useState(false);
  const [fontDialog, setFontDialog] = useState(false);
  const [numberingDialog, setNumberingDialog] = useState(false);
  const [find, setFind] = useState<FindMode | null>(null);
  const [tableDialog, setTableDialog] = useState(false);
  const [pageSetupDialog, setPageSetupDialog] = useState(false);
  const [columnsDialog, setColumnsDialog] = useState(false);
  const wrapRef = useRef<HTMLDivElement>(null);
  const reviewRef = useRef(review);
  useEffect(() => {
    reviewRef.current = review;
  });

  const insertPicture = async (file: File) => {
    if (!controller) return;
    const image = await readImageFile(file);
    const setup = controller.getPageSetup();
    const max = setup.width - setup.marginLeft - setup.marginRight;
    controller.insertImage({ ...image, ...fitImage(image.width, image.height, max) });
  };
  const pickPicture = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/bmp";
    input.onchange = () => {
      const file = input.files?.[0];
      if (file) void insertPicture(file);
    };
    input.click();
  };

  useEffect(() => {
    const onDoc = (e: MouseEvent) => {
      if (!wrapRef.current?.contains(e.target as Node)) setOpenMenu(null);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  useEffect(() => {
    const dom = controller?.view.dom;
    if (!dom) return;
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey) return;
      const key = e.key.toLowerCase();
      if (e.altKey) {
        if (e.code === "KeyM" && !readOnly && controller?.canAddComment()) {
          e.preventDefault();
          reviewRef.current.onNewComment();
        }
        return;
      }
      if (key === "k" && !readOnly) {
        e.preventDefault();
        reviewRef.current.onLink();
      } else if (key === "f" || (key === "h" && !readOnly)) {
        e.preventDefault();
        setFind(key === "f" ? "find" : "replace");
      } else if (key === "d" && !readOnly) {
        e.preventDefault();
        setFontDialog(true);
      }
    };
    dom.addEventListener("keydown", onKey);
    return () => dom.removeEventListener("keydown", onKey);
  }, [controller, readOnly]);

  const print = () => {
    document.documentElement.classList.add("printing");
    const done = () => {
      document.documentElement.classList.remove("printing");
      window.removeEventListener("afterprint", done);
    };
    window.addEventListener("afterprint", done);
    if (actions?.onPrint) actions.onPrint();
    else window.print();
  };
  const canUndo = !readOnly && !!controller?.canUndo();
  const canRedo = !readOnly && !!controller?.canRedo();
  const tracking = !!controller?.isTracking();
  const changes = controller?.listChanges() ?? [];
  const character = controller?.getCharacterFormat();
  const pasteMode = controller?.getPasteMode() ?? "auto";
  const paragraph = controller?.getParagraphFormat() ?? {};
  const align = paragraph.align ?? "left";
  const list = controller ? activeList(controller.state) : null;
  const { bullet_list, ordered_list } = schema.nodes;

  const setAlign = (value: TextAlign) =>
    controller?.setParagraphFormat({ align: value === "left" ? undefined : value });
  const setLineSpacing = (value: number) =>
    controller?.setParagraphFormat({ lineRule: "auto", lineValue: value });
  const toggleSpaceBefore = () =>
    controller?.setParagraphFormat((f) => ({ spaceBefore: f.spaceBefore ? undefined : SPACE_BEFORE_PX }));
  const toggleSpaceAfter = () =>
    controller?.setParagraphFormat((f) => ({ spaceAfter: f.spaceAfter ? undefined : SPACE_AFTER_PX }));

  const lineSpacingItems: MenuItem[] = [
    ...LINE_SPACINGS.map((value) => ({
      label: value.toFixed(value === 1.15 ? 2 : 1),
      checked: paragraph.lineRule === "auto" && paragraph.lineValue === value,
      disabled: readOnly,
      run: () => setLineSpacing(value),
    })),
    { label: "Line spacing options\u2026", disabled: readOnly, run: () => setParagraphDialog(true) },
    {
      label: paragraph.spaceBefore ? "Remove space before paragraph" : "Add space before paragraph",
      disabled: readOnly,
      run: toggleSpaceBefore,
      shortcut: "Ctrl+0",
    },
    {
      label: paragraph.spaceAfter ? "Remove space after paragraph" : "Add space after paragraph",
      disabled: readOnly,
      run: toggleSpaceAfter,
    },
  ];

  const listInfo = controller ? listAttrsAt(controller.state) : null;
  const inOrdered = listInfo?.type === ordered_list;
  const bulletItems: MenuItem[] = BULLET_STYLES.map(({ bullet, label }) => ({
    label,
    checked: listInfo?.type === bullet_list && listInfo.attrs.bullet === bullet,
    run: () => controller?.exec(setBullet(bullet)),
  }));
  const numberingItems: MenuItem[] = [
    ...NUMBER_FORMATS.map(({ format, label }) => ({
      label,
      checked: inOrdered && listInfo?.attrs.numFormat === format,
      run: () => controller?.exec(setNumbering(format)),
    })),
    ...MULTILEVEL_PRESETS.map((preset) => ({
      label: `Multilevel: ${preset.label}`,
      run: () => controller?.exec(applyMultilevel(preset)),
    })),
    { label: "Restart at 1", disabled: !inOrdered, run: () => controller?.exec(restartNumbering) },
    { label: "Continue numbering", disabled: !inOrdered, run: () => controller?.exec(continueNumbering) },
    { label: "Set numbering value\u2026", disabled: !inOrdered, run: () => setNumberingDialog(true) },
  ];

  const inTable = !!controller?.isInTable();
  const borders = controller?.tableBorders() ?? null;
  const locked = !!controller?.isTrackingLocked();
  const comments = controller?.getComments() ?? [];
  const { markup, setMarkup } = review;
  const toggleMarkup = (key: keyof MarkupFilter) => setMarkup({ ...markup, [key]: !markup[key] });

  const menuItems: Record<string, MenuItem[]> = {
    File: [
      { label: "Page Setup\u2026", disabled: readOnly, run: () => setPageSetupDialog(true) },
      { label: "Print\u2026", run: print, shortcut: "Ctrl+P" },
      { label: "Save as PDF\u2026", run: print },
    ],
    Edit: [
      { label: "Undo", disabled: !canUndo, run: () => controller?.undo(), shortcut: "Ctrl+Z" },
      { label: "Redo", disabled: !canRedo, run: () => controller?.redo(), shortcut: "Ctrl+Y" },
      { label: "Find\u2026", run: () => setFind("find"), shortcut: "Ctrl+F" },
      { label: "Replace\u2026", disabled: readOnly, run: () => setFind("replace"), shortcut: "Ctrl+H" },
      ...PASTE_MODES.map(({ mode, label }) => ({
        label,
        checked: pasteMode === mode,
        run: () => controller?.setPasteMode(mode),
      })),
    ],
    Format: [
      { label: "Font\u2026", disabled: readOnly, run: () => setFontDialog(true), shortcut: "Ctrl+D" },
      {
        label: "Bold",
        checked: !!character?.bold,
        disabled: readOnly,
        run: () => controller?.toggleMark(schema.marks.strong),
        shortcut: "Ctrl+B",
      },
      {
        label: "Italic",
        checked: !!character?.italic,
        disabled: readOnly,
        run: () => controller?.toggleMark(schema.marks.em),
        shortcut: "Ctrl+I",
      },
      {
        label: "Underline",
        checked: !!character?.underline,
        disabled: readOnly,
        run: () => controller?.exec(toggleUnderline()),
        shortcut: "Ctrl+U",
      },
      {
        label: "Strikethrough",
        checked: !!character?.strike,
        disabled: readOnly,
        run: () => controller?.exec(toggleStrike()),
      },
      {
        label: "Subscript",
        checked: character?.vertAlign === "subscript",
        disabled: readOnly,
        run: () => controller?.exec(toggleSubscript),
        shortcut: "Ctrl+=",
      },
      {
        label: "Superscript",
        checked: character?.vertAlign === "superscript",
        disabled: readOnly,
        run: () => controller?.exec(toggleSuperscript),
        shortcut: "Ctrl+Shift+=",
      },
      {
        label: "Small caps",
        checked: character?.caps === "small",
        disabled: readOnly,
        run: () => controller?.exec(toggleCaps("small")),
        shortcut: "Ctrl+Shift+K",
      },
      {
        label: "All caps",
        checked: character?.caps === "all",
        disabled: readOnly,
        run: () => controller?.exec(toggleCaps("all")),
        shortcut: "Ctrl+Shift+A",
      },
      {
        label: "Grow font",
        disabled: readOnly,
        run: () => controller?.exec(growFont(1)),
        shortcut: "Ctrl+Shift+>",
      },
      {
        label: "Shrink font",
        disabled: readOnly,
        run: () => controller?.exec(growFont(-1)),
        shortcut: "Ctrl+Shift+<",
      },
      {
        label: "Clear formatting",
        disabled: readOnly,
        run: () => controller?.exec(clearFormatting),
        shortcut: "Ctrl+Space",
      },
      { label: "Paragraph\u2026", disabled: readOnly, run: () => setParagraphDialog(true) },
      ...ALIGNMENTS.map(({ value, label, shortcut }) => ({
        label,
        checked: align === value,
        disabled: readOnly,
        run: () => setAlign(value),
        shortcut,
      })),
      {
        label: "Increase indent",
        disabled: readOnly,
        run: () => controller?.exec(indent(1)),
        shortcut: "Ctrl+M",
      },
      {
        label: "Decrease indent",
        disabled: readOnly,
        run: () => controller?.exec(indent(-1)),
        shortcut: "Ctrl+Shift+M",
      },
      { label: "Hanging indent", disabled: readOnly, run: () => controller?.exec(hangingIndent(1)) },
      {
        label: "Remove hanging indent",
        disabled: readOnly || !((paragraph.indentFirstLine ?? 0) < 0),
        run: () => controller?.exec(hangingIndent(-1)),
      },
      {
        label: "Bullets",
        checked: list === bullet_list,
        disabled: readOnly,
        run: () => controller?.exec(toggleList(bullet_list)),
        shortcut: "Ctrl+Shift+L",
      },
      {
        label: "Numbering",
        checked: list === ordered_list,
        disabled: readOnly,
        run: () => controller?.exec(toggleList(ordered_list)),
      },
    ],
    Insert: [
      {
        label: "Page break",
        disabled: readOnly,
        run: () => controller?.insertPageBreak(),
        shortcut: "Ctrl+Enter",
      },
      {
        label: "Section break (next page)",
        disabled: readOnly,
        run: () => controller?.exec(insertSectionBreak("nextPage")),
      },
      {
        label: "Section break (continuous)",
        disabled: readOnly,
        run: () => controller?.exec(insertSectionBreak("continuous")),
      },
      {
        label: "Section break (even page)",
        disabled: readOnly,
        run: () => controller?.exec(insertSectionBreak("evenPage")),
      },
      {
        label: "Section break (odd page)",
        disabled: readOnly,
        run: () => controller?.exec(insertSectionBreak("oddPage")),
      },
      { label: "Q line", disabled: readOnly, run: () => controller?.exec(insertQaLine("q")) },
      { label: "A line", disabled: readOnly, run: () => controller?.exec(insertQaLine("a")) },
      { label: "Speaker line", disabled: readOnly, run: () => controller?.exec(insertSpeakerLine()) },
      { label: "Timestamp", disabled: readOnly, run: () => controller?.exec(insertTimestamp()) },
      { label: "Table\u2026", disabled: readOnly, run: () => setTableDialog(true) },
      { label: "Picture\u2026", disabled: readOnly, run: () => pickPicture() },
      { label: "Link\u2026", disabled: readOnly, run: review.onLink, shortcut: "Ctrl+K" },
      { label: "Comment", disabled: readOnly || !controller?.canAddComment(), run: review.onNewComment },
    ],
    Layout: [
      { label: "Page Setup\u2026", disabled: readOnly, run: () => setPageSetupDialog(true) },
      {
        label: "Columns\u2026",
        disabled: readOnly,
        run: () => setColumnsDialog(true),
      },
    ],
    Table: [
      { label: "Insert table\u2026", disabled: readOnly, run: () => setTableDialog(true) },
      ...TABLE_COMMANDS.map(({ label, command }) => ({
        label,
        disabled: readOnly || !inTable || !controller?.can(command),
        run: () => controller?.execUntracked(command),
      })),
      {
        label: "View gridlines",
        checked: borders === true,
        disabled: readOnly || !inTable,
        run: () => controller?.setTableBorders(!borders),
      },
    ],
    Review: [
      {
        label: "Track changes",
        checked: tracking,
        disabled: readOnly || locked,
        run: () => controller?.setTrackChanges(!tracking),
        shortcut: "Ctrl+Shift+E",
      },
      {
        label: "Lock tracking",
        checked: locked,
        disabled: readOnly,
        run: () => controller?.setTrackingLocked(!locked),
      },
      {
        label: "Previous change",
        disabled: !changes.length,
        run: () => controller?.goToAdjacentChange(-1),
      },
      { label: "Next change", disabled: !changes.length, run: () => controller?.goToAdjacentChange(1) },
      {
        label: "Accept and move to next",
        disabled: readOnly || !changes.length,
        run: () => controller?.resolveSelection("accept"),
      },
      {
        label: "Reject and move to next",
        disabled: readOnly || !changes.length,
        run: () => controller?.resolveSelection("reject"),
      },
      {
        label: "Accept all changes",
        disabled: readOnly || !changes.length,
        run: () => controller?.acceptAll(),
      },
      {
        label: "Reject all changes",
        disabled: readOnly || !changes.length,
        run: () => controller?.rejectAll(),
      },
      {
        label: "New comment",
        disabled: readOnly || !controller?.canAddComment(),
        run: review.onNewComment,
        shortcut: "Ctrl+Alt+M",
      },
      {
        label: "Previous comment",
        disabled: !comments.length,
        run: () => controller?.goToAdjacentComment(-1),
      },
      {
        label: "Next comment",
        disabled: !comments.length,
        run: () => controller?.goToAdjacentComment(1),
      },
      {
        label: "Delete all comments",
        disabled: readOnly || !comments.length,
        run: () => comments.forEach((c) => controller?.deleteComment(c.id)),
      },
      {
        label: "Simple markup",
        checked: simpleMarkup,
        run: () => onSimpleMarkupChange(!simpleMarkup),
      },
      {
        label: "Show: insertions and deletions",
        checked: markup.insertions,
        run: () => toggleMarkup("insertions"),
      },
      { label: "Show: formatting", checked: markup.formatting, run: () => toggleMarkup("formatting") },
      { label: "Show: comments", checked: markup.comments, run: () => toggleMarkup("comments") },
    ],
    Audio: [{ label: "Play / pause", run: () => actions?.onPlayPause?.() }],
    View: [
      {
        label: "Navigation pane",
        checked: view.navigation,
        run: () => view.setNavigation(!view.navigation),
      },
      { label: "Ruler", checked: view.ruler, run: () => view.setRuler(!view.ruler) },
      ...ZOOM_LEVELS.map((zoom) => ({
        label: `Zoom ${Math.round(zoom * 100)}%`,
        checked: Math.abs(view.zoom - zoom) < 0.001,
        run: () => view.setZoom(zoom),
      })),
    ],
    Help: [{ label: "ProseMirror transcript editor", run: () => {} }],
  };

  return (
    <div
      ref={wrapRef}
      className="studio-editor-chrome relative z-30 shrink-0 border-b border-zinc-800 bg-zinc-950"
    >
      <div className="flex flex-wrap items-center gap-0.5 px-2 py-1 text-[13px] text-zinc-400">
        {MENUS.map((name) => {
          const items = menuItems[name] ?? [];
          return (
            <div key={name} className="relative">
              <button
                type="button"
                className={`rounded px-2 py-0.5 hover:bg-zinc-800 hover:text-zinc-200 ${
                  openMenu === name ? "bg-zinc-800 text-zinc-100" : ""
                }`}
                onClick={() => setOpenMenu((m) => (m === name ? null : name))}
                disabled={readOnly && name !== "View" && name !== "Help" && name !== "Review"}
              >
                {name}
              </button>
              {openMenu === name && items.length > 0 && (
                <MenuList items={items} onDone={() => setOpenMenu(null)} />
              )}
            </div>
          );
        })}
        <span className="ml-auto flex items-center gap-2">
          {tracking && (
            <span className="rounded border border-amber-800/60 bg-amber-950/50 px-2 py-0.5 text-[11px] font-medium text-amber-300">
              Tracking changes
            </span>
          )}
          {readOnly && (
            <span className="rounded border border-rose-900/60 bg-rose-950/50 px-2 py-0.5 text-[11px] font-medium text-rose-300">
              Read-Only
            </span>
          )}
        </span>
      </div>

      <div
        className="flex flex-wrap items-center gap-0.5 border-t border-zinc-800/80 px-2 py-1"
        role="toolbar"
        aria-label="Formatting"
      >
        <button
          type="button"
          className={iconBtn(false, !canUndo)}
          title="Undo"
          disabled={!canUndo}
          onClick={() => controller?.undo()}
        >
          {"\u21b6"}
        </button>
        <button
          type="button"
          className={iconBtn(false, !canRedo)}
          title="Redo"
          disabled={!canRedo}
          onClick={() => controller?.redo()}
        >
          {"\u21b7"}
        </button>
        <span className="mx-1 h-5 w-px bg-zinc-800" aria-hidden />
        <button type="button" className={iconBtn()} title="Print" onClick={print}>
          {"\u2399"}
        </button>
        <button
          type="button"
          className={iconBtn()}
          title="Play / pause audio"
          onClick={() => actions?.onPlayPause?.()}
        >
          {"\u25b6"}
        </button>
        <Separator />
        {controller && (
          <CharacterToolbar
            controller={controller}
            readOnly={readOnly}
            openMenu={openMenu}
            setOpenMenu={setOpenMenu}
          />
        )}
        <Separator />
        {ALIGNMENTS.map(({ value, label, shortcut }) => (
          <ToolButton
            key={value}
            title={`${label} (${shortcut})`}
            active={align === value}
            disabled={readOnly}
            onClick={() => setAlign(value)}
          >
            <Bars rows={ALIGN_ICONS[value]} />
          </ToolButton>
        ))}
        <div className="relative">
          <ToolButton
            title="Line and paragraph spacing"
            active={openMenu === "lineSpacing"}
            disabled={readOnly}
            onClick={() => setOpenMenu((m) => (m === "lineSpacing" ? null : "lineSpacing"))}
          >
            <span className="flex items-center gap-0.5 text-xs">
              <span aria-hidden>{"\u2195"}</span>
              <Bars rows={[[3, 12], [3, 12], [3, 12], [3, 12], [3, 12]]} />
            </span>
          </ToolButton>
          {openMenu === "lineSpacing" && (
            <MenuList items={lineSpacingItems} onDone={() => setOpenMenu(null)} />
          )}
        </div>
        <div className="relative flex items-center">
          <ToolButton
            title="Bullets (Ctrl+Shift+L)"
            active={list === bullet_list}
            disabled={readOnly}
            onClick={() => controller?.exec(toggleList(bullet_list))}
          >
            <ListIcon numbered={false} />
          </ToolButton>
          <ToolButton
            title="Bullet library"
            disabled={readOnly}
            className="!px-0.5"
            onClick={() => setOpenMenu((m) => (m === "bullets" ? null : "bullets"))}
          >
            <span className="text-[9px]">{"\u25be"}</span>
          </ToolButton>
          {openMenu === "bullets" && (
            <MenuList items={bulletItems} onDone={() => setOpenMenu(null)} />
          )}
        </div>
        <div className="relative flex items-center">
          <ToolButton
            title="Numbering"
            active={list === ordered_list}
            disabled={readOnly}
            onClick={() => controller?.exec(toggleList(ordered_list))}
          >
            <ListIcon numbered />
          </ToolButton>
          <ToolButton
            title="Numbering library"
            disabled={readOnly}
            className="!px-0.5"
            onClick={() => setOpenMenu((m) => (m === "numbering" ? null : "numbering"))}
          >
            <span className="text-[9px]">{"\u25be"}</span>
          </ToolButton>
          {openMenu === "numbering" && (
            <MenuList items={numberingItems} onDone={() => setOpenMenu(null)} />
          )}
        </div>
        <ToolButton
          title="Decrease indent (Ctrl+Shift+M)"
          disabled={readOnly}
          onClick={() => controller?.exec(indent(-1))}
        >
          <IndentIcon increase={false} />
        </ToolButton>
        <ToolButton
          title="Increase indent (Ctrl+M)"
          disabled={readOnly}
          onClick={() => controller?.exec(indent(1))}
        >
          <IndentIcon increase />
        </ToolButton>
        <ToolButton
          title="Paragraph settings"
          disabled={readOnly}
          onClick={() => setParagraphDialog(true)}
        >
          <span className="px-0.5 text-sm leading-4">{"\u00b6"}</span>
        </ToolButton>
        <Separator />
        {controller && (
          <StylesGallery
            controller={controller}
            readOnly={readOnly}
            open={openMenu === "styles"}
            onToggle={() => setOpenMenu((m) => (m === "styles" ? null : "styles"))}
            onClose={() => setOpenMenu(null)}
          />
        )}
        <span className="mx-1 h-5 w-px bg-zinc-800" aria-hidden />
        <button
          type="button"
          className={`${iconBtn(tracking, readOnly || locked)} px-2 text-xs font-medium`}
          title={locked ? "Track changes (locked)" : "Track changes (Ctrl+Shift+E)"}
          aria-pressed={tracking}
          disabled={readOnly || locked}
          onClick={() => controller?.setTrackChanges(!tracking)}
        >
          {locked ? "Track \u{1f512}" : "Track"}
        </button>
        <ToolButton
          title="New comment (Ctrl+Alt+M)"
          disabled={readOnly || !controller?.canAddComment()}
          onClick={review.onNewComment}
        >
          <span className="px-0.5 text-xs">Comment</span>
        </ToolButton>
        <Separator />
        <ToolButton title="Insert table" disabled={readOnly} onClick={() => setTableDialog(true)}>
          <span className="px-0.5 text-xs">Table</span>
        </ToolButton>
        <ToolButton title="Insert picture" disabled={readOnly} onClick={() => pickPicture()}>
          <span className="px-0.5 text-xs">Picture</span>
        </ToolButton>
        <ToolButton
          title="Link (Ctrl+K)"
          active={!!controller?.getLink()}
          disabled={readOnly}
          onClick={review.onLink}
        >
          <span className="px-0.5 text-xs">Link</span>
        </ToolButton>
        <Separator />
        <ToolButton title="Find (Ctrl+F)" active={find === "find"} onClick={() => setFind("find")}>
          <span className="px-0.5 text-xs">Find</span>
        </ToolButton>
        <ToolButton
          title="Replace (Ctrl+H)"
          active={find === "replace"}
          disabled={readOnly}
          onClick={() => setFind("replace")}
        >
          <span className="px-0.5 text-xs">Replace</span>
        </ToolButton>
        {changes.length > 0 && (
          <span className="ml-auto text-[11px] text-zinc-500">
            {changes.length} change{changes.length === 1 ? "" : "s"}
          </span>
        )}
      </div>

      {find && controller && (
        <FindReplaceDialog
          controller={controller}
          mode={find}
          readOnly={readOnly}
          onModeChange={setFind}
          onClose={() => {
            setFind(null);
            controller.focus();
          }}
        />
      )}

      {fontDialog && controller && character && (
        <FontDialog
          format={fontDialogValue(character)}
          onClose={() => {
            setFontDialog(false);
            controller.focus();
          }}
          onApply={(value) => {
            setFontDialog(false);
            const before = fontDialogValue(character);
            const patch: CharacterPatch = {};
            for (const key of Object.keys(value) as (keyof FontDialogValue)[]) {
              if (value[key] !== before[key]) Object.assign(patch, { [key]: value[key] ?? null });
            }
            controller.exec(applyCharacterFormat(patch));
          }}
        />
      )}

      {numberingDialog && controller && (
        <NumberingValueDialog
          value={Number(listInfo?.attrs.order ?? 1) + (listInfo?.index ?? 0)}
          onClose={() => {
            setNumberingDialog(false);
            controller.focus();
          }}
          onApply={(value) => {
            setNumberingDialog(false);
            controller.exec(setNumberingValue(value - (listInfo?.index ?? 0)));
          }}
        />
      )}

      {tableDialog && controller && (
        <InsertTableDialog
          onClose={() => {
            setTableDialog(false);
            controller.focus();
          }}
          onSubmit={(rows, cols) => {
            setTableDialog(false);
            controller.insertTable(rows, cols);
          }}
        />
      )}

      {paragraphDialog && controller && (
        <ParagraphDialog
          format={paragraph}
          onClose={() => {
            setParagraphDialog(false);
            controller.focus();
          }}
          onApply={(patch) => {
            setParagraphDialog(false);
            controller.setParagraphFormat(patch);
          }}
        />
      )}

      {pageSetupDialog && controller && (
        <PageSetupDialog
          setup={controller.getPageSetup()}
          onClose={() => {
            setPageSetupDialog(false);
            controller.focus();
          }}
          onApply={(setup) => {
            setPageSetupDialog(false);
            controller.setPageSetup(setup);
          }}
        />
      )}

      {columnsDialog && controller && (
        <SectionColumnsDialog
          count={
            sectionsOf(controller.state)[
              Math.min(
                sectionIndexAt(controller.state, controller.state.selection.head),
                sectionsOf(controller.state).length - 1,
              )
            ]?.columns?.count ?? 1
          }
          space={
            sectionsOf(controller.state)[
              Math.min(
                sectionIndexAt(controller.state, controller.state.selection.head),
                sectionsOf(controller.state).length - 1,
              )
            ]?.columns?.space ?? 48
          }
          onClose={() => {
            setColumnsDialog(false);
            controller.focus();
          }}
          onApply={(count, space) => {
            setColumnsDialog(false);
            controller.setSectionColumns(count, space);
          }}
        />
      )}
    </div>
  );
}
