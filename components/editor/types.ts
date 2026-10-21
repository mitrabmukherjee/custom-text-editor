import type { Ref } from "react";
import type { DocJSON } from "@/lib/pm/json";
import type { PageSetup, PaginationLayout } from "@/lib/pm/pagination";
import type { Author, TrackedChange } from "@/lib/pm/track-changes";
import type { StyleDefinition } from "@/lib/pm/format";
import type { PlainTextMode } from "@/lib/serializers/plainText";

export type { Author, DocJSON, PageSetup, PaginationLayout, PlainTextMode, TrackedChange };

export type TranscriptEditorHandle = {
  /** Plain text export; "accepted" (default) hides deletions, "original" hides insertions. */
  getPlainText: (options?: { mode?: PlainTextMode }) => string;
  getJson: () => DocJSON;
  /** Replace the document (resets undo history). */
  setJson: (json: DocJSON) => void;
  /** Replace the content in place (undoable, never tracked). Keeps document attrs (layout, comments). */
  setPlainText: (plain: string) => void;
  /** Replace the whole document with plain text and default studio layout (resets undo). */
  loadPlainText: (plain: string) => void;
  focus: () => void;
  importDocx: (buffer: ArrayBuffer) => Promise<void>;
  exportDocx: () => Promise<ArrayBuffer>;
  setTrackChanges: (enabled: boolean) => void;
  isTrackingChanges: () => boolean;
  setAuthor: (author: Author) => void;
  listChanges: () => TrackedChange[];
  acceptChange: (id: string) => void;
  rejectChange: (id: string) => void;
  acceptAll: () => void;
  rejectAll: () => void;
  goToChange: (id: string) => void;
  getPageCount: () => number;
  loadTemplateStyles: (styles: StyleDefinition[]) => void;
};

export type TranscriptEditorStudioActions = {
  onPlayPause?: () => void;
  onPrint?: () => void;
};

export type TranscriptEditorProps = {
  initialContent?: string;
  initialJson?: DocJSON;
  editable?: boolean;
  className?: string;
  placeholder?: string;
  /** Who tracked changes are attributed to. */
  author?: Author;
  /** Start (or switch) track changes on/off. */
  trackChanges?: boolean;
  /** Page geometry; defaults to US Letter with 1 inch margins. */
  pageSetup?: PageSetup;
  /** Initial review view: "all" shows deletions struck through, "simple" hides them. */
  markupView?: "all" | "simple";
  studioActions?: TranscriptEditorStudioActions;
  onReady?: () => void;
  onUpdate?: () => void;
  onChangesUpdate?: (changes: TrackedChange[]) => void;
  /** Track changes was switched on/off (from the toolbar, a menu, or the handle). */
  onTrackChangesChange?: (enabled: boolean) => void;
  onLayout?: (layout: PaginationLayout) => void;
  editorRef?: Ref<TranscriptEditorHandle | null>;
};
