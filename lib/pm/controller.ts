import type { MarkType, Node as PMNode } from "prosemirror-model";
import { toggleMark } from "prosemirror-commands";
import { redo, redoDepth, undo, undoDepth } from "prosemirror-history";
import {
  EditorState,
  TextSelection,
  type Command,
  type Plugin,
  type Transaction,
} from "prosemirror-state";
import { fixTables } from "prosemirror-tables";
import { EditorView } from "prosemirror-view";
import { docJsonToPlainText, plainTextToDoc, type PlainTextMode } from "@/lib/serializers/plainText";
import { ensureStudioLayout } from "./defaultLayout";
import { characterFormatAt } from "./character";
import { insertPageBreak, isMarkActive } from "./commands";
import { type HeaderFooterTarget, setHeaderFooterText } from "./header-footer";
import { mergeTemplateStyles } from "./styles";
import {
  sectionIndexAt,
  sectionsOf,
  setSectionColumns,
  setSectionLineNumbers,
  type LineNumberMode,
} from "./sections";
import {
  addComment,
  canAddComment,
  collectComments,
  commentThreads,
  deleteComment,
  editComment,
  replyToComment,
  resolveComment,
  type CommentEntry,
  type CommentThread,
} from "./comments";
import {
  findNext,
  findState,
  replaceAll,
  replaceMatch,
  setFindQuery,
  type FindQuery,
} from "./find";
import {
  insertImage,
  linkAt,
  removeLink,
  setLink,
  type ImageAttrs,
  type LinkInfo,
} from "./objects";
import { documentOutline, type OutlineEntry } from "./outline";
import { insertTable, isInTable, setTableBorders, tableBorders } from "./tables";
import { textStats, type TextStats } from "./stats";
import {
  JUSTIFY_SHRINK_VAR,
  pageFrames,
  type DocumentLayout,
  type PageFrame,
  type PageInfo,
  type ParagraphFormat,
  type SectionLayout,
  type StyleDefinition,
} from "./format";
import type { DocJSON } from "./json";
import { paragraphFormatAt, setParagraphFormat, type FormatPatch } from "./paragraph";
import {
  ALL_MARKUP,
  LETTER_PAGE_SETUP,
  getPageCount,
  paginationKey,
  paginationPlugin,
  wordLineBreaksPlugin,
  type MarkupFilter,
  type PageSetup,
  type PaginationLayout,
} from "./pagination";
import { documentUsesLayoutEngine } from "./word-layout/constants";
import { buildPaginationContext } from "./word-layout/pageGeometry";
import { basePlugins } from "./plugins";
import { pasteKey, type PasteMode } from "./plugins/paste";
import { schema } from "./schema";
import {
  SKIP_TRACKING,
  acceptAllChanges,
  acceptChange,
  acceptChanges,
  changesInSelection,
  collectChanges,
  goToChange,
  rejectAllChanges,
  rejectChange,
  rejectChanges,
  setAuthor,
  setTrackChanges,
  setTrackingLocked,
  trackChangesKey,
  trackChangesPlugin,
  trackTransaction,
  type Author,
  type TrackedChange,
} from "./track-changes";

export type ControllerOptions = {
  doc: DocJSON;
  editable: boolean;
  placeholder: string;
  author: Author;
  trackChanges: boolean;
  /** `null` disables pagination (plain, non-studio editor). */
  pageSetup: PageSetup | null;
  /** Use `pageSetup` even when the document carries its own Word page layout. */
  pinPageSetup?: boolean;
  onLayout?: (layout: PaginationLayout) => void;
};

const DEFAULT_TAB_STOP = 48;
const FIRST_PAGE: PageInfo = { section: 0, first: true, number: 1 };

export type TransactionListener = (info: { docChanged: boolean }) => void;

export function docFromJSON(json: DocJSON): PMNode {
  return schema.nodeFromJSON(json);
}

/** The item after (or before) the selection in document order, wrapping around. */
function adjacent<T extends { from: number }>(
  items: T[],
  selection: { from: number; to: number },
  direction: 1 | -1,
): T | undefined {
  if (!items.length) return undefined;
  if (direction > 0) return items.find((c) => c.from > selection.from) ?? items[0];
  return items.findLast((c) => c.from < selection.from) ?? items[items.length - 1];
}

/** Owns one ProseMirror view: tracked dispatch, plugins, and editor commands. */
export class EditorController {
  readonly view: EditorView;
  private editable: boolean;
  private pageSetup: PageSetup;
  private readonly pinPageSetup: boolean;
  private frameCache: {
    setup: PageSetup;
    layout: DocumentLayout | null;
    sections: SectionLayout[];
    frame: (page: PageInfo) => PageFrame;
  } | null = null;
  private sectionCache: { doc: PMNode; sections: SectionLayout[] } | null = null;
  private readonly plugins: Plugin[];
  private readonly paginationEnabled: boolean;
  private readonly listeners = new Set<TransactionListener>();
  private compositionRanges: { from: number; to: number }[] = [];
  private markupFilter: MarkupFilter = ALL_MARKUP;

  constructor(mount: HTMLElement, options: ControllerOptions) {
    this.editable = options.editable;
    this.pageSetup = options.pageSetup ?? LETTER_PAGE_SETUP;
    this.pinPageSetup = !!options.pinPageSetup;
    this.paginationEnabled = !!options.pageSetup;
    const initialTracking = { enabled: options.trackChanges, author: options.author };

    this.plugins = [
      ...basePlugins({ placeholder: options.placeholder }),
      trackChangesPlugin(() => ({ ...initialTracking })),
      ...(options.pageSetup
        ? [
            paginationPlugin({
              getSetup: () => this.getPageSetup(),
              getFrame: (page) => this.getPageFrame(page),
              getPaginationContext: () => this.getPaginationContext(),
              hasWordLayout: () => documentUsesLayoutEngine(this.view.state.doc),
              getNumberStart: (section) => this.getSections()[section]?.pageNumberStart,
              getDocGridLinePitchTwips: (section) =>
                this.getDocGridLinePitchTwips(section),
              getDefaultTabStop: () => this.getLayout()?.defaultTabStop ?? DEFAULT_TAB_STOP,
              getMarkupFilter: () => this.markupFilter,
              onLayout: options.onLayout,
            }),
            wordLineBreaksPlugin(),
          ]
        : []),
    ];

    const initialDoc = this.paginationEnabled
      ? ensureStudioLayout(options.doc, this.pageSetup)
      : options.doc;

    this.view = new EditorView(mount, {
      state: this.createState(docFromJSON(initialDoc)),
      editable: () => this.editable,
      dispatchTransaction: (tr) => this.dispatch(tr),
      attributes: (state) => {
        const shrink = (state.doc.attrs.layout as DocumentLayout | null)?.justifyShrink;
        return {
          class: "transcript-editor-prose",
          spellcheck: "true",
          ...(shrink ? { style: `${JUSTIFY_SHRINK_VAR}: ${shrink}` } : {}),
        };
      },
      handleDOMEvents: {
        compositionend: () => {
          setTimeout(() => this.flushComposition(), 0);
          return false;
        },
      },
    });
  }

  private createState(doc: PMNode, tracking = this.trackingState()) {
    let state = EditorState.create({ schema, doc, plugins: this.plugins });
    // Imported tables can be ragged (e.g. Word's gridBefore); table commands need them rectangular.
    const fix = fixTables(state);
    if (fix) state = state.apply(fix);
    // Carry track-changes settings over when the whole state is replaced.
    return tracking
      ? state.apply(state.tr.setMeta(trackChangesKey, tracking))
      : state;
  }

  private trackingState() {
    return this.view ? trackChangesKey.getState(this.view.state) : undefined;
  }

  get state(): EditorState {
    return this.view.state;
  }

  dispatch = (tr: Transaction) => {
    const state = this.view.state;
    if (tr.getMeta("composition") !== undefined && tr.docChanged) {
      this.recordComposition(tr, state);
    } else if (tr.docChanged) {
      this.compositionRanges = this.compositionRanges.map((r) => ({
        from: tr.mapping.map(r.from, -1),
        to: tr.mapping.map(r.to, 1),
      }));
    }
    const applied = trackTransaction(tr, state);
    this.view.updateState(state.apply(applied));
    for (const listener of this.listeners) listener({ docChanged: applied.docChanged });
  };

  private recordComposition(tr: Transaction, state: EditorState) {
    if (!trackChangesKey.getState(state)?.enabled) return;
    const ranges = this.compositionRanges.map((r) => ({
      from: tr.mapping.map(r.from, -1),
      to: tr.mapping.map(r.to, 1),
    }));
    tr.steps.forEach((step, i) => {
      step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        const rest = tr.mapping.slice(i + 1);
        ranges.push({ from: rest.map(newStart, -1), to: rest.map(newEnd, 1) });
      });
    });
    this.compositionRanges = ranges.filter((r) => r.to > r.from);
  }

  /** Mark text typed through an IME as an insertion once composition ends. */
  private flushComposition() {
    const ranges = this.compositionRanges;
    this.compositionRanges = [];
    const ts = trackChangesKey.getState(this.view.state);
    if (!ranges.length || !ts?.enabled || this.view.composing) return;
    const { insertion } = schema.marks;
    const mark = insertion.create({
      id: `ime${Date.now().toString(36)}`,
      authorId: ts.author.id,
      authorName: ts.author.name,
      date: new Date().toISOString(),
    });
    const tr = this.view.state.tr.setMeta(SKIP_TRACKING, true);
    for (const r of ranges) tr.addMark(r.from, r.to, mark);
    this.dispatch(tr);
  }

  subscribe(listener: TransactionListener) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  destroy() {
    this.listeners.clear();
    this.view.destroy();
  }

  run(command: Command): boolean {
    return command(this.view.state, this.dispatch, this.view);
  }

  can(command: Command): boolean {
    return command(this.view.state, undefined, this.view);
  }

  focus() {
    this.view.focus();
  }

  setEditable(editable: boolean) {
    this.editable = editable;
    this.view.setProps({ editable: () => this.editable });
  }

  // —— Content ——

  getJson(): DocJSON {
    return this.view.state.doc.toJSON() as DocJSON;
  }

  getPlainText(mode: PlainTextMode = "accepted"): string {
    return docJsonToPlainText(this.getJson(), { mode });
  }

  /** Replace the whole document and reset undo history. */
  load(json: DocJSON) {
    const doc = this.paginationEnabled ? ensureStudioLayout(json, this.pageSetup) : json;
    this.view.updateState(this.createState(docFromJSON(doc)));
    for (const listener of this.listeners) listener({ docChanged: true });
  }

  loadPlainText(plain: string) {
    const json = plainTextToDoc(plain);
    this.load(this.paginationEnabled ? ensureStudioLayout(json, this.pageSetup) : json);
  }

  /** Replace the content in place (undoable, never tracked). */
  setPlainText(plain: string) {
    const doc = docFromJSON(plainTextToDoc(plain));
    const { state } = this.view;
    this.dispatch(
      state.tr
        .replaceWith(0, state.doc.content.size, doc.content)
        .setMeta(SKIP_TRACKING, true),
    );
  }

  // —— Formatting ——

  /** Run a command from toolbar UI and return focus to the text. */
  exec(command: Command) {
    const ok = this.run(command);
    this.view.focus();
    return ok;
  }

  /** Run a command without tracking it (table structure, like Word's untracked layout edits). */
  execUntracked(command: Command) {
    const ok = command(this.view.state, (tr) => this.dispatch(tr.setMeta(SKIP_TRACKING, true)), this.view);
    this.view.focus();
    return ok;
  }

  /** Word's Insert Table, sized to the text width. */
  insertTable(rows: number, cols: number) {
    const setup = this.getPageSetup();
    return this.execUntracked(insertTable(rows, cols, setup.width - setup.marginLeft - setup.marginRight));
  }

  isInTable() {
    return isInTable(this.view.state);
  }

  tableBorders() {
    return tableBorders(this.view.state);
  }

  setTableBorders(borders: boolean) {
    return this.execUntracked(setTableBorders(borders));
  }

  getLink(): LinkInfo | null {
    return linkAt(this.view.state);
  }

  setLink(href: string, text?: string) {
    return this.exec(setLink(href, text));
  }

  removeLink() {
    return this.exec(removeLink);
  }

  insertImage(attrs: ImageAttrs) {
    return this.exec(insertImage(attrs));
  }

  toggleMark(type: MarkType) {
    this.exec(toggleMark(type));
  }

  getParagraphFormat(): ParagraphFormat {
    return paragraphFormatAt(this.view.state);
  }

  setParagraphFormat(patch: FormatPatch) {
    return this.exec(setParagraphFormat(patch));
  }

  isMarkActive(type: MarkType) {
    return isMarkActive(this.view.state, type);
  }

  /** Character formatting at the selection, with the effective font and size. */
  getCharacterFormat() {
    return characterFormatAt(this.view.state);
  }

  getPasteMode(): PasteMode {
    return pasteKey.getState(this.view.state)?.mode ?? "auto";
  }

  setPasteMode(mode: PasteMode) {
    this.dispatch(
      this.view.state.tr
        .setMeta(pasteKey, { mode })
        .setMeta("addToHistory", false)
        .setMeta(SKIP_TRACKING, true),
    );
  }

  undo() {
    this.run(undo);
  }

  redo() {
    this.run(redo);
  }

  canUndo() {
    return undoDepth(this.view.state) > 0;
  }

  canRedo() {
    return redoDepth(this.view.state) > 0;
  }

  insertPageBreak() {
    this.run(insertPageBreak);
    this.view.focus();
  }

  // —— Find and replace (leaves focus in the Find dialog) ——

  getFind(): { query: FindQuery; count: number; current: number } {
    const { query, matches, current } = findState(this.view.state);
    return { query, count: matches.length, current };
  }

  setFind(query: FindQuery) {
    this.run(setFindQuery(query));
  }

  findNext(direction: 1 | -1 = 1) {
    const ok = this.run(findNext(direction));
    if (ok) this.revealCurrentMatch();
    return ok;
  }

  /** Replace the selected match and select the next one. */
  replace(replacement: string) {
    if (this.run(replaceMatch(replacement))) this.run(findNext(1));
    this.revealCurrentMatch();
  }

  /** ProseMirror only scrolls to the selection while it has focus, which the Find dialog holds. */
  private revealCurrentMatch() {
    this.view.dom
      .querySelector(".pm-find-match--current")
      ?.scrollIntoView({ block: "center", inline: "nearest" });
  }

  /** Replace every match; returns how many were replaced. */
  replaceAll(replacement: string): number {
    const count = findState(this.view.state).matches.length;
    return this.run(replaceAll(replacement)) ? count : 0;
  }

  // —— Document info ——

  /** Word count for the document, or for the selection when `selection` is set and non-empty. */
  getStats(selection = false): TextStats {
    const { doc, selection: sel } = this.view.state;
    return selection && !sel.empty ? textStats(doc, sel.from, sel.to) : textStats(doc);
  }

  getOutline(): OutlineEntry[] {
    return documentOutline(this.view.state.doc);
  }

  /** Put the cursor at the start of the block at `pos` (e.g. a heading in the navigation pane). */
  goToBlock(pos: number) {
    const { doc } = this.view.state;
    const $pos = doc.resolve(Math.min(pos + 1, doc.content.size));
    this.dispatch(this.view.state.tr.setSelection(TextSelection.near($pos)));
    this.view.focus();
  }

  // —— Track changes ——

  isTracking(): boolean {
    return !!trackChangesKey.getState(this.view.state)?.enabled;
  }

  getAuthor(): Author | undefined {
    return trackChangesKey.getState(this.view.state)?.author;
  }

  /** Returns false when tracking is locked on. */
  setTrackChanges(enabled: boolean) {
    return this.run(setTrackChanges(enabled));
  }

  isTrackingLocked(): boolean {
    return !!trackChangesKey.getState(this.view.state)?.locked;
  }

  setTrackingLocked(locked: boolean) {
    this.run(setTrackingLocked(locked));
  }

  setAuthor(author: Author) {
    this.run(setAuthor(author));
  }

  listChanges(): TrackedChange[] {
    return collectChanges(this.view.state.doc);
  }

  acceptChange(id: string) {
    return this.run(acceptChange(id));
  }

  rejectChange(id: string) {
    return this.run(rejectChange(id));
  }

  acceptAll() {
    return this.run(acceptAllChanges);
  }

  rejectAll() {
    return this.run(rejectAllChanges);
  }

  goToChange(id: string) {
    const ok = this.run(goToChange(id));
    if (ok) this.view.focus();
    return ok;
  }

  /** Word's Previous / Next: the change before or after the selection, wrapping around. */
  goToAdjacentChange(direction: 1 | -1) {
    const target = adjacent(this.listChanges(), this.view.state.selection, direction);
    return target ? this.goToChange(target.id) : false;
  }

  /** Accept (or reject) the changes in the selection, then move to the next change, like Word. */
  resolveSelection(action: "accept" | "reject") {
    const ids = changesInSelection(this.view.state).map((c) => c.id);
    if (!ids.length) return this.goToAdjacentChange(1);
    this.run(action === "accept" ? acceptChanges(ids) : rejectChanges(ids));
    const next = this.listChanges().find((c) => c.from >= this.view.state.selection.from);
    if (next) this.goToChange(next.id);
    else this.view.focus();
    return true;
  }

  // —— Comments ——

  getComments(): CommentEntry[] {
    return collectComments(this.view.state.doc);
  }

  getCommentThreads(): Record<string, CommentThread> {
    return commentThreads(this.view.state.doc);
  }

  canAddComment() {
    return canAddComment(this.view.state);
  }

  addComment(text: string) {
    const author = this.getAuthor() ?? { id: "local", name: "Author" };
    return this.exec(addComment(text, author));
  }

  editComment(id: string, text: string) {
    return this.run(editComment(id, text));
  }

  replyToComment(id: string, text: string) {
    const author = this.getAuthor() ?? { id: "local", name: "Author" };
    return this.run(replyToComment(id, text, author));
  }

  resolveComment(id: string, resolved: boolean) {
    return this.run(resolveComment(id, resolved));
  }

  deleteComment(id: string) {
    return this.run(deleteComment(id));
  }

  goToComment(id: string) {
    const comment = this.getComments().find((c) => c.id === id);
    if (!comment) return false;
    this.dispatch(
      this.view.state.tr
        .setSelection(TextSelection.create(this.view.state.doc, comment.from, comment.to))
        .scrollIntoView(),
    );
    this.view.focus();
    return true;
  }

  goToAdjacentComment(direction: 1 | -1) {
    const target = adjacent(this.getComments(), this.view.state.selection, direction);
    return target ? this.goToComment(target.id) : false;
  }

  // —— Pagination ——

  /** Word layout imported with the document (page, header/footer, line numbers). */
  getLayout(): DocumentLayout | null {
    return (this.view.state.doc.attrs.layout as DocumentLayout | null) ?? null;
  }

  getPageSetup(): PageSetup {
    if (this.pinPageSetup) return this.pageSetup;
    return this.getLayout()?.page ?? this.pageSetup;
  }

  /** Layout of every section: the document's section breaks, then the document layout. */
  getSections(): SectionLayout[] {
    const doc = this.view.state.doc;
    if (this.sectionCache?.doc !== doc) {
      const prev = this.sectionCache?.sections;
      const sections = sectionsOf(this.view.state);
      const same = prev && JSON.stringify(prev) === JSON.stringify(sections);
      this.sectionCache = { doc, sections: same ? prev : sections };
    }
    return this.sectionCache!.sections;
  }

  getDocGridLinePitchTwips(section: number): number | undefined {
    return this.getSections()[section]?.docGridLinePitch;
  }

  /** Effective top/bottom margins of a page, grown to fit its header/footer. */
  getPageFrame(page: PageInfo = FIRST_PAGE): PageFrame {
    const setup = this.getPageSetup();
    const layout = this.getLayout();
    const sections = this.getSections();
    const cache = this.frameCache;
    if (!cache || cache.setup !== setup || cache.layout !== layout || cache.sections !== sections) {
      this.frameCache = { setup, layout, sections, frame: pageFrames(setup, layout, sections) };
    }
    return this.frameCache!.frame(page);
  }

  getPaginationContext() {
    return buildPaginationContext(this.getLayout(), this.getSections(), this.getPageSetup());
  }

  setPageSetup(setup: PageSetup) {
    const prev = this.getPageSetup();
    const keys = Object.keys(setup) as (keyof PageSetup)[];
    if (keys.every((k) => prev[k] === setup[k])) return;
    this.pageSetup = setup;
    const layout = this.getLayout();
    if (layout && !this.pinPageSetup) {
      this.dispatch(
        this.view.state.tr
          .setDocAttribute("layout", { ...layout, page: { ...setup } })
          .setMeta(SKIP_TRACKING, true),
      );
    }
    this.relayout();
  }

  getMarkupFilter(): MarkupFilter {
    return this.markupFilter;
  }

  /** Which revisions and comments get balloons and change bars (the shell hides their styling). */
  setMarkupFilter(filter: MarkupFilter) {
    const prev = this.markupFilter;
    if ((Object.keys(filter) as (keyof MarkupFilter)[]).every((k) => prev[k] === filter[k])) return;
    this.markupFilter = filter;
    this.relayout();
  }

  /** Re-paginate after a change the editor can't see, e.g. CSS hiding deleted text. */
  relayout() {
    if (!paginationKey.getState(this.view.state)) return;
    this.dispatch(
      this.view.state.tr
        .setMeta(paginationKey, { breaks: [] })
        .setMeta("addToHistory", false)
        .setMeta(SKIP_TRACKING, true),
    );
  }

  getPageCount(): number {
    return getPageCount(this.view.state);
  }

  setHeaderFooter(target: HeaderFooterTarget, text: string) {
    const tr = setHeaderFooterText(this.view.state, target, text);
    if (!tr) return;
    this.dispatch(tr);
    this.relayout();
  }

  setSectionColumns(count: number, space: number) {
    const idx = sectionIndexAt(this.view.state, this.view.state.selection.head);
    const tr = setSectionColumns(this.view.state, idx, count, space);
    if (!tr) return;
    this.dispatch(tr);
    this.relayout();
  }

  setLineNumbers(mode: LineNumberMode) {
    const idx = sectionIndexAt(this.view.state, this.view.state.selection.head);
    const tr = setSectionLineNumbers(this.view.state, idx, mode);
    if (!tr) return;
    this.dispatch(tr);
    this.relayout();
  }

  loadTemplateStyles(styles: StyleDefinition[]) {
    if (!styles.length) return;
    this.dispatch(mergeTemplateStyles(this.view.state, styles));
  }
}
