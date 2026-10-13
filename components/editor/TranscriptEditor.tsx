"use client";

import {
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import "prosemirror-view/style/prosemirror.css";
import { EditorController } from "@/lib/pm/controller";
import {
  ALL_MARKUP,
  LETTER_PAGE_SETUP,
  pageStride,
  stackHeight,
  type ChangeBalloon,
  type LinePlacement,
  type MarkupFilter,
  type PageSetup,
} from "@/lib/pm/pagination";
import type { PageInfo } from "@/lib/pm/format";
import type { Author } from "@/lib/pm/track-changes";
import { safeHref } from "@/lib/pm/schema";
import { TABLE_COMMANDS } from "@/lib/pm/tables";
import { plainTextToDoc } from "@/lib/serializers/plainText";
import { TextSelection } from "prosemirror-state";
import { CommentDialog } from "./CommentDialog";
import { ContextMenu, type ContextMenuState } from "./ContextMenu";
import { LinkDialog } from "./InsertDialogs";
import { HeaderFooterDialog } from "./HeaderFooterDialog";
import {
  getHeaderFooterText,
  headerFooterTargetAt,
  type HeaderFooterTarget,
} from "@/lib/pm/header-footer";
import { NavigationPane } from "./NavigationPane";
import { Ruler } from "./Ruler";
import { StatusBar } from "./StatusBar";
import { StudioEditorChrome } from "./StudioEditorChrome";
import { StudioPageBackgrounds, type BalloonAction } from "./StudioPageBackgrounds";
import type { MenuItem } from "./toolbarParts";
import type { TranscriptEditorHandle, TranscriptEditorProps } from "./types";
import { ZOOM_STEP, useZoom } from "./useZoom";
import "./styles/editor.css";

const DEFAULT_AUTHOR: Author = { id: "local", name: "Typist" };
const DEFAULT_PLACEHOLDER = "Start typing the transcript\u2026";
/** Word's markup area beside each page when formatting revisions are shown in balloons. */
const MARKUP_WIDTH = 300;

function pageCssVars(setup: PageSetup, firstPageTop: number, markupWidth: number) {
  return {
    "--page-width": `${setup.width}px`,
    "--markup-width": `${markupWidth}px`,
    "--page-height": `${setup.height}px`,
    "--page-gap": `${setup.gap}px`,
    "--margin-top": `${firstPageTop}px`,
    "--margin-right": `${setup.marginRight}px`,
    "--margin-bottom": `${setup.marginBottom}px`,
    "--margin-left": `${setup.marginLeft}px`,
  } as React.CSSProperties;
}

/** Room left above a heading or page scrolled to the top, below the sticky ruler. */
const SCROLL_MARGIN = 40;

type CommentDialogState = { mode: "new" } | { mode: "reply" | "edit"; id: string };
type LinkDialogState = { text: string; href: string; existing: boolean };
type HfDialogState = { target: HeaderFooterTarget; title: string };

async function loadDocxModule() {
  return import("@/lib/docx");
}

/** 0-based page holding the cursor, from its position in the page stack. */
function pageOfSelection(c: EditorController) {
  const dom = c.view.dom as HTMLElement;
  try {
    const rect = dom.getBoundingClientRect();
    const scale = dom.offsetHeight ? rect.height / dom.offsetHeight : 1;
    const y = (c.view.coordsAtPos(c.state.selection.head).top - rect.top) / scale;
    return Math.max(0, Math.floor(y / pageStride(c.getPageSetup())));
  } catch {
    return 0;
  }
}

export function TranscriptEditor(props: TranscriptEditorProps) {
  const {
    initialContent,
    initialJson,
    editable = true,
    className = "",
    author,
    trackChanges = false,
    pageSetup,
    markupView = "all",
    studioActions,
    editorRef,
  } = props;
  const latest = useRef(props);
  useEffect(() => {
    latest.current = props;
  });

  const loadedContent = useRef<string | undefined>(undefined);
  const [controller, setController] = useState<EditorController | null>(null);
  const [, setVersion] = useState(0);
  const [pageCount, setPageCount] = useState(1);
  const [pages, setPages] = useState<PageInfo[]>([]);
  const [lines, setLines] = useState<LinePlacement[]>([]);
  const [changedLines, setChangedLines] = useState<LinePlacement[]>([]);
  const [balloons, setBalloons] = useState<ChangeBalloon[]>([]);
  const [simpleMarkup, setSimpleMarkup] = useState(markupView === "simple");
  const [cursorPage, setCursorPage] = useState(0);
  const [navigation, setNavigation] = useState(false);
  const [ruler, setRuler] = useState(true);
  const [markup, setMarkup] = useState<MarkupFilter>(ALL_MARKUP);
  const [commentDialog, setCommentDialog] = useState<CommentDialogState | null>(null);
  const [linkDialog, setLinkDialog] = useState<LinkDialogState | null>(null);
  const [hfDialog, setHfDialog] = useState<HfDialogState | null>(null);
  const [menu, setMenu] = useState<ContextMenuState | null>(null);
  const closeMenu = useCallback(() => setMenu(null), []);
  const [zoom, setZoom] = useZoom();
  const scrollRef = useRef<HTMLDivElement>(null);
  const paperRef = useRef<HTMLDivElement>(null);
  const zoomRef = useRef(zoom);

  const mountRef = useCallback(
    (el: HTMLDivElement | null) => {
      if (!el) return;
      const p = latest.current;
      loadedContent.current = p.initialContent;
      const c = new EditorController(el, {
        doc: p.initialJson ?? plainTextToDoc(p.initialContent ?? ""),
        editable: p.editable ?? true,
        placeholder: p.placeholder ?? DEFAULT_PLACEHOLDER,
        author: p.author ?? DEFAULT_AUTHOR,
        trackChanges: p.trackChanges ?? false,
        pageSetup: p.pageSetup ?? LETTER_PAGE_SETUP,
        pinPageSetup: !!p.pageSetup,
        onLayout: (layout) => {
          setPageCount(layout.pageCount);
          setPages(layout.pages);
          setLines(layout.lines);
          setChangedLines(layout.changedLines);
          setBalloons(layout.balloons);
          latest.current.onLayout?.(layout);
        },
      });
      let frame = 0;
      let tracking = c.isTracking();
      const unsubscribe = c.subscribe(({ docChanged }) => {
        if (!frame) {
          frame = requestAnimationFrame(() => {
            frame = 0;
            setVersion((v) => v + 1);
            setCursorPage(pageOfSelection(c));
          });
        }
        if (c.isTracking() !== tracking) {
          tracking = c.isTracking();
          latest.current.onTrackChangesChange?.(tracking);
        }
        if (docChanged) {
          latest.current.onUpdate?.();
          latest.current.onChangesUpdate?.(c.listChanges());
        }
      });
      setController(c);
      return () => {
        cancelAnimationFrame(frame);
        unsubscribe();
        c.destroy();
      };
    },
    [],
  );

  useEffect(() => {
    if (!controller) return;
    latest.current.onChangesUpdate?.(controller.listChanges());
    latest.current.onReady?.();
  }, [controller]);

  useEffect(() => {
    controller?.setEditable(editable);
  }, [controller, editable]);

  useEffect(() => {
    controller?.relayout();
  }, [controller, simpleMarkup]);

  // After render, so the filter's CSS (hidden deletions) is in place before re-measuring.
  useEffect(() => {
    controller?.setMarkupFilter(markup);
  }, [controller, markup]);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey || !e.deltaY) return;
      e.preventDefault();
      setZoom((z) => z + (e.deltaY < 0 ? ZOOM_STEP : -ZOOM_STEP));
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, [setZoom]);

  // Keep the same part of the document in view when zooming.
  useLayoutEffect(() => {
    const el = scrollRef.current;
    const prev = zoomRef.current;
    zoomRef.current = zoom;
    if (el && prev !== zoom) el.scrollTop = (el.scrollTop * zoom) / prev;
  }, [zoom]);

  useEffect(() => {
    if (controller && controller.isTracking() !== trackChanges) {
      controller.setTrackChanges(trackChanges);
    }
  }, [controller, trackChanges]);

  const authorId = author?.id;
  const authorName = author?.name;
  useEffect(() => {
    if (controller && authorId) {
      controller.setAuthor({ id: authorId, name: authorName ?? authorId });
    }
  }, [controller, authorId, authorName]);

  useEffect(() => {
    if (!controller || !pageSetup) return;
    controller.setPageSetup(pageSetup);
  }, [controller, pageSetup]);

  useEffect(() => {
    if (!controller || initialJson) return;
    if (initialContent === undefined || initialContent === loadedContent.current) return;
    loadedContent.current = initialContent;
    controller.loadPlainText(initialContent);
  }, [controller, initialContent, initialJson]);

  useImperativeHandle(
    editorRef,
    (): TranscriptEditorHandle => ({
      getPlainText: (options) => controller?.getPlainText(options?.mode) ?? "",
      getJson: () => controller?.getJson() ?? { type: "doc", content: [] },
      setJson: (json) => controller?.load(json),
      setPlainText: (plain) => controller?.setPlainText(plain),
      focus: () => controller?.focus(),
      importDocx: async (buffer) => {
        if (!controller) throw new Error("Editor not ready");
        const { importDocx } = await loadDocxModule();
        const result = await importDocx(buffer);
        controller.load(result.doc);
      },
      exportDocx: async () => {
        if (!controller) throw new Error("Editor not ready");
        const { exportDocx } = await loadDocxModule();
        return exportDocx(controller.getJson(), controller.getPageSetup(), {
          trackRevisions: controller.isTracking(),
        });
      },
      setTrackChanges: (enabled) => controller?.setTrackChanges(enabled),
      isTrackingChanges: () => controller?.isTracking() ?? false,
      setAuthor: (next) => controller?.setAuthor(next),
      listChanges: () => controller?.listChanges() ?? [],
      acceptChange: (id) => void controller?.acceptChange(id),
      rejectChange: (id) => void controller?.rejectChange(id),
      acceptAll: () => void controller?.acceptAll(),
      rejectAll: () => void controller?.rejectAll(),
      goToChange: (id) => void controller?.goToChange(id),
      getPageCount: () => controller?.getPageCount() ?? 1,
      loadTemplateStyles: (styles) => controller?.loadTemplateStyles(styles),
    }),
    [controller],
  );

  const setup = controller?.getPageSetup() ?? pageSetup ?? LETTER_PAGE_SETUP;
  const layout = controller?.getLayout() ?? null;
  const firstPageTop = controller?.getPageFrame().top ?? setup.marginTop;
  const shownBalloons = simpleMarkup ? balloons.filter((b) => b.changeType === "comment") : balloons;
  const markupWidth =
    shownBalloons.length || (!simpleMarkup && changedLines.length) ? MARKUP_WIDTH : 0;
  const paperWidth = (setup.width + markupWidth) * zoom;
  const page = Math.min(cursorPage, pageCount - 1);
  const threads = controller?.getCommentThreads() ?? {};

  const changeMenu = (id: string): MenuItem[] => [
    { label: "Accept change", run: () => controller?.acceptChange(id) },
    { label: "Reject change", run: () => controller?.rejectChange(id) },
  ];
  const commentMenu = (id: string): MenuItem[] => {
    const resolved = !!threads[id]?.resolved;
    return [
      { label: "Reply to comment", run: () => setCommentDialog({ mode: "reply", id }) },
      { label: "Edit comment", run: () => setCommentDialog({ mode: "edit", id }) },
      {
        label: resolved ? "Reopen comment" : "Resolve comment",
        run: () => controller?.resolveComment(id, !resolved),
      },
      { label: "Delete comment", run: () => controller?.deleteComment(id) },
    ];
  };

  const onBalloonAction = (balloon: ChangeBalloon, action: BalloonAction) => {
    if (!controller) return;
    const { id } = balloon;
    switch (action) {
      case "select":
        if (balloon.changeType === "comment") controller.goToComment(id);
        else controller.goToChange(id);
        break;
      case "accept":
        controller.acceptChange(id);
        break;
      case "reject":
        controller.rejectChange(id);
        break;
      case "reply":
        setCommentDialog({ mode: "reply", id });
        break;
      case "resolve":
      case "reopen":
        controller.resolveComment(id, action === "resolve");
        break;
      case "delete":
        controller.deleteComment(id);
        break;
    }
  };

  const openLinkDialog = () => {
    if (!controller || !editable) return;
    const link = controller.getLink();
    const { from, to } = controller.state.selection;
    setLinkDialog(
      link
        ? { text: link.text, href: link.href, existing: true }
        : { text: controller.state.doc.textBetween(from, to, " "), href: "", existing: false },
    );
  };
  const linkMenu = (anchor: HTMLElement): MenuItem[] => {
    const href = safeHref(anchor.dataset.href);
    return [
      { label: "Edit hyperlink\u2026", run: openLinkDialog },
      {
        label: "Open hyperlink",
        disabled: !href || href.startsWith("#"),
        run: () => href && window.open(href, "_blank", "noopener,noreferrer"),
      },
      {
        label: "Copy hyperlink",
        disabled: !href,
        run: () => href && void navigator.clipboard?.writeText(href),
      },
      { label: "Remove hyperlink", run: () => controller?.removeLink() },
    ];
  };

  /** Word's right-click Accept / Reject on a revision, and comment / link / table actions. */
  const onDocumentMenu = (e: React.MouseEvent) => {
    if (!editable || !controller) return;
    const target = e.target as HTMLElement;
    const change = target.closest<HTMLElement>("[data-change-id]")?.dataset.changeId;
    const comment = target.closest<HTMLElement>("[data-comment-id]")?.dataset.commentId;
    const anchor = target.closest<HTMLElement>("a.pm-link");
    // Like Word, right-clicking outside the selection moves the cursor there first.
    const hit = controller.view.posAtCoords({ left: e.clientX, top: e.clientY });
    const { from, to } = controller.state.selection;
    if (hit && (hit.pos < from || hit.pos > to)) {
      controller.run((state, dispatch) => {
        dispatch?.(state.tr.setSelection(TextSelection.near(state.doc.resolve(hit.pos))));
        return true;
      });
    }
    const inCell = !!target.closest(".pm-table td, .pm-table th") && controller.isInTable();
    const items = [
      ...(change ? changeMenu(change) : []),
      ...(comment ? commentMenu(comment) : []),
      ...(anchor ? linkMenu(anchor) : []),
      ...(inCell
        ? TABLE_COMMANDS.map(({ label, command }) => ({
            label,
            disabled: !controller.can(command),
            run: () => controller.execUntracked(command),
          }))
        : []),
    ];
    if (!items.length) return;
    e.preventDefault();
    setMenu({ x: e.clientX, y: e.clientY, items });
  };

  const submitComment = (text: string) => {
    if (!controller || !commentDialog) return;
    if (commentDialog.mode === "new") controller.addComment(text);
    else if (commentDialog.mode === "reply") controller.replyToComment(commentDialog.id, text);
    else controller.editComment(commentDialog.id, text);
    setCommentDialog(null);
  };
  const dialogThread = commentDialog && commentDialog.mode !== "new" ? threads[commentDialog.id] : undefined;
  const dialogQuote =
    commentDialog && commentDialog.mode !== "new"
      ? controller?.getComments().find((c) => c.id === commentDialog.id)?.quote
      : undefined;

  /** Scroll so client y `top` sits just below the ruler. */
  const scrollToClientY = (top: number) => {
    const el = scrollRef.current;
    if (el) el.scrollTop += top - el.getBoundingClientRect().top - SCROLL_MARGIN;
  };
  const goToHeading = (pos: number) => {
    if (!controller) return;
    controller.goToBlock(pos);
    try {
      scrollToClientY(controller.view.coordsAtPos(controller.state.selection.head).top);
    } catch {
      // Position not rendered; the cursor still moved.
    }
  };
  const goToPage = (index: number) => {
    const paper = paperRef.current;
    if (paper) scrollToClientY(paper.getBoundingClientRect().top + index * pageStride(setup) * zoom);
  };

  return (
    <div
      className={[
        "transcript-editor-root",
        simpleMarkup ? "transcript-editor-root--simple-markup" : "",
        markup.insertions ? "" : "transcript-editor-root--hide-insdel",
        markup.comments ? "" : "transcript-editor-root--hide-comments",
        className,
      ]
        .filter(Boolean)
        .join(" ")}
    >
      <StudioEditorChrome
        controller={controller}
        readOnly={!editable}
        actions={studioActions}
        simpleMarkup={simpleMarkup}
        onSimpleMarkupChange={setSimpleMarkup}
        view={{ zoom, setZoom, navigation, setNavigation, ruler, setRuler }}
        review={{
          markup,
          setMarkup,
          onNewComment: () => setCommentDialog({ mode: "new" }),
          onLink: openLinkDialog,
        }}
      />
      <div className="transcript-editor-body">
        {navigation && controller && (
          <NavigationPane
            controller={controller}
            setup={setup}
            pageCount={pageCount}
            lines={lines}
            currentPage={page}
            onHeading={goToHeading}
            onPage={goToPage}
            onClose={() => setNavigation(false)}
          />
        )}
        <div ref={scrollRef} className="transcript-editor-scroll">
          {ruler && controller && (
            <div className="studio-ruler-bar">
              <div style={{ width: paperWidth }}>
                <Ruler
                  setup={setup}
                  zoom={zoom}
                  format={controller.getParagraphFormat()}
                  disabled={!editable}
                  onChange={(patch) => controller.setParagraphFormat(patch)}
                />
              </div>
            </div>
          )}
          <div
            className="paper-zoom"
            style={{ width: paperWidth, height: stackHeight(setup, pageCount) * zoom }}
          >
            <div
              ref={paperRef}
              className={["paper", layout ? "paper--word" : ""].filter(Boolean).join(" ")}
              style={{
                ...pageCssVars(setup, firstPageTop, markupWidth),
                transform: zoom === 1 ? undefined : `scale(${zoom})`,
              }}
            >
              <StudioPageBackgrounds
                setup={setup}
                pageCount={pageCount}
                pages={pages}
                layout={layout}
                sections={controller?.getSections() ?? []}
                lines={lines}
                changedLines={changedLines}
                balloons={shownBalloons}
                markupWidth={markupWidth}
                simpleMarkup={simpleMarkup}
                threads={threads}
                readOnly={!editable}
                onBalloonAction={onBalloonAction}
                onBalloonMenu={(balloon, x, y) => {
                  if (!editable) return;
                  const items =
                    balloon.changeType === "comment" ? commentMenu(balloon.id) : changeMenu(balloon.id);
                  setMenu({ x, y, items });
                }}
                onHeaderFooterEdit={(page, part) => {
                  if (!controller || !editable) return;
                  const target = headerFooterTargetAt(
                    controller.state,
                    page,
                    part,
                    !!layout?.evenAndOddHeaders,
                  );
                  const title = `${part === "header" ? "Header" : "Footer"} (page ${page.number})`;
                  setHfDialog({ target, title });
                }}
              />
              <div ref={mountRef} className="transcript-editor-mount" onContextMenu={onDocumentMenu} />
            </div>
          </div>
        </div>
      </div>
      {controller && (
        <StatusBar
          page={page}
          pageCount={pageCount}
          stats={controller.getStats()}
          selection={controller.state.selection.empty ? null : controller.getStats(true)}
          tracking={controller.isTracking()}
          readOnly={!editable || controller.isTrackingLocked()}
          onToggleTracking={() => controller.setTrackChanges(!controller.isTracking())}
          zoom={zoom}
          onZoom={setZoom}
        />
      )}
      {hfDialog && controller && (
        <HeaderFooterDialog
          title={hfDialog.title}
          initial={getHeaderFooterText(controller.state, hfDialog.target)}
          onClose={() => {
            setHfDialog(null);
            controller.focus();
          }}
          onSubmit={(text) => {
            controller.setHeaderFooter(hfDialog.target, text);
            setHfDialog(null);
            controller.focus();
          }}
        />
      )}
      {linkDialog && controller && (
        <LinkDialog
          initialText={linkDialog.text}
          initialHref={linkDialog.href}
          onClose={() => {
            setLinkDialog(null);
            controller.focus();
          }}
          onSubmit={(href, text) => {
            setLinkDialog(null);
            controller.setLink(href, text);
          }}
          onRemove={
            linkDialog.existing
              ? () => {
                  setLinkDialog(null);
                  controller.removeLink();
                }
              : undefined
          }
        />
      )}
      {menu && <ContextMenu {...menu} onClose={closeMenu} />}
      {commentDialog && (
        <CommentDialog
          key={commentDialog.mode === "new" ? "new" : `${commentDialog.mode}:${commentDialog.id}`}
          title={
            commentDialog.mode === "new"
              ? "New comment"
              : commentDialog.mode === "reply"
                ? `Reply to ${dialogThread?.authorName ?? "comment"}`
                : "Edit comment"
          }
          initial={commentDialog.mode === "edit" ? dialogThread?.text : ""}
          quote={commentDialog.mode === "new" ? undefined : dialogQuote}
          onClose={() => setCommentDialog(null)}
          onSubmit={submitComment}
        />
      )}
    </div>
  );
}

export type { TranscriptEditorHandle, TranscriptEditorProps };
