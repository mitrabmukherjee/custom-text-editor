import type { CSSProperties, ReactNode } from "react";
import {
  headerFooterFor,
  headerFooterHeight,
  paragraphStyle,
  runStyle,
  type DocumentLayout,
  type HFFrame,
  type HFParagraph,
  type HFRun,
  type LineNumbering,
  type PageInfo,
  type SectionLayout,
} from "@/lib/pm/format";
import {
  pageStride,
  stackHeight,
  type ChangeBalloon,
  type LinePlacement,
  type PageSetup,
} from "@/lib/pm/pagination";
import { displayColor } from "@/lib/pm/textColors";
import type { CommentThread } from "@/lib/pm/comments";
import { authorColor, REVISION_BAR_COLOR } from "@/lib/pm/track-changes";

type Props = {
  setup: PageSetup;
  pageCount: number;
  /** Section and displayed number of each page. */
  pages: PageInfo[];
  layout: DocumentLayout | null;
  /** Every section's layout (see `sectionLayouts`). */
  sections: SectionLayout[];
  lines: LinePlacement[];
  changedLines: LinePlacement[];
  balloons: ChangeBalloon[];
  /** Width of Word's markup area to the right of each page; 0 hides it. */
  markupWidth: number;
  simpleMarkup: boolean;
  threads: Record<string, CommentThread>;
  readOnly: boolean;
  onHeaderFooterEdit?: (page: PageInfo, part: "header" | "footer") => void;
  onBalloonAction?: (balloon: ChangeBalloon, action: BalloonAction) => void;
  onBalloonMenu?: (balloon: ChangeBalloon, x: number, y: number) => void;
};

export type BalloonAction = "select" | "accept" | "reject" | "reply" | "resolve" | "reopen" | "delete";

/** Word shows reviewer initials in balloons (e.g. "QC"), not always the full display name. */
function authorInitials(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) return "?";
  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length === 1) {
    const one = parts[0];
    if (one.length <= 4 && /^[A-Z][A-Z0-9.]*$/.test(one)) return one;
    return one.slice(0, 2).toUpperCase();
  }
  return parts.map((p) => p[0]).join("").slice(0, 4).toUpperCase();
}

const BALLOON_INSET = 14;
const BALLOON_GAP = 8;
const BALLOON_LINE = 15;
const BALLOON_CHAR = 6.2;
/** Padding, border, author line, and one line of body (11px/15px font). */
const BALLOON_CHROME = 12 + BALLOON_LINE;

function estimateBalloonHeight(label: string, charsPerLine: number): number {
  const bodyLines = label
    .split("\n")
    .reduce((n, part) => n + Math.max(1, Math.ceil(part.length / charsPerLine)), 0);
  return BALLOON_CHROME + BALLOON_LINE * bodyLines;
}

function runContent(run: HFRun, page: number, pageCount: number): string {
  if (run.kind === "tab") return "";
  if (run.kind === "field") return String(run.field === "PAGE" ? page : pageCount);
  return run.text;
}

function RunSpan({ run, page, pageCount }: { run: HFRun; page: number; pageCount: number }) {
  if (run.kind === "tab") return null;
  const style: CSSProperties = {
    ...runStyle(run),
    fontWeight: run.bold ? "bold" : undefined,
    fontStyle: run.italic ? "italic" : undefined,
    textDecoration: run.underline ? "underline" : undefined,
    color: run.color ? displayColor(run.color) : undefined,
  };
  return <span style={style}>{runContent(run, page, pageCount)}</span>;
}

/** A header/footer paragraph; text after each tab is placed at the paragraph's tab stops. */
function HFLine({
  paragraph,
  page,
  pageCount,
}: {
  paragraph: HFParagraph;
  page: number;
  pageCount: number;
}) {
  const segments: HFRun[][] = [[]];
  for (const run of paragraph.runs) {
    if (run.kind === "tab") segments.push([]);
    else segments[segments.length - 1].push(run);
  }
  const stops = paragraph.format.tabs ?? [];
  const render = (runs: HFRun[]): ReactNode =>
    runs.map((run, i) => <RunSpan key={i} run={run} page={page} pageCount={pageCount} />);
  return (
    <div className="studio-hf-paragraph" style={paragraphStyle(paragraph.format)}>
      {render(segments[0])}
      {segments.length === 1 && !paragraph.runs.length ? "\u00a0" : null}
      {segments.slice(1).map((runs, i) => {
        const stop = stops[i];
        if (!stop) return <span key={i}>{"\u2003"}{render(runs)}</span>;
        const transform =
          stop.align === "center" ? "translateX(-50%)" : stop.align === "right" ? "translateX(-100%)" : undefined;
        return (
          <span key={i} className="studio-hf-tabbed" style={{ left: stop.pos, transform }}>
            {render(runs)}
          </span>
        );
      })}
    </div>
  );
}

function HeaderFooter({
  paragraphs,
  page,
  pageCount,
  style,
  editable,
  onEdit,
}: {
  paragraphs: HFParagraph[] | undefined;
  page: number;
  pageCount: number;
  style: CSSProperties;
  editable?: boolean;
  onEdit?: () => void;
}) {
  const flow = paragraphs?.filter((p) => !p.floating) ?? [];
  if (!flow.length && !editable) return null;
  return (
    <div
      className={`studio-hf${editable ? " studio-hf--editable" : ""}`}
      style={style}
      onDoubleClick={editable ? onEdit : undefined}
      title={editable ? "Double-click to edit" : undefined}
    >
      {flow.map((p, i) => (
        <HFLine key={i} paragraph={p} page={page} pageCount={pageCount} />
      ))}
    </div>
  );
}

/** Absolute position of a Word frame on the page, from its anchors and alignment. */
function frameStyle(
  frame: HFFrame,
  setup: PageSetup,
  textTop: number,
): CSSProperties {
  const textWidth = setup.width - setup.marginLeft - setup.marginRight;
  const [hBase, hSpan] =
    frame.hAnchor === "page" ? [0, setup.width] : [setup.marginLeft, textWidth];
  const [vBase, vSpan] =
    frame.vAnchor === "page"
      ? [0, setup.height]
      : frame.vAnchor === "margin"
        ? [setup.marginTop, setup.height - setup.marginTop - setup.marginBottom]
        : [textTop, 0];
  const style: CSSProperties = { position: "absolute", width: frame.width, minHeight: frame.height };
  const tx: string[] = [];
  const x = frame.xAlign;
  if (x === "center") {
    style.left = hBase + hSpan / 2;
    tx.push("translateX(-50%)");
  } else if (x === "right" || x === "outside") {
    style.left = hBase + hSpan;
    tx.push("translateX(-100%)");
  } else {
    style.left = hBase + (x ? 0 : (frame.x ?? 0));
  }
  const y = frame.yAlign;
  if (y === "center") {
    style.top = vBase + vSpan / 2;
    tx.push("translateY(-50%)");
  } else if (y === "bottom" || y === "outside") {
    style.top = vBase + vSpan;
    tx.push("translateY(-100%)");
  } else {
    style.top = vBase + (y ? 0 : (frame.y ?? 0));
  }
  if (tx.length) style.transform = tx.join(" ");
  return style;
}

/** Header/footer paragraphs in Word frames, positioned on the page and taking no space. */
function Frames({
  paragraphs,
  setup,
  textTop,
  page,
  pageCount,
}: {
  paragraphs: HFParagraph[] | undefined;
  setup: PageSetup;
  textTop: number;
  page: number;
  pageCount: number;
}) {
  const framed = paragraphs?.filter((p) => p.frame) ?? [];
  return (
    <>
      {framed.map((p, i) => (
        <div key={i} className="studio-hf studio-hf-frame" style={frameStyle(p.frame!, setup, textTop)}>
          <HFLine paragraph={p} page={page} pageCount={pageCount} />
        </div>
      ))}
    </>
  );
}

type NumberedLine = { n: number; line: LinePlacement; ln: LineNumbering };

/** Line numbers for every page, following each section's countBy / restart rules. */
function lineNumbersByPage(sections: SectionLayout[], lines: LinePlacement[]) {
  const out = new Map<number, NumberedLine[]>();
  let counter: number | null = null;
  let currentPage = -1;
  let currentSection = -1;
  for (const line of lines) {
    const ln = sections[Math.min(line.section ?? 0, sections.length - 1)]?.lineNumbers;
    const newSection = (line.section ?? 0) !== currentSection;
    const newPage = line.page !== currentPage;
    currentSection = line.section ?? 0;
    currentPage = line.page;
    if (!ln) continue;
    if (
      counter === null ||
      (newSection ? ln.restart !== "continuous" : newPage && ln.restart === "newPage")
    ) {
      counter = ln.start;
    }
    if (counter % ln.countBy === 0) {
      const list = out.get(line.page) ?? [];
      list.push({ n: counter, line, ln });
      out.set(line.page, list);
    }
    counter += 1;
  }
  return out;
}

/** Consecutive changed lines on one page merged into Word's vertical change bars. */
function changeBarsFor(changedLines: LinePlacement[], page: number) {
  const bars: { top: number; bottom: number }[] = [];
  for (const line of changedLines) {
    if (line.page !== page) continue;
    const last = bars[bars.length - 1];
    if (last && line.top <= last.bottom + 2) last.bottom = Math.max(last.bottom, line.bottom);
    else bars.push({ top: line.top, bottom: line.bottom });
  }
  return bars;
}

/**
 * Balloons for one page, stacked in the markup column without overlap.
 * Heights are estimated conservatively (author + wrapped label); positions stay inside the page sheet.
 */
function balloonsFor(
  balloons: ChangeBalloon[],
  page: number,
  width: number,
  pageHeight: number,
) {
  const charsPerLine = Math.max(8, Math.floor((width - 16) / BALLOON_CHAR));
  const pageBottom = pageHeight - BALLOON_INSET;
  let floor = BALLOON_INSET;
  return balloons
    .filter((b) => b.page === page)
    .sort((a, b) => a.top - b.top || a.left - b.left)
    .map((b) => {
      const height = estimateBalloonHeight(b.label, charsPerLine);
      let top = Math.max(b.top - 4, floor);
      if (top + height > pageBottom) top = Math.max(floor, pageBottom - height);
      floor = top + height + BALLOON_GAP;
      return { balloon: b, top };
    });
}

function ActionButton({ label, glyph, onClick }: { label: string; glyph: string; onClick: () => void }) {
  return (
    <button
      type="button"
      title={label}
      aria-label={label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={(e) => {
        e.stopPropagation();
        onClick();
      }}
    >
      {glyph}
    </button>
  );
}

function Balloon({
  balloon,
  thread,
  readOnly,
  style,
  onAction,
  onMenu,
}: {
  balloon: ChangeBalloon;
  thread: CommentThread | undefined;
  readOnly: boolean;
  style: CSSProperties;
  onAction?: (balloon: ChangeBalloon, action: BalloonAction) => void;
  onMenu?: (balloon: ChangeBalloon, x: number, y: number) => void;
}) {
  const act = (action: BalloonAction) => onAction?.(balloon, action);
  const isComment = balloon.changeType === "comment";
  const resolved = isComment && !!thread?.resolved;
  return (
    <div
      className={`studio-balloon${isComment ? " studio-balloon--comment" : " studio-balloon--revision"}${resolved ? " studio-balloon--resolved" : ""}`}
      data-balloon-id={balloon.id}
      style={style}
      onContextMenu={(e) => {
        if (!onMenu) return;
        e.preventDefault();
        onMenu(balloon, e.clientX, e.clientY);
      }}
    >
      <button type="button" className="studio-balloon-body" onClick={() => act("select")}>
        <strong>{authorInitials(balloon.authorName || "Unknown")}</strong>
        {isComment && thread ? (
          <>
            <span>{thread.text}</span>
            {(thread.replies ?? []).map((reply) => (
              <span key={reply.id} className="studio-balloon-reply">
                <strong>{reply.authorName}</strong> {reply.text}
              </span>
            ))}
          </>
        ) : (
          <span>{balloon.label}</span>
        )}
      </button>
      {readOnly ? null : (
        <span className="studio-balloon-actions">
          {isComment ? (
            <>
              <ActionButton label="Reply" glyph={"\u21a9"} onClick={() => act("reply")} />
              <ActionButton
                label={resolved ? "Reopen" : "Resolve"}
                glyph={resolved ? "\u21ba" : "\u2713"}
                onClick={() => act(resolved ? "reopen" : "resolve")}
              />
              <ActionButton label="Delete comment" glyph={"\u2715"} onClick={() => act("delete")} />
            </>
          ) : (
            <>
              <ActionButton label="Accept" glyph={"\u2713"} onClick={() => act("accept")} />
              <ActionButton label="Reject" glyph={"\u2715"} onClick={() => act("reject")} />
            </>
          )}
        </span>
      )}
    </div>
  );
}

function MarkupPane({
  setup,
  balloons,
  threads,
  readOnly,
  page,
  width,
  onAction,
  onMenu,
}: {
  setup: PageSetup;
  balloons: ChangeBalloon[];
  threads: Record<string, CommentThread>;
  readOnly: boolean;
  page: number;
  width: number;
  onAction?: (balloon: ChangeBalloon, action: BalloonAction) => void;
  onMenu?: (balloon: ChangeBalloon, x: number, y: number) => void;
}) {
  const placed = balloonsFor(balloons, page, width - 2 * BALLOON_INSET, setup.height);
  const x = setup.width + BALLOON_INSET;
  return (
    <>
      <div className="studio-markup-pane" style={{ left: setup.width, width }} />
      {placed.length ? (
        <svg
          className="studio-markup-leaders"
          width={setup.width + width}
          height={setup.height}
        >
          {placed.map(({ balloon, top }) => {
            const y = balloon.bottom - 3;
            return (
              <polyline
                key={`${balloon.changeType}:${balloon.id}`}
                points={`${balloon.left},${y} ${setup.width + 4},${y} ${x},${top + 10}`}
                stroke={authorColor(balloon.authorId || "unknown")}
              />
            );
          })}
        </svg>
      ) : null}
      {placed.map(({ balloon, top }) => (
        <Balloon
          key={`${balloon.changeType}:${balloon.id}`}
          balloon={balloon}
          thread={balloon.changeType === "comment" ? threads[balloon.id] : undefined}
          readOnly={readOnly}
          style={{
            top,
            left: x,
            width: width - 2 * BALLOON_INSET,
            borderColor: authorColor(balloon.authorId || "unknown"),
            borderLeftWidth: 3,
          }}
          onAction={onAction}
          onMenu={onMenu}
        />
      ))}
    </>
  );
}

/** Page sheets, Word headers/footers, line numbers and markup drawn behind the editable flow. */
export function StudioPageBackgrounds({
  setup,
  pageCount,
  pages,
  layout,
  sections,
  lines,
  changedLines,
  balloons,
  markupWidth,
  simpleMarkup,
  threads,
  readOnly,
  onHeaderFooterEdit,
  onBalloonAction,
  onBalloonMenu,
}: Props) {
  const stride = pageStride(setup);
  const textWidth = setup.width - setup.marginLeft - setup.marginRight;
  const allSections = sections.length ? sections : layout ? [layout] : [];
  const anyLineNumbers = allSections.find((s) => s.lineNumbers)?.lineNumbers;
  const barX = anyLineNumbers
    ? setup.marginLeft - anyLineNumbers.distance - 30
    : setup.marginLeft - 14;
  const numbered = layout ? lineNumbersByPage(allSections, lines) : new Map<number, NumberedLine[]>();
  return (
    <div
      className="studio-page-stack"
      aria-hidden
      style={{ width: setup.width + markupWidth, height: stackHeight(setup, pageCount) }}
    >
      {Array.from({ length: pageCount }, (_, i) => {
        const info: PageInfo = pages[i] ?? { section: 0, first: i === 0, number: i + 1 };
        const section = allSections[Math.min(info.section, allSections.length - 1)] ?? layout ?? {};
        const evenAndOdd = !!layout?.evenAndOddHeaders;
        const header = headerFooterFor(section.headers, section, info, evenAndOdd);
        const footer = headerFooterFor(section.footers, section, info, evenAndOdd);
        return (
        <div
          key={i}
          className="studio-page-surface"
          style={{ top: i * stride, width: setup.width + markupWidth, height: setup.height }}
        >
          {changeBarsFor(changedLines, i).map((bar) => (
            <span
              key={bar.top}
              className="studio-change-bar"
              style={{
                left: Math.max(2, barX),
                top: bar.top,
                height: bar.bottom - bar.top,
                background: REVISION_BAR_COLOR,
              }}
            />
          ))}
          {markupWidth ? (
            <MarkupPane
              setup={setup}
              balloons={balloons}
              page={i}
              width={markupWidth}
              threads={threads}
              readOnly={readOnly}
              onAction={onBalloonAction}
              onMenu={onBalloonMenu}
            />
          ) : null}
          {layout ? (
            <>
              <HeaderFooter
                paragraphs={header}
                page={info.number}
                pageCount={pageCount}
                style={{ top: section.headerDistance ?? 48, left: setup.marginLeft, width: textWidth }}
                editable={!readOnly && !!onHeaderFooterEdit}
                onEdit={() => onHeaderFooterEdit?.(info, "header")}
              />
              <HeaderFooter
                paragraphs={footer}
                page={info.number}
                pageCount={pageCount}
                style={{ bottom: section.footerDistance ?? 48, left: setup.marginLeft, width: textWidth }}
                editable={!readOnly && !!onHeaderFooterEdit}
                onEdit={() => onHeaderFooterEdit?.(info, "footer")}
              />
              <Frames
                paragraphs={header}
                setup={setup}
                textTop={section.headerDistance ?? 48}
                page={info.number}
                pageCount={pageCount}
              />
              <Frames
                paragraphs={footer}
                setup={setup}
                textTop={setup.height - (section.footerDistance ?? 48) - headerFooterHeight(footer)}
                page={info.number}
                pageCount={pageCount}
              />
              {(numbered.get(i) ?? []).map(({ n, line, ln }) => (
                <span
                  key={`${n}:${line.top}`}
                  className="studio-line-number"
                  style={{
                    top: line.top,
                    height: line.bottom - line.top,
                    lineHeight: `${line.bottom - line.top}px`,
                    right: setup.width + markupWidth - setup.marginLeft + ln.distance,
                    ...runStyle(ln),
                  }}
                >
                  {n}
                </span>
              ))}
            </>
          ) : (
            <span
              className="studio-page-number"
              style={{ bottom: Math.max(0, setup.marginBottom / 2 - 8), right: markupWidth }}
            >
              Page {info.number}
            </span>
          )}
        </div>
        );
      })}
    </div>
  );
}
