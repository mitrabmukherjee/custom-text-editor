"use client";

import { useMemo, useState } from "react";
import type { EditorController } from "@/lib/pm/controller";
import { outlineIndexAt } from "@/lib/pm/outline";
import type { LinePlacement, PageSetup } from "@/lib/pm/pagination";

const THUMB_WIDTH = 112;

type Props = {
  controller: EditorController;
  setup: PageSetup;
  pageCount: number;
  lines: LinePlacement[];
  currentPage: number;
  onHeading: (pos: number) => void;
  onPage: (page: number) => void;
  onClose: () => void;
};

/** Word's Navigation pane: the headings outline and page thumbnails. */
export function NavigationPane({
  controller,
  setup,
  pageCount,
  lines,
  currentPage,
  onHeading,
  onPage,
  onClose,
}: Props) {
  const [tab, setTab] = useState<"headings" | "pages">("headings");
  const outline = controller.getOutline();
  const active = outlineIndexAt(outline, controller.state.selection.from);
  const minLevel = Math.min(...outline.map((e) => e.level));

  const scale = THUMB_WIDTH / setup.width;
  const linesByPage = useMemo(() => {
    const out = new Map<number, LinePlacement[]>();
    for (const line of lines) {
      const list = out.get(line.page) ?? [];
      list.push(line);
      out.set(line.page, list);
    }
    return out;
  }, [lines]);

  return (
    <aside className="studio-nav-pane flex w-60 shrink-0 flex-col border-r border-zinc-800 bg-zinc-950 text-sm">
      <div className="flex items-center justify-between px-3 pb-1 pt-2">
        <h2 className="text-xs font-semibold uppercase tracking-wide text-zinc-400">Navigation</h2>
        <button
          type="button"
          className="rounded px-1.5 text-zinc-500 hover:bg-zinc-800 hover:text-zinc-100"
          title="Close"
          aria-label="Close navigation pane"
          onClick={onClose}
        >
          {"\u2715"}
        </button>
      </div>
      <div className="flex gap-1 border-b border-zinc-800 px-3 text-xs" role="tablist">
        {(["headings", "pages"] as const).map((t) => (
          <button
            key={t}
            type="button"
            role="tab"
            aria-selected={tab === t}
            className={`-mb-px border-b-2 px-2 py-1.5 capitalize ${
              tab === t ? "border-sky-500 text-zinc-100" : "border-transparent text-zinc-500 hover:text-zinc-300"
            }`}
            onClick={() => setTab(t)}
          >
            {t}
          </button>
        ))}
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto p-2">
        {tab === "headings" ? (
          outline.length ? (
            <ul className="space-y-px">
              {outline.map((entry, i) => (
                <li key={entry.pos}>
                  <button
                    type="button"
                    className={`w-full truncate rounded px-2 py-1 text-left ${
                      i === active ? "bg-sky-900/50 text-sky-100" : "text-zinc-300 hover:bg-zinc-800"
                    } ${entry.level === minLevel ? "font-medium" : ""}`}
                    style={{ paddingLeft: 8 + (entry.level - minLevel) * 14 }}
                    title={entry.text}
                    onClick={() => onHeading(entry.pos)}
                  >
                    {entry.text}
                  </button>
                </li>
              ))}
            </ul>
          ) : (
            <p className="px-2 py-3 text-xs leading-relaxed text-zinc-500">
              This document has no headings. Apply a Heading style to text to list it here.
            </p>
          )
        ) : (
          <ul className="flex flex-col items-center gap-3 py-1">
            {Array.from({ length: pageCount }, (_, page) => (
              <li key={page} className="flex flex-col items-center gap-1">
                <button
                  type="button"
                  className={`studio-nav-thumb relative overflow-hidden ${
                    page === currentPage ? "ring-2 ring-sky-500" : "ring-1 ring-zinc-700 hover:ring-zinc-500"
                  }`}
                  style={{ width: THUMB_WIDTH, height: setup.height * scale }}
                  aria-label={`Page ${page + 1}`}
                  onClick={() => onPage(page)}
                >
                  {(linesByPage.get(page) ?? []).map((line, i) => (
                    <span
                      key={i}
                      className="absolute bg-zinc-500/70"
                      style={{
                        top: line.top * scale,
                        left: setup.marginLeft * scale,
                        width: (setup.width - setup.marginLeft - setup.marginRight) * scale,
                        height: Math.max(1, (line.bottom - line.top) * scale * 0.45),
                      }}
                    />
                  ))}
                </button>
                <span className="text-[11px] text-zinc-500">{page + 1}</span>
              </li>
            ))}
          </ul>
        )}
      </div>
    </aside>
  );
}
