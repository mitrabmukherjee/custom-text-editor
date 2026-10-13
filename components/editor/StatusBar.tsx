"use client";

import { useState } from "react";
import type { TextStats } from "@/lib/pm/stats";
import { Dialog } from "./dialogParts";
import { ZOOM_MAX, ZOOM_MIN, ZOOM_STEP } from "./useZoom";

type Props = {
  page: number;
  pageCount: number;
  stats: TextStats;
  /** Counts for a non-empty selection. */
  selection: TextStats | null;
  tracking: boolean;
  readOnly?: boolean;
  onToggleTracking: () => void;
  zoom: number;
  onZoom: (zoom: number) => void;
};

const n = (value: number) => value.toLocaleString();

const item = "rounded px-1.5 py-0.5 hover:bg-zinc-800 hover:text-zinc-200";

/** Word's status bar: page, word count, track changes and zoom. */
export function StatusBar({
  page,
  pageCount,
  stats,
  selection,
  tracking,
  readOnly,
  onToggleTracking,
  zoom,
  onZoom,
}: Props) {
  const [wordCount, setWordCount] = useState(false);
  const percent = Math.round(zoom * 100);
  return (
    <div className="studio-status-bar flex h-6 shrink-0 items-center gap-1 overflow-hidden whitespace-nowrap border-t border-zinc-800 bg-zinc-950 px-2 text-[11px] text-zinc-400">
      <span className="px-1.5">
        Page {n(page + 1)} of {n(pageCount)}
      </span>
      <button type="button" className={item} title="Word Count" onClick={() => setWordCount(true)}>
        {selection ? `${n(selection.words)} of ${n(stats.words)} words` : `${n(stats.words)} words`}
      </button>
      <span className="px-1.5">{n(selection?.characters ?? stats.characters)} characters</span>
      <button
        type="button"
        className={item}
        disabled={readOnly}
        title="Track Changes"
        onClick={onToggleTracking}
      >
        Track Changes: {tracking ? "On" : "Off"}
      </button>
      <span className="ml-auto flex items-center gap-1">
        <button
          type="button"
          className={`${item} px-2`}
          title="Zoom out"
          aria-label="Zoom out"
          disabled={zoom <= ZOOM_MIN}
          onClick={() => onZoom(zoom - ZOOM_STEP)}
        >
          {"\u2212"}
        </button>
        <input
          type="range"
          className="studio-zoom-slider w-28"
          min={ZOOM_MIN * 100}
          max={ZOOM_MAX * 100}
          step={ZOOM_STEP * 100}
          value={percent}
          aria-label="Zoom"
          onChange={(e) => onZoom(Number(e.target.value) / 100)}
        />
        <button
          type="button"
          className={`${item} px-2`}
          title="Zoom in"
          aria-label="Zoom in"
          disabled={zoom >= ZOOM_MAX}
          onClick={() => onZoom(zoom + ZOOM_STEP)}
        >
          +
        </button>
        <button type="button" className={`${item} w-12 text-right`} title="Reset zoom to 100%" onClick={() => onZoom(1)}>
          {percent}%
        </button>
      </span>
      {wordCount && (
        <Dialog
          title="Word Count"
          width="w-72"
          okLabel="Close"
          hideCancel
          onClose={() => setWordCount(false)}
          onSubmit={() => setWordCount(false)}
        >
          <dl className="grid grid-cols-[1fr_auto] gap-x-6 gap-y-1.5 text-sm text-zinc-300">
            {(
              [
                ["Pages", pageCount],
                ["Words", (selection ?? stats).words],
                ["Characters (no spaces)", (selection ?? stats).charactersNoSpaces],
                ["Characters (with spaces)", (selection ?? stats).characters],
                ["Paragraphs", (selection ?? stats).paragraphs],
              ] as const
            ).map(([label, value]) => (
              <div key={label} className="contents">
                <dt>{label}</dt>
                <dd className="text-right tabular-nums">{n(value)}</dd>
              </div>
            ))}
          </dl>
          {selection && <p className="text-xs text-zinc-500">Statistics for the selected text.</p>}
        </Dialog>
      )}
    </div>
  );
}
