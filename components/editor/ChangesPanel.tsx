"use client";

import { authorColor, type ChangeType, type TrackedChange } from "@/lib/pm/track-changes";

const VERB: Record<ChangeType, string> = {
  insertion: "formatted",
  deletion: "formatted",
  format: "formatted",
  paragraph: "formatted",
};

const TEXT_CLASS: Record<ChangeType, string> = {
  insertion: "text-emerald-200",
  deletion: "text-rose-200 line-through",
  format: "text-zinc-200",
  paragraph: "text-zinc-200",
};

type Props = {
  changes: TrackedChange[];
  onAccept: (id: string) => void;
  onReject: (id: string) => void;
  onSelect: (id: string) => void;
  onAcceptAll: () => void;
  onRejectAll: () => void;
  readOnly?: boolean;
};

function smallBtn(tone: "accept" | "reject" | "neutral") {
  const tones = {
    accept: "border-emerald-800 text-emerald-300 hover:bg-emerald-950",
    reject: "border-rose-900 text-rose-300 hover:bg-rose-950",
    neutral: "border-zinc-700 text-zinc-300 hover:bg-zinc-800",
  };
  return `rounded border px-2 py-0.5 text-[11px] font-medium disabled:opacity-40 ${tones[tone]}`;
}

function snippet(text: string) {
  const t = text.replace(/\s+/g, " ").trim();
  return t.length > 80 ? `${t.slice(0, 80)}\u2026` : t || "(whitespace)";
}

/** List of tracked changes with accept/reject controls. */
export function ChangesPanel({
  changes,
  onAccept,
  onReject,
  onSelect,
  onAcceptAll,
  onRejectAll,
  readOnly,
}: Props) {
  return (
    <div>
      <div className="mb-2 flex items-center gap-2">
        <span className="text-xs text-zinc-400">
          {changes.length} change{changes.length === 1 ? "" : "s"}
        </span>
        <span className="ml-auto flex gap-1">
          <button
            type="button"
            className={smallBtn("accept")}
            disabled={readOnly || !changes.length}
            onClick={onAcceptAll}
          >
            Accept all
          </button>
          <button
            type="button"
            className={smallBtn("reject")}
            disabled={readOnly || !changes.length}
            onClick={onRejectAll}
          >
            Reject all
          </button>
        </span>
      </div>
      {changes.length === 0 ? (
        <p className="text-xs text-zinc-600">No tracked changes.</p>
      ) : (
        <ul className="space-y-2">
          {changes.map((change) => (
            <li
              key={`${change.type}:${change.id}`}
              className="rounded-md border border-zinc-800 bg-zinc-900 p-2"
            >
              <button
                type="button"
                className="block w-full text-left"
                onClick={() => onSelect(change.id)}
              >
                <span className="flex items-center gap-1.5 text-[11px] text-zinc-400">
                  <span
                    className="inline-block h-2 w-2 rounded-full"
                    style={{ background: authorColor(change.authorId || "unknown") }}
                  />
                  <span className="font-medium text-zinc-200">
                    {change.authorName || "Unknown"}
                  </span>
                  <span>{VERB[change.type]}</span>
                  {change.date && (
                    <span className="ml-auto text-zinc-600">
                      {new Date(change.date).toLocaleTimeString()}
                    </span>
                  )}
                </span>
                <span className={`mt-1 block text-xs ${TEXT_CLASS[change.type]}`}>
                  {snippet(change.text)}
                </span>
                {change.description && (
                  <span className="mt-0.5 block text-[11px] text-zinc-400">
                    Formatted: {change.description}
                  </span>
                )}
              </button>
              <span className="mt-2 flex gap-1">
                <button
                  type="button"
                  className={smallBtn("accept")}
                  disabled={readOnly}
                  onClick={() => onAccept(change.id)}
                >
                  Accept
                </button>
                <button
                  type="button"
                  className={smallBtn("reject")}
                  disabled={readOnly}
                  onClick={() => onReject(change.id)}
                >
                  Reject
                </button>
              </span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
