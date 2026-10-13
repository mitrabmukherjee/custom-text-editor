"use client";

import { useEffect, useRef, useState, type KeyboardEvent } from "react";
import type { EditorController } from "@/lib/pm/controller";
import { Check, inputClass } from "./dialogParts";

export type FindMode = "find" | "replace";

type Props = {
  controller: EditorController;
  mode: FindMode;
  readOnly?: boolean;
  onModeChange: (mode: FindMode) => void;
  onClose: () => void;
};

/** Short single-paragraph selection to search for, like Word; otherwise the last search. */
function initialText(controller: EditorController) {
  const { doc, selection } = controller.state;
  if (!selection.empty) {
    const text = doc.textBetween(selection.from, selection.to, "\n");
    if (text.length <= 255 && !text.includes("\n")) return text;
  }
  return controller.getFind().query.text;
}

const button =
  "rounded border border-zinc-700 px-2.5 py-1 text-xs text-zinc-200 hover:bg-zinc-800 disabled:opacity-40";

/** Word's Find and Replace: a floating panel that leaves the document editable. */
export function FindReplaceDialog({ controller, mode, readOnly, onModeChange, onClose }: Props) {
  const [text, setText] = useState(() => initialText(controller));
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(() => !!controller.getFind().query.matchCase);
  const [wholeWord, setWholeWord] = useState(() => !!controller.getFind().query.wholeWord);
  const [message, setMessage] = useState<string | null>(null);
  const findRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    controller.setFind({ text, matchCase, wholeWord });
  }, [controller, text, matchCase, wholeWord]);

  useEffect(() => () => controller.setFind({ text: "" }), [controller]);

  useEffect(() => {
    findRef.current?.focus();
    findRef.current?.select();
  }, [mode]);

  const { count, current } = controller.getFind();
  const status = message ?? (!text ? "" : count ? `${current + 1} of ${count}` : "No results");
  const replacing = mode === "replace" && !readOnly;

  const next = (direction: 1 | -1) => {
    setMessage(null);
    controller.findNext(direction);
  };
  const replace = () => {
    setMessage(null);
    controller.replace(replacement);
  };
  const replaceAll = () => {
    const n = controller.replaceAll(replacement);
    setMessage(`Replaced ${n} occurrence${n === 1 ? "" : "s"}`);
  };
  const onKey = (e: KeyboardEvent, enter: () => void) => {
    if (e.key === "Escape") {
      e.preventDefault();
      onClose();
    } else if (e.key === "Enter") {
      e.preventDefault();
      enter();
    }
  };

  return (
    <div
      role="dialog"
      aria-label="Find and replace"
      className="absolute right-4 top-full z-40 mt-2 w-80 space-y-2 rounded-lg border border-zinc-700 bg-zinc-900 p-3 shadow-2xl"
    >
      <div className="flex items-center gap-1 text-xs">
        {(["find", "replace"] as const).map((m) => (
          <button
            key={m}
            type="button"
            disabled={m === "replace" && readOnly}
            className={`rounded px-2 py-0.5 disabled:opacity-40 ${
              mode === m ? "bg-zinc-700 text-zinc-100" : "text-zinc-400 hover:bg-zinc-800"
            }`}
            onClick={() => onModeChange(m)}
          >
            {m === "find" ? "Find" : "Replace"}
          </button>
        ))}
        <span className="flex-1" />
        <button
          type="button"
          className="rounded px-1.5 text-zinc-400 hover:bg-zinc-800 hover:text-zinc-100"
          title="Close (Esc)"
          aria-label="Close"
          onClick={onClose}
        >
          {"\u2715"}
        </button>
      </div>
      <div className="flex items-center gap-2">
        <input
          ref={findRef}
          className={inputClass}
          placeholder="Find what"
          aria-label="Find what"
          value={text}
          onChange={(e) => {
            setMessage(null);
            setText(e.target.value);
          }}
          onKeyDown={(e) => onKey(e, () => next(e.shiftKey ? -1 : 1))}
        />
        <span className="w-20 shrink-0 text-right text-[11px] text-zinc-500" aria-live="polite">
          {status}
        </span>
      </div>
      {replacing && (
        <input
          className={inputClass}
          placeholder="Replace with"
          aria-label="Replace with"
          value={replacement}
          onChange={(e) => setReplacement(e.target.value)}
          onKeyDown={(e) => onKey(e, replace)}
        />
      )}
      <div className="flex gap-4">
        <Check label="Match case" checked={matchCase} onChange={setMatchCase} />
        <Check label="Whole words only" checked={wholeWord} onChange={setWholeWord} />
      </div>
      <div className="flex flex-wrap justify-end gap-1.5">
        {replacing && (
          <>
            <button type="button" className={button} disabled={!count} onClick={replace}>
              Replace
            </button>
            <button type="button" className={button} disabled={!count} onClick={replaceAll}>
              Replace All
            </button>
          </>
        )}
        <button type="button" className={button} disabled={!count} onClick={() => next(-1)}>
          {"\u2191"} Previous
        </button>
        <button type="button" className={button} disabled={!count} onClick={() => next(1)}>
          {"\u2193"} Next
        </button>
      </div>
    </div>
  );
}
