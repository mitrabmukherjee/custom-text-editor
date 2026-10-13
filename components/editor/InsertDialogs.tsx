"use client";

import { useState } from "react";
import { Dialog, Field, NumberInput, inputClass } from "./dialogParts";

export function InsertTableDialog({
  onClose,
  onSubmit,
}: {
  onClose: () => void;
  onSubmit: (rows: number, cols: number) => void;
}) {
  const [cols, setCols] = useState(3);
  const [rows, setRows] = useState(2);
  return (
    <Dialog title="Insert Table" width="w-[18rem]" onClose={onClose} onSubmit={() => onSubmit(rows, cols)}>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Number of columns">
          <NumberInput value={cols} step={1} min={1} max={63} suffix="" onChange={(v) => setCols(Math.min(63, Math.max(1, v)))} />
        </Field>
        <Field label="Number of rows">
          <NumberInput value={rows} step={1} min={1} max={500} suffix="" onChange={(v) => setRows(Math.min(500, Math.max(1, v)))} />
        </Field>
      </div>
    </Dialog>
  );
}

export function LinkDialog({
  initialText,
  initialHref,
  onClose,
  onSubmit,
  onRemove,
}: {
  initialText: string;
  initialHref: string;
  onClose: () => void;
  onSubmit: (href: string, text: string) => void;
  /** Shown when editing an existing link. */
  onRemove?: () => void;
}) {
  const [text, setText] = useState(initialText);
  const [href, setHref] = useState(initialHref);
  return (
    <Dialog
      title={onRemove ? "Edit Hyperlink" : "Insert Hyperlink"}
      width="w-[26rem]"
      onClose={onClose}
      onSubmit={() => href.trim() && onSubmit(href.trim(), text)}
      extraButtons={
        onRemove && (
          <button
            type="button"
            className="rounded border border-zinc-700 px-3 py-1 text-sm text-zinc-300 hover:bg-zinc-800"
            onClick={onRemove}
          >
            Remove Link
          </button>
        )
      }
    >
      <Field label="Text to display">
        <input className={inputClass} value={text} onChange={(e) => setText(e.target.value)} />
      </Field>
      <Field label="Address">
        <input
          className={inputClass}
          value={href}
          autoFocus
          placeholder="https://"
          onChange={(e) => setHref(e.target.value)}
        />
      </Field>
    </Dialog>
  );
}
