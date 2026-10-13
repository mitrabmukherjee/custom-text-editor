"use client";

import { useState } from "react";
import { Dialog, Field, inputClass } from "./dialogParts";

type Props = {
  title: string;
  initial: string;
  onClose: () => void;
  onSubmit: (text: string) => void;
};

export function HeaderFooterDialog({ title, initial, onClose, onSubmit }: Props) {
  const [text, setText] = useState(initial);
  return (
    <Dialog title={title} width="w-[26rem]" onClose={onClose} onSubmit={() => onSubmit(text)}>
      <p className="text-xs text-zinc-400">
        Use {"{PAGE}"}, {"{NUMPAGES}"}, or {"{SECTIONPAGES}"} for fields. One line per paragraph.
      </p>
      <Field label="Content">
        <textarea
          className={`${inputClass} h-28 resize-y font-mono text-xs`}
          value={text}
          autoFocus
          onChange={(e) => setText(e.target.value)}
        />
      </Field>
    </Dialog>
  );
}
