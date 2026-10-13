"use client";

import { useState } from "react";
import { Dialog, inputClass } from "./dialogParts";

type Props = {
  title: string;
  initial?: string;
  /** Text the comment is about, shown for context. */
  quote?: string;
  onClose: () => void;
  onSubmit: (text: string) => void;
};

export function CommentDialog({ title, initial = "", quote, onClose, onSubmit }: Props) {
  const [text, setText] = useState(initial);
  const submit = () => {
    if (text.trim()) onSubmit(text.trim());
  };
  return (
    <Dialog title={title} width="w-[26rem]" okLabel="Post" onClose={onClose} onSubmit={submit}>
      {quote && (
        <blockquote className="border-l-2 border-zinc-600 pl-2 text-xs italic text-zinc-400">
          {quote.length > 160 ? `${quote.slice(0, 160)}\u2026` : quote}
        </blockquote>
      )}
      <textarea
        className={`${inputClass} h-28 resize-y`}
        aria-label="Comment"
        placeholder="Type your comment"
        value={text}
        autoFocus
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            submit();
          }
        }}
      />
    </Dialog>
  );
}
