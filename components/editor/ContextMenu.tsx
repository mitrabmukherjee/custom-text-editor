"use client";

import { useEffect, useRef } from "react";
import { createPortal } from "react-dom";
import type { MenuItem } from "./toolbarParts";

export type ContextMenuState = { x: number; y: number; items: MenuItem[] };

/** Right-click menu at the pointer, rendered on `document.body` so page zoom doesn't move it. */
export function ContextMenu({ x, y, items, onClose }: ContextMenuState & { onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey);
    window.addEventListener("scroll", onClose, true);
    window.addEventListener("blur", onClose);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      document.removeEventListener("keydown", onKey);
      window.removeEventListener("scroll", onClose, true);
      window.removeEventListener("blur", onClose);
    };
  }, [onClose]);

  return createPortal(
    <div
      ref={ref}
      role="menu"
      className="fixed z-[200] min-w-[11rem] rounded-md border border-zinc-700 bg-zinc-900 py-1 shadow-xl"
      style={{ left: Math.min(x, window.innerWidth - 200), top: Math.min(y, window.innerHeight - items.length * 32 - 8) }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {items.map((item) => (
        <button
          key={item.label}
          type="button"
          role="menuitem"
          disabled={item.disabled}
          className="flex w-full items-center gap-2 whitespace-nowrap px-3 py-1.5 text-left text-sm text-zinc-200 hover:bg-zinc-800 disabled:opacity-40"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            onClose();
            item.run();
          }}
        >
          <span className="flex-1">{item.label}</span>
          {item.shortcut && <span className="text-[11px] text-zinc-500">{item.shortcut}</span>}
        </button>
      ))}
    </div>,
    document.body,
  );
}
