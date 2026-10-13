"use client";

import type { ReactNode } from "react";

export type MenuItem = {
  label: string;
  run: () => void;
  disabled?: boolean;
  checked?: boolean;
  shortcut?: string;
};

/** Small line-art icon: rows of [x, width] bars on a 16×16 grid. */
export function Bars({ rows }: { rows: [number, number][] }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden>
      {rows.map(([x, w], i) => (
        <rect key={i} x={x} y={2 + i * 3} width={w} height="1.5" fill="currentColor" />
      ))}
    </svg>
  );
}

export function MenuList({
  items,
  onDone,
  align = "left",
}: {
  items: MenuItem[];
  onDone: () => void;
  align?: "left" | "right";
}) {
  return (
    <div
      className={`absolute top-full z-50 mt-0.5 min-w-[13rem] rounded-md border border-zinc-700 bg-zinc-900 py-1 shadow-xl ${
        align === "left" ? "left-0" : "right-0"
      }`}
      role="menu"
    >
      {items.map((item) => (
        <button
          key={item.label}
          type="button"
          role={item.checked === undefined ? "menuitem" : "menuitemcheckbox"}
          aria-checked={item.checked}
          disabled={item.disabled}
          className="flex w-full items-center gap-2 whitespace-nowrap px-3 py-1.5 text-left text-sm text-zinc-200 hover:bg-zinc-800 disabled:opacity-40"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            if (item.disabled) return;
            item.run();
            onDone();
          }}
        >
          <span className="w-3 text-xs">{item.checked ? "\u2713" : ""}</span>
          <span className="flex-1">{item.label}</span>
          {item.shortcut && <span className="text-[11px] text-zinc-500">{item.shortcut}</span>}
        </button>
      ))}
    </div>
  );
}

export function ToolButton({
  title,
  active = false,
  disabled = false,
  onClick,
  children,
  className = "",
}: {
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
  className?: string;
}) {
  return (
    <button
      type="button"
      className={`${iconBtn(active, disabled)} ${className}`}
      title={title}
      aria-label={title}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

export function iconBtn(active = false, disabled = false) {
  return `rounded p-1.5 text-zinc-400 ${
    disabled
      ? "cursor-not-allowed opacity-40"
      : active
        ? "bg-zinc-700 text-zinc-100"
        : "hover:bg-zinc-800 hover:text-zinc-100"
  }`;
}

export function Separator() {
  return <span className="mx-1 h-5 w-px bg-zinc-800" aria-hidden />;
}
