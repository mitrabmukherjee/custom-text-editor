"use client";

import { useEffect, type ReactNode } from "react";

export const inputClass =
  "w-full rounded border border-zinc-700 bg-zinc-950 px-2 py-1 text-sm text-zinc-100 focus:border-sky-600 focus:outline-none";

export function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <label className="flex flex-col gap-1 text-xs text-zinc-400">
      {label}
      {children}
    </label>
  );
}

export function NumberInput({
  value,
  onChange,
  step,
  min,
  max,
  suffix,
  disabled,
}: {
  value: number;
  onChange: (v: number) => void;
  step: number;
  min?: number;
  max?: number;
  suffix: string;
  disabled?: boolean;
}) {
  return (
    <span className="flex items-center gap-1">
      <input
        type="number"
        className={inputClass}
        value={value}
        step={step}
        min={min}
        max={max}
        disabled={disabled}
        onChange={(e) => onChange(Number(e.target.value) || 0)}
      />
      <span className="w-5 text-[11px] text-zinc-500">{suffix}</span>
    </span>
  );
}

export function Check({
  label,
  checked,
  onChange,
  disabled,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <label className={`flex items-center gap-2 text-sm text-zinc-200 ${disabled ? "opacity-40" : ""}`}>
      <input
        type="checkbox"
        checked={checked}
        disabled={disabled}
        onChange={(e) => onChange(e.target.checked)}
      />
      {label}
    </label>
  );
}

export function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <fieldset className="border-t border-zinc-800 pt-3">
      <legend className="pr-2 text-xs font-semibold uppercase tracking-wide text-zinc-500">
        {title}
      </legend>
      {children}
    </fieldset>
  );
}

/** Modal dialog shell with OK / Cancel; Escape and clicking outside cancel. */
export function Dialog({
  title,
  onClose,
  onSubmit,
  width = "w-[30rem]",
  okLabel = "OK",
  extraButtons,
  hideCancel,
  children,
}: {
  title: string;
  onClose: () => void;
  onSubmit: () => void;
  width?: string;
  okLabel?: string;
  extraButtons?: ReactNode;
  hideCancel?: boolean;
  children: ReactNode;
}) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  return (
    <div
      className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        role="dialog"
        aria-label={title}
        className={`max-h-[92vh] ${width} max-w-[95vw] space-y-4 overflow-y-auto rounded-lg border border-zinc-700 bg-zinc-900 p-5 shadow-2xl`}
        onSubmit={(e) => {
          e.preventDefault();
          onSubmit();
        }}
      >
        <h2 className="text-sm font-semibold text-zinc-100">{title}</h2>
        {children}
        <div className="flex items-center justify-end gap-2 pt-1">
          {extraButtons}
          <span className="flex-1" />
          {!hideCancel && (
            <button
              type="button"
              className="rounded border border-zinc-700 px-3 py-1 text-sm text-zinc-300 hover:bg-zinc-800"
              onClick={onClose}
            >
              Cancel
            </button>
          )}
          <button
            type="submit"
            className="rounded bg-sky-700 px-3 py-1 text-sm font-medium text-white hover:bg-sky-600"
          >
            {okLabel}
          </button>
        </div>
      </form>
    </div>
  );
}
