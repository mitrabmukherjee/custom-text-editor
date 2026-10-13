"use client";

import { useState } from "react";
import { FONT_FAMILIES, FONT_SIZES } from "@/lib/pm/character";
import { FONT_COLORS } from "@/lib/pm/textColors";
import type { RunFormat } from "@/lib/pm/track-changes";
import { UNDERLINE_STYLES } from "./CharacterToolbar";
import { Check, Dialog, Field, Section, inputClass } from "./dialogParts";

export type FontDialogValue = Required<Pick<RunFormat, "bold" | "italic">> &
  Omit<RunFormat, "bold" | "italic" | "fontFamily" | "fontSize"> & {
    family: string;
    size: number;
  };

type Props = {
  format: FontDialogValue;
  onApply: (format: FontDialogValue) => void;
  onClose: () => void;
};

type FontStyle = "regular" | "italic" | "bold" | "boldItalic";

const STYLES: { value: FontStyle; label: string }[] = [
  { value: "regular", label: "Regular" },
  { value: "italic", label: "Italic" },
  { value: "bold", label: "Bold" },
  { value: "boldItalic", label: "Bold Italic" },
];

/** Word's Format → Font dialog (Ctrl+D). */
export function FontDialog({ format, onApply, onClose }: Props) {
  const [value, setValue] = useState<FontDialogValue>(format);
  const set = (patch: Partial<FontDialogValue>) => setValue((v) => ({ ...v, ...patch }));
  const style: FontStyle = value.bold
    ? value.italic
      ? "boldItalic"
      : "bold"
    : value.italic
      ? "italic"
      : "regular";
  const families = FONT_FAMILIES.includes(value.family)
    ? FONT_FAMILIES
    : [...FONT_FAMILIES, value.family].sort();

  return (
    <Dialog title="Font" onClose={onClose} onSubmit={() => onApply(value)}>
      <Section title="Font">
        <div className="grid grid-cols-[1fr_8rem_5rem] gap-3">
          <Field label="Font">
            <select
              className={inputClass}
              value={value.family}
              onChange={(e) => set({ family: e.target.value })}
            >
              {families.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Font style">
            <select
              className={inputClass}
              value={style}
              onChange={(e) => {
                const s = e.target.value as FontStyle;
                set({ bold: s === "bold" || s === "boldItalic", italic: s === "italic" || s === "boldItalic" });
              }}
            >
              {STYLES.map((s) => (
                <option key={s.value} value={s.value}>
                  {s.label}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Size">
            <input
              className={inputClass}
              list="te-font-dialog-sizes"
              value={value.size}
              onChange={(e) => {
                const n = parseFloat(e.target.value);
                if (Number.isFinite(n) && n > 0) set({ size: n });
              }}
            />
            <datalist id="te-font-dialog-sizes">
              {FONT_SIZES.map((s) => (
                <option key={s} value={s} />
              ))}
            </datalist>
          </Field>
        </div>
        <div className="mt-3 grid grid-cols-2 gap-3">
          <Field label="Font color">
            <select
              className={inputClass}
              value={value.color ?? ""}
              onChange={(e) => set({ color: e.target.value || undefined })}
            >
              <option value="">Automatic</option>
              {FONT_COLORS.map((c) => (
                <option key={c.hex} value={c.hex}>
                  {c.name}
                </option>
              ))}
              {value.color && !FONT_COLORS.some((c) => c.hex === value.color) && (
                <option value={value.color}>Custom ({value.color})</option>
              )}
            </select>
          </Field>
          <Field label="Underline style">
            <select
              className={inputClass}
              value={value.underline ?? ""}
              onChange={(e) =>
                set({ underline: (e.target.value || undefined) as RunFormat["underline"] })
              }
            >
              <option value="">(none)</option>
              {UNDERLINE_STYLES.map((u) => (
                <option key={u.style} value={u.style}>
                  {u.label}
                </option>
              ))}
            </select>
          </Field>
        </div>
      </Section>
      <Section title="Effects">
        <div className="grid grid-cols-2 gap-2">
          <Check
            label="Strikethrough"
            checked={value.strike === "single"}
            onChange={(on) => set({ strike: on ? "single" : undefined })}
          />
          <Check
            label="Double strikethrough"
            checked={value.strike === "double"}
            onChange={(on) => set({ strike: on ? "double" : undefined })}
          />
          <Check
            label="Superscript"
            checked={value.vertAlign === "superscript"}
            onChange={(on) => set({ vertAlign: on ? "superscript" : undefined })}
          />
          <Check
            label="Subscript"
            checked={value.vertAlign === "subscript"}
            onChange={(on) => set({ vertAlign: on ? "subscript" : undefined })}
          />
          <Check
            label="Small caps"
            checked={value.caps === "small"}
            onChange={(on) => set({ caps: on ? "small" : undefined })}
          />
          <Check
            label="All caps"
            checked={value.caps === "all"}
            onChange={(on) => set({ caps: on ? "all" : undefined })}
          />
        </div>
      </Section>
      <Section title="Preview">
        <div
          className="flex h-14 items-center justify-center overflow-hidden rounded border border-zinc-800 bg-zinc-950 text-zinc-100"
          style={{
            fontFamily: value.family,
            fontSize: `${Math.min(value.size, 36)}pt`,
            fontWeight: value.bold ? 700 : 400,
            fontStyle: value.italic ? "italic" : "normal",
            color: value.color ?? undefined,
            textDecoration: [
              value.underline ? "underline" : "",
              value.strike ? "line-through" : "",
            ].join(" ") || "none",
            textTransform: value.caps === "all" ? "uppercase" : undefined,
            fontVariant: value.caps === "small" ? "small-caps" : undefined,
            verticalAlign: value.vertAlign === "superscript" ? "super" : value.vertAlign === "subscript" ? "sub" : undefined,
          }}
        >
          {value.family}
        </div>
      </Section>
    </Dialog>
  );
}
