"use client";

import { useState } from "react";
import {
  FONT_FAMILIES,
  FONT_SIZES,
  clearFormatting,
  growFont,
  setFontFamily,
  setFontSize,
  setHighlight,
  setMark,
  setTextColor,
  toggleStrike,
  toggleSubscript,
  toggleSuperscript,
  toggleUnderline,
} from "@/lib/pm/character";
import type { EditorController } from "@/lib/pm/controller";
import { schema } from "@/lib/pm/schema";
import { FONT_COLORS, HIGHLIGHT_COLORS } from "@/lib/pm/textColors";
import { MenuList, ToolButton, type MenuItem } from "./toolbarParts";

type Props = {
  controller: EditorController;
  readOnly?: boolean;
  openMenu: string | null;
  setOpenMenu: (update: (current: string | null) => string | null) => void;
};

export const UNDERLINE_STYLES: { style: string; label: string; shortcut?: string }[] = [
  { style: "single", label: "Single", shortcut: "Ctrl+U" },
  { style: "double", label: "Double", shortcut: "Ctrl+Shift+D" },
  { style: "thick", label: "Thick" },
  { style: "dotted", label: "Dotted" },
  { style: "dash", label: "Dashed" },
  { style: "dotDash", label: "Dot-dash" },
  { style: "wave", label: "Wave" },
  { style: "words", label: "Words only", shortcut: "Ctrl+Shift+W" },
];

/** Word's highlight palette order. */
const HIGHLIGHT_ORDER = [
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
];

const titleCase = (name: string) =>
  name.replace(/([A-Z])/g, " $1").replace(/^./, (c) => c.toUpperCase());

function Swatches({
  colors,
  noneLabel,
  onPick,
}: {
  colors: { name: string; hex: string }[];
  noneLabel: string;
  onPick: (hex: string | null) => void;
}) {
  return (
    <div
      className="absolute left-0 top-full z-50 mt-0.5 w-[11.5rem] rounded-md border border-zinc-700 bg-zinc-900 p-2 shadow-xl"
      role="menu"
    >
      <button
        type="button"
        className="mb-2 w-full rounded px-2 py-1 text-left text-sm text-zinc-200 hover:bg-zinc-800"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => onPick(null)}
      >
        {noneLabel}
      </button>
      <div className="grid grid-cols-5 gap-1.5">
        {colors.map((c) => (
          <button
            key={c.hex}
            type="button"
            title={c.name}
            aria-label={c.name}
            className="h-6 w-full rounded-sm border border-zinc-600 hover:ring-2 hover:ring-zinc-300"
            style={{ background: c.hex }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onPick(c.hex)}
          />
        ))}
      </div>
    </div>
  );
}

function FontSizeBox({
  size,
  disabled,
  onApply,
}: {
  size: number;
  disabled?: boolean;
  onApply: (size: number) => void;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const apply = () => {
    if (draft === null) return;
    const n = parseFloat(draft);
    setDraft(null);
    if (Number.isFinite(n) && n > 0) onApply(n);
  };
  return (
    <>
      <input
        className="h-7 w-12 rounded border border-zinc-700 bg-zinc-900 px-1.5 text-xs text-zinc-200 outline-none focus:border-zinc-500"
        title="Font size"
        aria-label="Font size"
        list="te-font-sizes"
        disabled={disabled}
        value={draft ?? String(size)}
        onChange={(e) => {
          const value = e.target.value;
          setDraft(value);
          // Picking from the datalist: apply straight away.
          const inputType = (e.nativeEvent as InputEvent).inputType;
          if (FONT_SIZES.includes(parseFloat(value)) && (!inputType || inputType === "insertReplacementText")) {
            setDraft(null);
            onApply(parseFloat(value));
          }
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            apply();
          } else if (e.key === "Escape") {
            setDraft(null);
          }
        }}
        onBlur={apply}
      />
      <datalist id="te-font-sizes">
        {FONT_SIZES.map((s) => (
          <option key={s} value={s} />
        ))}
      </datalist>
    </>
  );
}

/** Word's Home ▸ Font group: font, size, grow/shrink, clear, B I U S x₂ x², colour, highlight. */
export function CharacterToolbar({ controller, readOnly, openMenu, setOpenMenu }: Props) {
  const [lastColor, setLastColor] = useState("#ff0000");
  const [lastHighlight, setLastHighlight] = useState(HIGHLIGHT_COLORS.yellow);
  const format = controller.getCharacterFormat();
  const families = FONT_FAMILIES.includes(format.family)
    ? FONT_FAMILIES
    : [...FONT_FAMILIES, format.family].sort();
  const toggle = (name: string) => setOpenMenu((m) => (m === name ? null : name));
  const close = () => setOpenMenu(() => null);

  const underlineItems: MenuItem[] = [
    ...UNDERLINE_STYLES.map(({ style, label, shortcut }) => ({
      label,
      shortcut,
      checked: format.underline === style,
      run: () => controller.exec(setMark(schema.marks.underline, { style })),
    })),
    { label: "No underline", run: () => controller.exec(setMark(schema.marks.underline, null)) },
  ];

  return (
    <>
      <select
        className="h-7 w-36 rounded border border-zinc-700 bg-zinc-900 px-1 text-xs text-zinc-200 outline-none focus:border-zinc-500"
        title="Font"
        aria-label="Font"
        disabled={readOnly}
        value={format.family}
        onChange={(e) => controller.exec(setFontFamily(e.target.value))}
      >
        {families.map((f) => (
          <option key={f} value={f} style={{ fontFamily: f }}>
            {f}
          </option>
        ))}
      </select>
      <FontSizeBox
        size={format.size}
        disabled={readOnly}
        onApply={(size) => controller.exec(setFontSize(size))}
      />
      <ToolButton
        title="Increase font size (Ctrl+Shift+>)"
        disabled={readOnly}
        onClick={() => controller.exec(growFont(1))}
      >
        <span className="text-xs font-semibold leading-4">A{"\u02c4"}</span>
      </ToolButton>
      <ToolButton
        title="Decrease font size (Ctrl+Shift+<)"
        disabled={readOnly}
        onClick={() => controller.exec(growFont(-1))}
      >
        <span className="text-[10px] font-semibold leading-4">A{"\u02c5"}</span>
      </ToolButton>
      <ToolButton
        title="Clear formatting (Ctrl+Space)"
        disabled={readOnly}
        onClick={() => controller.exec(clearFormatting)}
      >
        <span className="relative text-xs leading-4">
          A<span className="absolute -bottom-0.5 -right-1.5 text-[9px] text-rose-400">{"\u2715"}</span>
        </span>
      </ToolButton>
      <span className="mx-1 h-5 w-px bg-zinc-800" aria-hidden />
      <ToolButton
        title="Bold (Ctrl+B)"
        active={!!format.bold}
        disabled={readOnly}
        onClick={() => controller.toggleMark(schema.marks.strong)}
      >
        <strong className="text-xs font-bold leading-4">B</strong>
      </ToolButton>
      <ToolButton
        title="Italic (Ctrl+I)"
        active={!!format.italic}
        disabled={readOnly}
        onClick={() => controller.toggleMark(schema.marks.em)}
      >
        <em className="font-serif text-xs italic leading-4">I</em>
      </ToolButton>
      <div className="relative flex items-center">
        <ToolButton
          title="Underline (Ctrl+U)"
          active={!!format.underline}
          disabled={readOnly}
          onClick={() => controller.exec(toggleUnderline(format.underline ?? "single"))}
        >
          <span className="text-xs leading-4 underline">U</span>
        </ToolButton>
        <ToolButton
          title="Underline style"
          disabled={readOnly}
          className="!px-0.5"
          onClick={() => toggle("underline")}
        >
          <span className="text-[9px]">{"\u25be"}</span>
        </ToolButton>
        {openMenu === "underline" && <MenuList items={underlineItems} onDone={close} />}
      </div>
      <ToolButton
        title="Strikethrough"
        active={!!format.strike}
        disabled={readOnly}
        onClick={() => controller.exec(toggleStrike())}
      >
        <span className="text-xs leading-4 line-through">ab</span>
      </ToolButton>
      <ToolButton
        title="Subscript (Ctrl+=)"
        active={format.vertAlign === "subscript"}
        disabled={readOnly}
        onClick={() => controller.exec(toggleSubscript)}
      >
        <span className="text-xs leading-4">
          x<sub className="text-[8px]">2</sub>
        </span>
      </ToolButton>
      <ToolButton
        title="Superscript (Ctrl+Shift+=)"
        active={format.vertAlign === "superscript"}
        disabled={readOnly}
        onClick={() => controller.exec(toggleSuperscript)}
      >
        <span className="text-xs leading-4">
          x<sup className="text-[8px]">2</sup>
        </span>
      </ToolButton>
      <div className="relative flex items-center">
        <ToolButton
          title="Text highlight color"
          disabled={readOnly}
          onClick={() => controller.exec(setHighlight(lastHighlight))}
        >
          <span className="flex flex-col items-center text-[11px] leading-3">
            <span>ab</span>
            <span className="mt-0.5 h-1 w-4 rounded-sm" style={{ background: lastHighlight }} />
          </span>
        </ToolButton>
        <ToolButton
          title="Highlight colors"
          disabled={readOnly}
          className="!px-0.5"
          onClick={() => toggle("highlight")}
        >
          <span className="text-[9px]">{"\u25be"}</span>
        </ToolButton>
        {openMenu === "highlight" && (
          <Swatches
            colors={HIGHLIGHT_ORDER.map((name) => ({ name: titleCase(name), hex: HIGHLIGHT_COLORS[name] }))}
            noneLabel="No Color"
            onPick={(hex) => {
              if (hex) setLastHighlight(hex);
              controller.exec(setHighlight(hex));
              close();
            }}
          />
        )}
      </div>
      <div className="relative flex items-center">
        <ToolButton
          title="Font color"
          disabled={readOnly}
          onClick={() => controller.exec(setTextColor(lastColor))}
        >
          <span className="flex flex-col items-center text-[11px] font-semibold leading-3">
            <span>A</span>
            <span className="mt-0.5 h-1 w-4 rounded-sm" style={{ background: lastColor }} />
          </span>
        </ToolButton>
        <ToolButton
          title="Font colors"
          disabled={readOnly}
          className="!px-0.5"
          onClick={() => toggle("fontColor")}
        >
          <span className="text-[9px]">{"\u25be"}</span>
        </ToolButton>
        {openMenu === "fontColor" && (
          <Swatches
            colors={[
              ...FONT_COLORS,
              { name: "Black", hex: "#000000" },
              { name: "Dark Gray", hex: "#595959" },
              { name: "Gray", hex: "#808080" },
              { name: "Light Gray", hex: "#bfbfbf" },
              { name: "White", hex: "#ffffff" },
            ]}
            noneLabel="Automatic"
            onPick={(hex) => {
              if (hex) setLastColor(hex);
              controller.exec(setTextColor(hex));
              close();
            }}
          />
        )}
      </div>
    </>
  );
}
