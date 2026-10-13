import type { Mark } from "prosemirror-model";
import type { MarkJSON } from "../json";
import { colorName } from "../textColors";

export type UnderlineStyle =
  | "single"
  | "words"
  | "double"
  | "thick"
  | "dotted"
  | "dash"
  | "dotDash"
  | "wave";
export type StrikeStyle = "single" | "double";
export type CapsStyle = "all" | "small";
export type VertAlign = "superscript" | "subscript";

/**
 * Character formatting of a run. Formatting revisions record it so rejecting them can
 * restore the text. Optional fields are absent in revisions saved by older versions.
 */
export type RunFormat = {
  bold: boolean;
  italic: boolean;
  /** `null` means the paragraph's font. */
  fontFamily: string | null;
  /** Points; `null` means the paragraph's size. */
  fontSize: number | null;
  underline?: UnderlineStyle | null;
  strike?: StrikeStyle | null;
  /** "#rrggbb"; `null` is Word's "Automatic". */
  color?: string | null;
  /** "#rrggbb" highlight or run shading. */
  highlight?: string | null;
  caps?: CapsStyle | null;
  vertAlign?: VertAlign | null;
};

export const PLAIN_RUN: Required<RunFormat> = {
  bold: false,
  italic: false,
  fontFamily: null,
  fontSize: null,
  underline: null,
  strike: null,
  color: null,
  highlight: null,
  caps: null,
  vertAlign: null,
};

/** Marks whose changes are tracked as formatting revisions. */
export const FORMAT_MARK_NAMES = [
  "strong",
  "em",
  "textStyle",
  "underline",
  "strike",
  "textColor",
  "highlight",
  "caps",
  "vertAlign",
];

export function fullRunFormat(format: RunFormat): Required<RunFormat> {
  return {
    ...PLAIN_RUN,
    ...Object.fromEntries(Object.entries(format).filter(([, v]) => v !== undefined)),
  } as Required<RunFormat>;
}

export function runFormatOf(marks: readonly Mark[]): RunFormat {
  const find = (name: string) => marks.find((m) => m.type.name === name);
  const style = find("textStyle");
  return {
    bold: !!find("strong"),
    italic: !!find("em"),
    fontFamily: (style?.attrs.fontFamily as string | null) ?? null,
    fontSize: (style?.attrs.fontSize as number | null) ?? null,
    underline: (find("underline")?.attrs.style as UnderlineStyle | undefined) ?? null,
    strike: (find("strike")?.attrs.style as StrikeStyle | undefined) ?? null,
    color: (find("textColor")?.attrs.color as string | undefined) ?? null,
    highlight: (find("highlight")?.attrs.color as string | undefined) ?? null,
    caps: (find("caps")?.attrs.style as CapsStyle | undefined) ?? null,
    vertAlign: (find("vertAlign")?.attrs.align as VertAlign | undefined) ?? null,
  };
}

/** The formatting marks for a run format (JSON, so the DOCX importer can use it too). */
export function runFormatMarks(format: RunFormat): MarkJSON[] {
  const f = fullRunFormat(format);
  const marks: MarkJSON[] = [];
  if (f.fontFamily || f.fontSize) {
    marks.push({ type: "textStyle", attrs: { fontFamily: f.fontFamily, fontSize: f.fontSize } });
  }
  if (f.bold) marks.push({ type: "strong" });
  if (f.italic) marks.push({ type: "em" });
  if (f.underline) marks.push({ type: "underline", attrs: { style: f.underline } });
  if (f.strike) marks.push({ type: "strike", attrs: { style: f.strike } });
  if (f.color) marks.push({ type: "textColor", attrs: { color: f.color } });
  if (f.highlight) marks.push({ type: "highlight", attrs: { color: f.highlight } });
  if (f.caps) marks.push({ type: "caps", attrs: { style: f.caps } });
  if (f.vertAlign) marks.push({ type: "vertAlign", attrs: { align: f.vertAlign } });
  return marks;
}

export function sameRunFormat(a: RunFormat, b: RunFormat) {
  const x = fullRunFormat(a);
  const y = fullRunFormat(b);
  return (Object.keys(PLAIN_RUN) as (keyof RunFormat)[]).every((k) => x[k] === y[k]);
}

const UNDERLINE_NAMES: Record<UnderlineStyle, string> = {
  single: "Underline",
  words: "Underline (Words only)",
  double: "Double underline",
  thick: "Thick underline",
  dotted: "Dotted underline",
  dash: "Dashed underline",
  dotDash: "Dot-dash underline",
  wave: "Wavy underline",
};

/** Word's balloon wording, e.g. "Font: Italic" or "Font: Arial, 14 pt, Not Bold"; "" if equal. */
export function describeFormatChange(oldFormat: RunFormat, nowFormat: RunFormat): string {
  const old = fullRunFormat(oldFormat);
  const now = fullRunFormat(nowFormat);
  const parts: string[] = [];
  if (now.fontFamily !== old.fontFamily) parts.push(now.fontFamily ?? "(Default)");
  if (now.fontSize !== old.fontSize) parts.push(now.fontSize ? `${now.fontSize} pt` : "(Default) size");
  if (now.bold !== old.bold) parts.push(now.bold ? "Bold" : "Not Bold");
  if (now.italic !== old.italic) parts.push(now.italic ? "Italic" : "Not Italic");
  if (now.underline !== old.underline) {
    parts.push(now.underline ? UNDERLINE_NAMES[now.underline] : "No underline");
  }
  if (now.strike !== old.strike) {
    parts.push(
      now.strike === "double"
        ? "Double strikethrough"
        : now.strike
          ? "Strikethrough"
          : "Not Strikethrough",
    );
  }
  if (now.color !== old.color) {
    parts.push(`Font color: ${now.color ? colorName(now.color) : "Auto"}`);
  }
  if (now.highlight !== old.highlight) {
    parts.push(now.highlight ? `Highlight (${colorName(now.highlight)})` : "Not Highlight");
  }
  if (now.caps !== old.caps) {
    parts.push(
      now.caps === "all"
        ? "All caps"
        : now.caps === "small"
          ? "Small caps"
          : old.caps === "all"
            ? "Not All caps"
            : "Not Small caps",
    );
  }
  if (now.vertAlign !== old.vertAlign) {
    parts.push(
      now.vertAlign === "superscript"
        ? "Superscript"
        : now.vertAlign === "subscript"
          ? "Subscript"
          : "Not Superscript/ Subscript",
    );
  }
  return parts.length ? `Font: ${parts.join(", ")}` : "";
}
