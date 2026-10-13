import type { ParagraphFormat } from "../format";

const ALIGN_NAMES = { left: "Left", center: "Centered", right: "Right", justify: "Justified" };

const inches = (px: number | undefined) => `${Math.round(((px ?? 0) / 96) * 100) / 100}"`;
const points = (px: number | undefined) => `${Math.round(((px ?? 0) * 72) / 96 * 10) / 10} pt`;

function lineSpacing(f: ParagraphFormat): string {
  const rule = f.lineRule ?? "auto";
  const value = f.lineValue ?? 1;
  if (rule === "exact") return `Exactly ${points(value)}`;
  if (rule === "atLeast") return `At least ${points(value)}`;
  if (value === 1) return "single";
  if (value === 1.5) return "1.5 lines";
  if (value === 2) return "Double";
  return `Multiple ${Math.round(value * 100) / 100} li`;
}

const FLAGS: [keyof ParagraphFormat, string, string][] = [
  ["keepNext", "Keep with next", "Don't keep with next"],
  ["keepLines", "Keep lines together", "Don't keep lines together"],
  ["widowControl", "Widow/Orphan control", "No widow/orphan control"],
  ["pageBreakBefore", "Page break before", "No page break before"],
  ["contextualSpacing", "Don't add space between paragraphs of the same style", "Add space between paragraphs of the same style"],
  ["suppressLineNumbers", "Suppress line numbers", "Don't suppress line numbers"],
];

/** Word's balloon wording for a paragraph formatting revision; "" when nothing visible changed. */
export function describeParagraphChange(
  oldFormat: ParagraphFormat | null,
  nowFormat: ParagraphFormat | null,
  styleName: (id: string) => string = (id) => id,
): string {
  const old = oldFormat ?? {};
  const now = nowFormat ?? {};
  const parts: string[] = [];
  if (now.styleId !== old.styleId && (now.styleId || old.styleId)) {
    parts.push(`Style: ${styleName(now.styleId ?? "Normal")}`);
  }
  if ((now.align ?? "left") !== (old.align ?? "left")) parts.push(ALIGN_NAMES[now.align ?? "left"]);

  const indent: string[] = [];
  if ((now.indentLeft ?? 0) !== (old.indentLeft ?? 0)) indent.push(`Left: ${inches(now.indentLeft)}`);
  if ((now.indentRight ?? 0) !== (old.indentRight ?? 0)) indent.push(`Right: ${inches(now.indentRight)}`);
  const first = now.indentFirstLine ?? 0;
  if (first !== (old.indentFirstLine ?? 0)) {
    indent.push(first < 0 ? `Hanging: ${inches(-first)}` : `First line: ${inches(first)}`);
  }
  if (indent.length) parts.push(`Indent: ${indent.join(", ")}`);

  if ((now.spaceBefore ?? 0) !== (old.spaceBefore ?? 0)) parts.push(`Space Before: ${points(now.spaceBefore)}`);
  if ((now.spaceAfter ?? 0) !== (old.spaceAfter ?? 0)) parts.push(`Space After: ${points(now.spaceAfter)}`);
  if (lineSpacing(now) !== lineSpacing(old)) parts.push(`Line spacing: ${lineSpacing(now)}`);
  for (const [key, on, off] of FLAGS) {
    if (!!now[key] !== !!old[key]) parts.push(now[key] ? on : off);
  }
  if (JSON.stringify(now.tabs ?? []) !== JSON.stringify(old.tabs ?? [])) parts.push("Tab stops");
  return parts.join(", ");
}
