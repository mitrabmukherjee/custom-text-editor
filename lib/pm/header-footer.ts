import type { EditorState, Transaction } from "prosemirror-state";
import type { HFBlock, HFParagraph, HeaderFooterSet, PageInfo } from "./format";
import { hfSetParagraphs } from "./format";
import { patchSectionLayout, sectionsOf } from "./sections";

export type HeaderFooterTarget = {
  section: number;
  part: "header" | "footer";
  variant: keyof HeaderFooterSet;
};

/** Plain text for a default header/footer (PAGE becomes `{PAGE}`). */
export function hfPlainText(paragraphs: HFParagraph[] | undefined): string {
  if (!paragraphs?.length) return "";
  return paragraphs
    .map((p) =>
      p.runs
        .map((r) => {
          if (r.kind === "field") {
            if (r.field === "PAGE") return "{PAGE}";
            if (r.field === "SECTIONPAGES") return "{SECTIONPAGES}";
            return "{NUMPAGES}";
          }
          if (r.kind === "tab") return "\t";
          return r.text;
        })
        .join(""),
    )
    .join("\n");
}

function parseRuns(line: string): HFParagraph["runs"] {
  const runs: HFParagraph["runs"] = [];
  const re = /\{(PAGE|NUMPAGES|SECTIONPAGES)\}/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(line))) {
    if (m.index > last) runs.push({ kind: "text", text: line.slice(last, m.index) });
    const field =
      m[1] === "PAGE" ? "PAGE" : m[1] === "SECTIONPAGES" ? "SECTIONPAGES" : "NUMPAGES";
    runs.push({ kind: "field", field });
    last = m.index + m[0].length;
  }
  if (last < line.length) runs.push({ kind: "text", text: line.slice(last) });
  if (!runs.length) runs.push({ kind: "text", text: "" });
  return runs;
}

export function textToHFParagraphs(text: string): HFParagraph[] {
  const lines = text.split(/\n/);
  if (!lines.length) return [{ format: {}, runs: [{ kind: "text", text: "" }] }];
  return lines.map((line) => ({ format: {}, runs: parseRuns(line) }));
}

export function setHeaderFooterText(
  state: EditorState,
  target: HeaderFooterTarget,
  text: string,
): Transaction | null {
  const paragraphs = textToHFParagraphs(text);
  return patchSectionLayout(state, target.section, (section) => {
    const key = target.part === "header" ? "headers" : "footers";
    const set: HeaderFooterSet = { ...(section[key] ?? {}) };
    set[target.variant] = paragraphs.map((p) => ({ ...p, kind: "paragraph" as const }));
    return { ...section, [key]: set };
  });
}

/** Header/footer target for a page in the document (for double-click editing). */
export function headerFooterTargetAt(
  state: EditorState,
  page: Pick<PageInfo, "section" | "first" | "number">,
  part: "header" | "footer",
  evenAndOdd: boolean,
): HeaderFooterTarget {
  const sections = sectionsOf(state);
  const section = sections[Math.min(page.section, sections.length - 1)] ?? {};
  const set = part === "header" ? section.headers : section.footers;
  let variant: keyof HeaderFooterSet = "default";
  if (page.first && section.titlePage && set?.first) variant = "first";
  else if (evenAndOdd && page.number % 2 === 0 && set?.even) variant = "even";
  return { section: page.section, part, variant };
}

export function getHeaderFooterText(state: EditorState, target: HeaderFooterTarget): string {
  const section = sectionsOf(state)[Math.min(target.section, sectionsOf(state).length - 1)] ?? {};
  const set = target.part === "header" ? section.headers : section.footers;
  const blocks = set?.[target.variant] ?? set?.default;
  return hfPlainText(hfSetParagraphs(blocks));
}
