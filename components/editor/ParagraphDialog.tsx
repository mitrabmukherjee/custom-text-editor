"use client";

import { useState } from "react";
import type { ParagraphFormat, TextAlign } from "@/lib/pm/format";
import { Check, Dialog, Field, NumberInput, Section, inputClass } from "./dialogParts";

type Props = {
  format: ParagraphFormat;
  onApply: (patch: Partial<ParagraphFormat>) => void;
  onClose: () => void;
};

type Special = "none" | "firstLine" | "hanging";
type Spacing = "default" | "single" | "onePointFive" | "double" | "atLeast" | "exactly" | "multiple";

const PX_PER_IN = 96;
const PX_PER_PT = 96 / 72;

const toIn = (px: number | undefined) => round((px ?? 0) / PX_PER_IN);
const toPt = (px: number | undefined) => round((px ?? 0) / PX_PER_PT);
const round = (n: number) => Math.round(n * 100) / 100;
/** 0 means "not set", so it isn't written into the format. */
const orUndefined = (n: number) => (n ? n : undefined);

function spacingOf(format: ParagraphFormat): { kind: Spacing; at: number } {
  const { lineRule, lineValue } = format;
  if (!lineRule || lineValue === undefined) return { kind: "default", at: 1 };
  if (lineRule === "exact") return { kind: "exactly", at: toPt(lineValue) };
  if (lineRule === "atLeast") return { kind: "atLeast", at: toPt(lineValue) };
  if (lineValue === 1) return { kind: "single", at: 1 };
  if (lineValue === 1.5) return { kind: "onePointFive", at: 1.5 };
  if (lineValue === 2) return { kind: "double", at: 2 };
  return { kind: "multiple", at: round(lineValue) };
}

function lineOf(kind: Spacing, at: number): Pick<ParagraphFormat, "lineRule" | "lineValue"> {
  switch (kind) {
    case "default":
      return { lineRule: undefined, lineValue: undefined };
    case "single":
      return { lineRule: "auto", lineValue: 1 };
    case "onePointFive":
      return { lineRule: "auto", lineValue: 1.5 };
    case "double":
      return { lineRule: "auto", lineValue: 2 };
    case "multiple":
      return { lineRule: "auto", lineValue: at > 0 ? at : 1 };
    case "atLeast":
      return { lineRule: "atLeast", lineValue: Math.max(0, at) * PX_PER_PT };
    case "exactly":
      return { lineRule: "exact", lineValue: Math.max(1, at) * PX_PER_PT };
  }
}

/** Word's Format → Paragraph dialog: alignment, indentation, spacing and line/page breaks. */
export function ParagraphDialog({ format, onApply, onClose }: Props) {
  const first = format.indentFirstLine ?? 0;
  const initialSpacing = spacingOf(format);
  const [align, setAlign] = useState<TextAlign>(format.align ?? "left");
  // Like Word, "Left" is where the first line starts when the paragraph hangs.
  const [left, setLeft] = useState(toIn((format.indentLeft ?? 0) + Math.min(0, first)));
  const [right, setRight] = useState(toIn(format.indentRight));
  const [special, setSpecial] = useState<Special>(
    first > 0 ? "firstLine" : first < 0 ? "hanging" : "none",
  );
  const [by, setBy] = useState(toIn(Math.abs(first)) || 0.5);
  const [before, setBefore] = useState(toPt(format.spaceBefore));
  const [after, setAfter] = useState(toPt(format.spaceAfter));
  const [spacing, setSpacing] = useState<Spacing>(initialSpacing.kind);
  const [at, setAt] = useState(initialSpacing.at);
  const [widowControl, setWidowControl] = useState(!!format.widowControl);
  const [keepNext, setKeepNext] = useState(!!format.keepNext);
  const [keepLines, setKeepLines] = useState(!!format.keepLines);
  const [suppressLineNumbers, setSuppressLineNumbers] = useState(!!format.suppressLineNumbers);
  const [pageBreakBefore, setPageBreakBefore] = useState(!!format.pageBreakBefore);
  const [contextualSpacing, setContextualSpacing] = useState(!!format.contextualSpacing);

  const changeSpacing = (kind: Spacing) => {
    setSpacing(kind);
    if (kind === "multiple") setAt(3);
    else if (kind === "atLeast" || kind === "exactly") setAt(12);
  };

  const apply = () => {
    const indent = special === "none" ? 0 : (special === "firstLine" ? 1 : -1) * by * PX_PER_IN;
    onApply({
      align: align === "left" ? undefined : align,
      indentLeft: orUndefined(left * PX_PER_IN - Math.min(0, indent)),
      indentRight: orUndefined(right * PX_PER_IN),
      indentFirstLine: orUndefined(indent),
      spaceBefore: orUndefined(before * PX_PER_PT),
      spaceAfter: orUndefined(after * PX_PER_PT),
      ...lineOf(spacing, at),
      widowControl: widowControl || undefined,
      keepNext: keepNext || undefined,
      keepLines: keepLines || undefined,
      suppressLineNumbers: suppressLineNumbers || undefined,
      pageBreakBefore: pageBreakBefore || undefined,
      contextualSpacing: contextualSpacing || undefined,
    });
  };

  const atUnit = spacing === "multiple" ? "×" : "pt";
  const atEnabled = spacing === "multiple" || spacing === "atLeast" || spacing === "exactly";

  return (
    <Dialog title="Paragraph" onClose={onClose} onSubmit={apply}>
        <Section title="General">
          <Field label="Alignment">
            <select
              className={inputClass}
              value={align}
              onChange={(e) => setAlign(e.target.value as TextAlign)}
            >
              <option value="left">Left</option>
              <option value="center">Centered</option>
              <option value="right">Right</option>
              <option value="justify">Justified</option>
            </select>
          </Field>
        </Section>

        <Section title="Indentation">
          <div className="grid grid-cols-2 gap-3">
            <Field label="Left">
              <NumberInput value={left} onChange={setLeft} step={0.1} suffix="in" />
            </Field>
            <Field label="Special">
              <select
                className={inputClass}
                value={special}
                onChange={(e) => setSpecial(e.target.value as Special)}
              >
                <option value="none">(none)</option>
                <option value="firstLine">First line</option>
                <option value="hanging">Hanging</option>
              </select>
            </Field>
            <Field label="Right">
              <NumberInput value={right} onChange={setRight} step={0.1} suffix="in" />
            </Field>
            <Field label="By">
              <NumberInput
                value={by}
                onChange={setBy}
                step={0.1}
                min={0}
                suffix="in"
                disabled={special === "none"}
              />
            </Field>
          </div>
        </Section>

        <Section title="Spacing">
          <div className="grid grid-cols-2 gap-3">
            <Field label="Before">
              <NumberInput value={before} onChange={setBefore} step={6} min={0} suffix="pt" />
            </Field>
            <Field label="Line spacing">
              <select
                className={inputClass}
                value={spacing}
                onChange={(e) => changeSpacing(e.target.value as Spacing)}
              >
                {spacing === "default" && <option value="default">(document default)</option>}
                <option value="single">Single</option>
                <option value="onePointFive">1.5 lines</option>
                <option value="double">Double</option>
                <option value="atLeast">At least</option>
                <option value="exactly">Exactly</option>
                <option value="multiple">Multiple</option>
              </select>
            </Field>
            <Field label="After">
              <NumberInput value={after} onChange={setAfter} step={6} min={0} suffix="pt" />
            </Field>
            <Field label="At">
              <NumberInput
                value={at}
                onChange={setAt}
                step={spacing === "multiple" ? 0.25 : 1}
                min={0}
                suffix={atUnit}
                disabled={!atEnabled}
              />
            </Field>
          </div>
          <div className="mt-3">
            <Check
              label="Don't add space between paragraphs of the same style"
              checked={contextualSpacing}
              onChange={setContextualSpacing}
            />
          </div>
        </Section>

        <Section title="Line and page breaks">
          <div className="grid grid-cols-2 gap-2">
            <Check label="Widow/Orphan control" checked={widowControl} onChange={setWidowControl} />
            <Check label="Keep with next" checked={keepNext} onChange={setKeepNext} />
            <Check label="Keep lines together" checked={keepLines} onChange={setKeepLines} />
            <Check label="Page break before" checked={pageBreakBefore} onChange={setPageBreakBefore} />
            <Check
              label="Suppress line numbers"
              checked={suppressLineNumbers}
              onChange={setSuppressLineNumbers}
            />
          </div>
        </Section>
    </Dialog>
  );
}
