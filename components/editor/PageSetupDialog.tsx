"use client";

import { useState } from "react";
import type { PageSetup } from "@/lib/pm/pagination/pageSetup";
import { PAGE_SETUP_PRESETS } from "@/lib/pm/pagination/pageSetup";
import { Check, Dialog, Field, NumberInput, Section } from "./dialogParts";

const PX_PER_IN = 96;

type Props = {
  setup: PageSetup;
  onClose: () => void;
  onApply: (setup: PageSetup) => void;
};

const toIn = (px: number) => Math.round((px / PX_PER_IN) * 100) / 100;
const toPx = (inch: number) => Math.round(inch * PX_PER_IN);

export function PageSetupDialog({ setup, onClose, onApply }: Props) {
  const [preset, setPreset] = useState("custom");
  const [portrait, setPortrait] = useState(setup.height >= setup.width);
  const [width, setWidth] = useState(toIn(setup.width));
  const [height, setHeight] = useState(toIn(setup.height));
  const [marginTop, setMarginTop] = useState(toIn(setup.marginTop));
  const [marginRight, setMarginRight] = useState(toIn(setup.marginRight));
  const [marginBottom, setMarginBottom] = useState(toIn(setup.marginBottom));
  const [marginLeft, setMarginLeft] = useState(toIn(setup.marginLeft));

  const pickPreset = (id: string) => {
    setPreset(id);
    const p = PAGE_SETUP_PRESETS.find((x) => x.id === id);
    if (!p) return;
    const s = p.setup;
    setPortrait(s.height >= s.width);
    setWidth(toIn(s.width));
    setHeight(toIn(s.height));
    setMarginTop(toIn(s.marginTop));
    setMarginRight(toIn(s.marginRight));
    setMarginBottom(toIn(s.marginBottom));
    setMarginLeft(toIn(s.marginLeft));
  };

  const apply = () => {
    let w = toPx(width);
    let h = toPx(height);
    if (!portrait && w < h) [w, h] = [h, w];
    if (portrait && w > h) [w, h] = [h, w];
    onApply({
      ...setup,
      width: w,
      height: h,
      marginTop: toPx(marginTop),
      marginRight: toPx(marginRight),
      marginBottom: toPx(marginBottom),
      marginLeft: toPx(marginLeft),
    });
  };

  return (
    <Dialog title="Page Setup" width="w-[28rem]" onClose={onClose} onSubmit={apply}>
      <Section title="Paper">
        <Field label="Paper size">
          <select
            className="w-full rounded border border-zinc-700 bg-zinc-950 px-2 py-1 text-sm"
            value={preset}
            onChange={(e) => pickPreset(e.target.value)}
          >
            <option value="custom">Custom</option>
            {PAGE_SETUP_PRESETS.map((p) => (
              <option key={p.id} value={p.id}>{p.label}</option>
            ))}
          </select>
        </Field>
        <div className="mt-2 flex gap-4">
          <Check label="Portrait" checked={portrait} onChange={setPortrait} />
          <Check label="Landscape" checked={!portrait} onChange={(v) => setPortrait(!v)} />
        </div>
        <div className="mt-3 grid grid-cols-2 gap-3">
          <Field label="Width">
            <NumberInput value={width} onChange={setWidth} step={0.1} min={1} suffix="in" />
          </Field>
          <Field label="Height">
            <NumberInput value={height} onChange={setHeight} step={0.1} min={1} suffix="in" />
          </Field>
        </div>
      </Section>
      <Section title="Margins">
        <div className="grid grid-cols-2 gap-3">
          <Field label="Top">
            <NumberInput value={marginTop} onChange={setMarginTop} step={0.1} min={0} suffix="in" />
          </Field>
          <Field label="Bottom">
            <NumberInput value={marginBottom} onChange={setMarginBottom} step={0.1} min={0} suffix="in" />
          </Field>
          <Field label="Left">
            <NumberInput value={marginLeft} onChange={setMarginLeft} step={0.1} min={0} suffix="in" />
          </Field>
          <Field label="Right">
            <NumberInput value={marginRight} onChange={setMarginRight} step={0.1} min={0} suffix="in" />
          </Field>
        </div>
      </Section>
    </Dialog>
  );
}
