"use client";

import { useState } from "react";
import { Dialog, Field, NumberInput } from "./dialogParts";

type Props = {
  count: number;
  space: number;
  onClose: () => void;
  onApply: (count: number, space: number) => void;
};

export function SectionColumnsDialog({ count, space, onClose, onApply }: Props) {
  const [cols, setCols] = useState(count);
  const [gap, setGap] = useState(space);
  return (
    <Dialog title="Columns" width="w-[20rem]" onClose={onClose} onSubmit={() => onApply(cols, gap)}>
      <Field label="Number of columns">
        <NumberInput value={cols} onChange={setCols} step={1} min={1} max={9} suffix="" />
      </Field>
      <Field label="Spacing between columns">
        <NumberInput value={gap} onChange={setGap} step={6} min={0} suffix="px" />
      </Field>
      <p className="text-xs text-zinc-500">
        Columns are exported to Word; the editor still displays a single column on screen.
      </p>
    </Dialog>
  );
}
