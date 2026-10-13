"use client";

import { useState } from "react";
import { Dialog, Field, NumberInput } from "./dialogParts";

/** Word's "Set Numbering Value" for the selected list item. */
export function NumberingValueDialog({
  value,
  onApply,
  onClose,
}: {
  value: number;
  onApply: (value: number) => void;
  onClose: () => void;
}) {
  const [n, setN] = useState(value);
  return (
    <Dialog title="Set Numbering Value" width="w-72" onClose={onClose} onSubmit={() => onApply(n)}>
      <Field label="Set value to">
        <NumberInput value={n} onChange={setN} step={1} min={0} suffix="" />
      </Field>
    </Dialog>
  );
}
