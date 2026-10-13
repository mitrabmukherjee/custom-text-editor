"use client";

import { useRef, useState, type PointerEvent, type ReactNode } from "react";
import type { ParagraphFormat, TabStop } from "@/lib/pm/format";
import type { PageSetup } from "@/lib/pm/pagination";

/** Word snaps ruler drags to 1/16 inch (hold Alt to place freely). */
const SNAP = 6;
const TICK = 12;
const TICKS_PER_INCH = 8;
const HEIGHT = 22;
/** Dragging a tab this far off the ruler removes it. */
const REMOVE_DISTANCE = 24;

const TAB_ALIGNS: TabStop["align"][] = ["left", "center", "right", "decimal"];

type Marker = "first" | "hanging" | "left" | "right";
type Drag = { marker: Marker; x: number } | { marker: "tab"; index: number; x: number; remove: boolean };

type Props = {
  setup: PageSetup;
  zoom: number;
  format: ParagraphFormat;
  disabled?: boolean;
  onChange: (patch: Partial<ParagraphFormat>) => void;
};

const round = (n: number) => Math.round(n * 100) / 100;
const orUndefined = (n: number) => (Math.abs(n) < 0.01 ? undefined : round(n));

/** Format change for a drag ending at page x `x` (px from the page's left edge). */
function dragPatch(drag: Drag, format: ParagraphFormat, setup: PageSetup): Partial<ParagraphFormat> {
  const left = format.indentLeft ?? 0;
  const first = format.indentFirstLine ?? 0;
  const offset = drag.x - setup.marginLeft;
  switch (drag.marker) {
    case "first":
      return { indentFirstLine: orUndefined(offset - left) };
    case "hanging":
      // The hanging marker moves the left indent but leaves the first line where it is.
      return { indentLeft: orUndefined(offset), indentFirstLine: orUndefined(left + first - offset) };
    case "left":
      return { indentLeft: orUndefined(offset) };
    case "right":
      return { indentRight: orUndefined(setup.width - setup.marginRight - drag.x) };
    case "tab": {
      const tabs = (format.tabs ?? []).filter((_, i) => i !== drag.index);
      if (!drag.remove) tabs.push({ ...format.tabs![drag.index], pos: round(offset) });
      tabs.sort((a, b) => a.pos - b.pos);
      return { tabs: tabs.length ? tabs : undefined };
    }
  }
}

function TabGlyph({ align }: { align: TabStop["align"] }) {
  const stem = align === "left" ? 1.5 : align === "right" ? 6.5 : 4;
  const [x1, x2] = align === "left" ? [1, 8] : align === "right" ? [0, 7] : [0.5, 7.5];
  return (
    <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden>
      <rect x={stem - 0.75} y="0" width="1.5" height="8" fill="currentColor" />
      <rect x={x1} y="6.5" width={x2 - x1} height="1.5" fill="currentColor" />
      {align === "decimal" && <rect x="6" y="3" width="1.5" height="1.5" fill="currentColor" />}
    </svg>
  );
}

const TRIANGLE_DOWN = "M0.5 0.5 H9.5 L5 6 Z";
const TRIANGLE_UP = "M5 0.5 L9.5 6 H0.5 Z";

function MarkerShape({ path, rect }: { path?: string; rect?: boolean }) {
  return (
    <svg width="10" height="7" viewBox="0 0 10 7" aria-hidden>
      {rect ? (
        <rect x="0.5" y="0.5" width="9" height="6" className="studio-ruler-shape" />
      ) : (
        <path d={path} className="studio-ruler-shape" />
      )}
    </svg>
  );
}

/** Word's horizontal ruler: margins, indent markers and tab stops for the selected paragraph. */
export function Ruler({ setup, zoom, format, disabled, onChange }: Props) {
  const stripRef = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<Drag | null>(null);
  const [tabAlign, setTabAlign] = useState<TabStop["align"]>("left");
  const { width, marginLeft, marginRight } = setup;

  const shown = drag ? { ...format, ...dragPatch(drag, format, setup) } : format;
  const left = marginLeft + (shown.indentLeft ?? 0);
  const firstLine = left + (shown.indentFirstLine ?? 0);
  const right = width - marginRight - (shown.indentRight ?? 0);
  const tabs = format.tabs ?? [];

  const pageX = (e: PointerEvent) => {
    const rect = stripRef.current!.getBoundingClientRect();
    const x = (e.clientX - rect.left) / zoom;
    const snapped = e.altKey ? x : marginLeft + Math.round((x - marginLeft) / SNAP) * SNAP;
    return Math.max(0, Math.min(width, snapped));
  };
  const offRuler = (e: PointerEvent) => {
    const rect = stripRef.current!.getBoundingClientRect();
    return e.clientY < rect.top - REMOVE_DISTANCE || e.clientY > rect.bottom + REMOVE_DISTANCE;
  };

  const startDrag = (e: PointerEvent, start: Drag) => {
    if (disabled || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    setDrag(start);
  };
  const moveDrag = (e: PointerEvent) => {
    if (!drag) return;
    const x = pageX(e);
    setDrag(drag.marker === "tab" ? { ...drag, x, remove: offRuler(e) } : { ...drag, x });
  };
  const endDrag = () => {
    if (!drag) return;
    setDrag(null);
    onChange(dragPatch(drag, format, setup));
  };

  const addTab = (e: PointerEvent) => {
    if (disabled || e.button !== 0 || e.target !== e.currentTarget) return;
    e.preventDefault();
    const x = pageX(e);
    if (x <= marginLeft || x >= width - marginRight) return;
    const next = [...tabs, { pos: round(x - marginLeft), align: tabAlign }].sort((a, b) => a.pos - b.pos);
    onChange({ tabs: next });
  };

  const handle = (key: string, x: number, top: number, title: string, start: Drag, child: ReactNode) => (
    <div
      key={key}
      className="studio-ruler-marker"
      style={{ left: x * zoom, top }}
      title={title}
      onPointerDown={(e) => startDrag(e, start)}
      onPointerMove={moveDrag}
      onPointerUp={endDrag}
      onPointerCancel={() => setDrag(null)}
    >
      {child}
    </div>
  );

  const ticks: ReactNode[] = [];
  const firstTick = -Math.floor(marginLeft / TICK);
  const lastTick = Math.floor((width - marginLeft) / TICK);
  for (let k = firstTick; k <= lastTick; k++) {
    if (k === 0) continue;
    const x = (marginLeft + k * TICK) * zoom;
    if (k % TICKS_PER_INCH === 0) {
      ticks.push(
        <text key={k} x={x} y={15} textAnchor="middle" className="studio-ruler-number">
          {Math.abs(k / TICKS_PER_INCH)}
        </text>,
      );
    } else {
      const h = k % (TICKS_PER_INCH / 2) === 0 ? 6 : 3;
      ticks.push(<rect key={k} x={x - 0.5} y={(HEIGHT - h) / 2} width={1} height={h} className="studio-ruler-tick" />);
    }
  }

  return (
    <div
      className={`studio-ruler${disabled ? " studio-ruler--disabled" : ""}`}
      style={{ width: width * zoom }}
      onMouseDown={(e) => e.preventDefault()}
    >
      <button
        type="button"
        className="studio-ruler-tab-selector"
        title={`${tabAlign[0].toUpperCase()}${tabAlign.slice(1)} tab (click to change)`}
        onClick={() => setTabAlign((a) => TAB_ALIGNS[(TAB_ALIGNS.indexOf(a) + 1) % TAB_ALIGNS.length])}
      >
        <TabGlyph align={tabAlign} />
      </button>
      <div ref={stripRef} className="studio-ruler-strip" style={{ height: HEIGHT }} onPointerDown={addTab}>
        <svg width={width * zoom} height={HEIGHT} className="studio-ruler-scale" aria-hidden>
          <rect x={0} y={0} width={marginLeft * zoom} height={HEIGHT} className="studio-ruler-margin" />
          <rect
            x={(width - marginRight) * zoom}
            y={0}
            width={marginRight * zoom}
            height={HEIGHT}
            className="studio-ruler-margin"
          />
          {ticks}
        </svg>
        {handle("first", firstLine, 0, "First Line Indent", { marker: "first", x: firstLine }, <MarkerShape path={TRIANGLE_DOWN} />)}
        {handle("hanging", left, 9, "Hanging Indent", { marker: "hanging", x: left }, <MarkerShape path={TRIANGLE_UP} />)}
        {handle("left", left, 15, "Left Indent", { marker: "left", x: left }, <MarkerShape rect />)}
        {handle("right", right, 12, "Right Indent", { marker: "right", x: right }, <MarkerShape path={TRIANGLE_UP} />)}
        {tabs.map((tab, i) => {
          const dragging = drag?.marker === "tab" && drag.index === i ? drag : null;
          const x = dragging ? dragging.x : marginLeft + tab.pos;
          return (
            <div
              key={`tab${i}`}
              className={`studio-ruler-marker studio-ruler-tab${dragging?.remove ? " studio-ruler-tab--removing" : ""}`}
              style={{ left: x * zoom, top: 12 }}
              title={`${tab.align[0].toUpperCase()}${tab.align.slice(1)} Tab`}
              onPointerDown={(e) => startDrag(e, { marker: "tab", index: i, x, remove: false })}
              onPointerMove={moveDrag}
              onPointerUp={endDrag}
              onPointerCancel={() => setDrag(null)}
            >
              <TabGlyph align={tab.align} />
            </div>
          );
        })}
      </div>
    </div>
  );
}
