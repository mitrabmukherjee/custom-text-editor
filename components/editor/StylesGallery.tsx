"use client";

import { clearFormatting, sequence } from "@/lib/pm/character";
import type { EditorController } from "@/lib/pm/controller";
import {
  applyStyle,
  currentStyleId,
  displayName,
  findStyle,
  galleryStyles,
  updateStyleToMatch,
} from "@/lib/pm/styles";

type Props = {
  controller: EditorController;
  readOnly?: boolean;
  open: boolean;
  onToggle: () => void;
  onClose: () => void;
};

/** Word's Home ▸ Styles gallery as a dropdown: apply, update to match, clear. */
export function StylesGallery({ controller, readOnly, open, onToggle, onClose }: Props) {
  const state = controller.state;
  const currentId = currentStyleId(state);
  const current = findStyle(state.doc, currentId);
  const styles = open ? galleryStyles(state.doc) : [];
  const run = (fn: () => void) => {
    fn();
    onClose();
  };

  return (
    <div className="relative">
      <button
        type="button"
        className="flex h-7 w-32 items-center justify-between rounded border border-zinc-700 bg-zinc-900 px-2 text-xs text-zinc-200 hover:border-zinc-500 disabled:opacity-40"
        title="Styles"
        aria-label="Styles"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={readOnly}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onToggle}
      >
        <span className="truncate">{current ? displayName(current) : currentId}</span>
        <span className="text-[9px] text-zinc-500">{"\u25be"}</span>
      </button>
      {open && (
        <div
          className="absolute left-0 top-full z-50 mt-0.5 max-h-[24rem] w-64 overflow-y-auto rounded-md border border-zinc-700 bg-zinc-900 py-1 shadow-xl"
          role="menu"
        >
          {styles.map((style) => {
            const runFormat = (style.run ?? {}) as Record<string, unknown>;
            return (
              <button
                key={style.id}
                type="button"
                role="menuitemradio"
                aria-checked={style.id === currentId}
                className={`flex w-full items-center gap-2 px-3 py-1.5 text-left text-zinc-100 hover:bg-zinc-800 ${
                  style.id === currentId ? "bg-zinc-800/70" : ""
                }`}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => run(() => controller.exec(applyStyle(style.id)))}
              >
                <span className="w-3 text-[10px] text-zinc-500">{style.type === "character" ? "a" : "\u00b6"}</span>
                <span
                  className="flex-1 truncate"
                  style={{
                    fontFamily: style.format?.fontFamily,
                    fontSize: `${Math.min(style.format?.fontSize ?? 12, 18)}pt`,
                    fontWeight: runFormat.bold ? 700 : 400,
                    fontStyle: runFormat.italic ? "italic" : "normal",
                  }}
                >
                  {displayName(style)}
                </span>
              </button>
            );
          })}
          <div className="my-1 border-t border-zinc-800" />
          <button
            type="button"
            className="w-full px-3 py-1.5 text-left text-sm text-zinc-200 hover:bg-zinc-800"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => run(() => controller.exec(updateStyleToMatch))}
          >
            Update {current ? displayName(current) : currentId} to match selection
          </button>
          <button
            type="button"
            className="w-full px-3 py-1.5 text-left text-sm text-zinc-200 hover:bg-zinc-800"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => run(() => controller.exec(sequence(applyStyle("Normal"), clearFormatting)))}
          >
            Clear all
          </button>
        </div>
      )}
    </div>
  );
}
