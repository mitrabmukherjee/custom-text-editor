/**
 * postMessage contract between a host page and the /embed editor iframe.
 * All message types are prefixed with "te:". DOCX payloads are ArrayBuffers
 * and are sent as transferables.
 */

import type { DocJSON } from "@/lib/pm/json";
import type { StyleDefinition } from "@/lib/pm/format";

export type EditorActionName = "playPause" | "print";

/** "accepted" hides deletions (the text as it will read); "original" hides insertions. */
export type PlainTextMode = "accepted" | "original";

export type HostToEditorMessage =
  | { type: "te:load"; plainText: string; editable?: boolean }
  | { type: "te:load-docx"; buffer: ArrayBuffer; editable?: boolean }
  | { type: "te:get-plain"; requestId: string; mode?: PlainTextMode }
  | { type: "te:get-json"; requestId: string }
  | { type: "te:get-docx"; requestId: string }
  | { type: "te:set-plain"; plainText: string }
  | { type: "te:set-editable"; editable: boolean }
  | { type: "te:set-author"; id: string; name: string }
  | { type: "te:set-track-changes"; enabled: boolean }
  | { type: "te:focus" }
  | { type: "te:load-template"; requestId?: string; styles: StyleDefinition[] };

export type EditorToHostMessage =
  | { type: "te:ready" }
  | { type: "te:loaded"; requestId?: string }
  | { type: "te:plain"; requestId: string; text: string }
  | { type: "te:json"; requestId: string; json: DocJSON }
  | { type: "te:docx"; requestId: string; buffer: ArrayBuffer }
  | { type: "te:change"; pageCount: number; changeCount: number }
  | { type: "te:layout"; pageCount: number }
  | { type: "te:action"; name: EditorActionName }
  | { type: "te:error"; message: string; requestId?: string };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return (
    value instanceof ArrayBuffer ||
    Object.prototype.toString.call(value) === "[object ArrayBuffer]"
  );
}

const optionalBoolean = (v: unknown) => v === undefined || typeof v === "boolean";
const isCount = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= 0;

export function isHostMessage(data: unknown): data is HostToEditorMessage {
  if (!isRecord(data)) return false;
  switch (data.type) {
    case "te:load":
      return typeof data.plainText === "string" && optionalBoolean(data.editable);
    case "te:load-docx":
      return isArrayBuffer(data.buffer) && optionalBoolean(data.editable);
    case "te:get-plain":
      return (
        typeof data.requestId === "string" &&
        (data.mode === undefined || data.mode === "accepted" || data.mode === "original")
      );
    case "te:get-json":
    case "te:get-docx":
      return typeof data.requestId === "string";
    case "te:set-plain":
      return typeof data.plainText === "string";
    case "te:set-editable":
      return typeof data.editable === "boolean";
    case "te:set-author":
      return typeof data.id === "string" && data.id !== "" && typeof data.name === "string";
    case "te:set-track-changes":
      return typeof data.enabled === "boolean";
    case "te:focus":
      return true;
    case "te:load-template":
      return (
        (data.requestId === undefined || typeof data.requestId === "string") &&
        Array.isArray(data.styles) &&
        data.styles.every(
          (s) =>
            typeof s === "object" &&
            s !== null &&
            typeof (s as StyleDefinition).id === "string" &&
            typeof (s as StyleDefinition).name === "string",
        )
      );
    default:
      return false;
  }
}

export function isEditorMessage(data: unknown): data is EditorToHostMessage {
  if (!isRecord(data)) return false;
  switch (data.type) {
    case "te:ready":
    case "te:loaded":
      return data.requestId === undefined || typeof data.requestId === "string";
    case "te:plain":
      return typeof data.requestId === "string" && typeof data.text === "string";
    case "te:json":
      return (
        typeof data.requestId === "string" &&
        isRecord(data.json) &&
        data.json.type === "doc"
      );
    case "te:docx":
      return typeof data.requestId === "string" && isArrayBuffer(data.buffer);
    case "te:change":
      return isCount(data.pageCount) && isCount(data.changeCount);
    case "te:layout":
      return isCount(data.pageCount);
    case "te:action":
      return data.name === "playPause" || data.name === "print";
    case "te:error":
      return typeof data.message === "string";
    default:
      return false;
  }
}
