"use client";

import { useCallback, useSyncExternalStore } from "react";

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2;
export const ZOOM_STEP = 0.1;

const KEY = "te:zoom";
const listeners = new Set<() => void>();
let current: number | null = null;

export function clampZoom(zoom: number) {
  const z = Math.round(zoom * 100) / 100;
  return Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Number.isFinite(z) ? z : 1));
}

function read(): number {
  if (current === null) {
    try {
      const stored = Number(sessionStorage.getItem(KEY));
      current = stored ? clampZoom(stored) : 1;
    } catch {
      current = 1;
    }
  }
  return current;
}

function subscribe(listener: () => void) {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** Zoom level kept for the browser session (sessionStorage), like Word keeping it per window. */
export function useZoom(): [number, (zoom: number | ((zoom: number) => number)) => void] {
  const zoom = useSyncExternalStore(subscribe, read, () => 1);
  const setZoom = useCallback((next: number | ((zoom: number) => number)) => {
    current = clampZoom(typeof next === "function" ? next(read()) : next);
    try {
      sessionStorage.setItem(KEY, String(current));
    } catch {
      // Storage unavailable (e.g. a sandboxed iframe): zoom lasts until reload.
    }
    for (const listener of listeners) listener();
  }, []);
  return [zoom, setZoom];
}
