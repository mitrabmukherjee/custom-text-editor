import type { TextMeasurer } from "./measure";
import { setHarfBuzzMeasurer } from "./measure";

/**
 * Hook for HarfBuzz WASM when vendored. Call `registerHarfBuzzMeasurer` from a
 * bootstrap script after loading WASM + font binaries.
 */
export function registerHarfBuzzMeasurer(measurer: TextMeasurer) {
  setHarfBuzzMeasurer(measurer);
}

export async function loadHarfBuzz(): Promise<boolean> {
  return false;
}
