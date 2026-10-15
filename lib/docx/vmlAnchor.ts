/** Round-trippable VML signature placement from `w:pict` / `v:shape`. */
export type VmlAnchorPlacement = "underlineParagraph";

export type VmlAnchorJSON = {
  placement: VmlAnchorPlacement;
  /** From VML style margin-left (pt). */
  marginLeftPt: number;
  /** From VML style margin-top (pt). */
  marginTopPt: number;
  widthPt: number;
  heightPt: number;
  /** Carrier paragraph `w:ind` left (px). */
  carrierIndentLeftPx?: number;
  /** Body paragraphs between carrier and underline target. */
  paragraphsBetween: number;
};

export function parseVmlAnchorAttr(raw: unknown): VmlAnchorJSON | null {
  if (!raw || typeof raw !== "object") return null;
  const o = raw as VmlAnchorJSON;
  if (o.placement !== "underlineParagraph") return null;
  if (!Number.isFinite(o.marginLeftPt) || !Number.isFinite(o.marginTopPt)) return null;
  return o;
}
