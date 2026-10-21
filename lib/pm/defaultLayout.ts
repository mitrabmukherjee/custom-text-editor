import type { DocJSON } from "./json";
import type { DocumentLayout } from "./format";
import { LETTER_PAGE_SETUP, type PageSetup } from "./pagination/pageSetup";
import { BUILTIN_STYLES } from "./styles";
import { plainTextToDoc } from "@/lib/serializers/plainText";

const DEFAULT_TAB_STOP = 48;
/** Matches importDocx when compatibilityMode >= 15. */
const JUSTIFY_SHRINK = 0.25;
const MODERN_WORD_COMPAT = 15;

/** Synthetic Word layout for studio documents created without a .docx import. */
export function defaultStudioLayout(pageSetup: PageSetup = LETTER_PAGE_SETUP): DocumentLayout {
  return {
    page: pageSetup,
    defaultTabStop: DEFAULT_TAB_STOP,
    justifyShrink: JUSTIFY_SHRINK,
    compat: { compatibilityMode: MODERN_WORD_COMPAT },
    styles: [...BUILTIN_STYLES],
    layoutEngine: false,
  };
}

function hasLayout(doc: DocJSON): boolean {
  const layout = doc.attrs?.layout;
  return layout != null && typeof layout === "object";
}

/** Attach default studio layout when the document has none (imported layout is preserved). */
export function ensureStudioLayout(doc: DocJSON, pageSetup?: PageSetup): DocJSON {
  if (hasLayout(doc)) return doc;
  return {
    ...doc,
    attrs: {
      ...doc.attrs,
      layout: defaultStudioLayout(pageSetup ?? LETTER_PAGE_SETUP),
    },
  };
}

export function studioDocFromPlainText(plain: string, pageSetup?: PageSetup): DocJSON {
  return ensureStudioLayout(plainTextToDoc(plain), pageSetup);
}
