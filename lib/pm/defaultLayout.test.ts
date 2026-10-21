import { describe, expect, it } from "vitest";
import { LETTER_PAGE_SETUP } from "./pagination/pageSetup";
import {
  defaultStudioLayout,
  ensureStudioLayout,
  studioDocFromPlainText,
} from "./defaultLayout";
import type { DocumentLayout } from "./format";

describe("defaultStudioLayout", () => {
  it("includes page geometry and Word defaults", () => {
    const layout = defaultStudioLayout(LETTER_PAGE_SETUP);
    expect(layout.page).toEqual(LETTER_PAGE_SETUP);
    expect(layout.defaultTabStop).toBe(48);
    expect(layout.justifyShrink).toBe(0.25);
    expect(layout.compat?.compatibilityMode).toBe(15);
    expect(layout.styles?.some((s) => s.id === "Normal")).toBe(true);
    expect(layout.layoutEngine).toBe(false);
  });
});

describe("ensureStudioLayout", () => {
  it("adds layout when missing", () => {
    const doc = ensureStudioLayout({ type: "doc", content: [{ type: "paragraph" }] });
    expect((doc.attrs?.layout as DocumentLayout).page).toEqual(LETTER_PAGE_SETUP);
    expect((doc.attrs?.layout as DocumentLayout).layoutEngine).toBe(false);
  });

  it("is idempotent", () => {
    const first = ensureStudioLayout({ type: "doc", content: [{ type: "paragraph" }] });
    const second = ensureStudioLayout(first);
    expect(second).toBe(first);
  });

  it("does not replace an existing layout", () => {
    const imported: DocumentLayout = {
      page: { ...LETTER_PAGE_SETUP, marginTop: 120 },
      styles: [{ id: "Normal", name: "Normal", type: "paragraph" }],
    };
    const doc = {
      type: "doc" as const,
      attrs: { layout: imported },
      content: [{ type: "paragraph" }],
    };
    const out = ensureStudioLayout(doc);
    expect((out.attrs?.layout as DocumentLayout).page?.marginTop).toBe(120);
  });
});

describe("studioDocFromPlainText", () => {
  it("builds paragraphs with studio layout", () => {
    const doc = studioDocFromPlainText("Hello");
    expect(doc.content?.[0]?.type).toBe("paragraph");
    expect((doc.attrs?.layout as DocumentLayout).page).toEqual(LETTER_PAGE_SETUP);
  });

  it("round-trips through exportDocx and importDocx with layout", async () => {
    const { exportDocx, importDocx } = await import("@/lib/docx");
    const doc = studioDocFromPlainText("Typed then saved");
    const buffer = await exportDocx(doc, LETTER_PAGE_SETUP);
    const { doc: back } = await importDocx(buffer);
    expect((back as { attrs?: { layout?: unknown } }).attrs?.layout).toBeTruthy();
  });
});
