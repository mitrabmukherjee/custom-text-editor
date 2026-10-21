import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { importDocx } from "./import";
import { layoutDocument } from "@/lib/pm/word-layout/layoutDocument";
import { schema } from "@/lib/pm/schema";

const DEFAULT_PATH =
  "C:/Users/Steora/Downloads/file_1783367057023_Compared_MISSONS__Wayne__Dorothy_Pace_-_May_16__2024__ID1518937.3-_36_pgs_-_Shoumi_draft.docx";

/** Engine page count with hidden deletions omitted (matches default markup UI). */
const EXPECTED_PAGE_COUNT = 34;

describe("MISSIONS Pace transcript pagination", () => {
  it("paginates imported grid transcript with engine layout", async () => {
    const envPath = process.env.MISSIONS_DOCX;
    const path = envPath && existsSync(envPath) ? envPath : DEFAULT_PATH;
    if (!existsSync(path)) return;

    const { doc, pageSetup, layout } = await importDocx(readFileSync(path));
    const root = schema.nodeFromJSON(doc);
    const setup = pageSetup ?? layout?.page!;
    const { pagination, lines } = layoutDocument(root, {
      defaultSetup: setup,
      includeDeletionText: false,
    });

    expect(pagination.pageCount).toBe(EXPECTED_PAGE_COUNT);
    expect(lines.length).toBeGreaterThan(1500);
    expect(pagination.pages.some((p) => p.first)).toBe(true);
  });
});
