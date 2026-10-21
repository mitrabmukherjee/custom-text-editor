import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { importDocx } from "./import";
import { schema } from "@/lib/pm/schema";
import { layoutDocument } from "@/lib/pm/word-layout/layoutDocument";
import { assertGolden, layoutGoldenFrom } from "@/lib/pm/word-layout/golden";
import type { LayoutGolden } from "@/lib/pm/word-layout/types";

const GOLDEN_PATH = path.join(
  process.cwd(),
  "public",
  "fixtures",
  "papasotiriou-may-2013.layout-golden.json",
);

const DEFAULT_DOCX =
  "C:/Users/Steora/Downloads/MASTER-BATCH-2013/MASTER-BATCH-2013/TF-plus-data-v2-2013-FLAT-B2/PAPASOTIRIOU - MAY 10  2013 - ID 61861 - COMPLETE PROCEEDING - MARGARET OSADET - JG - RELEASED.docx";

describe("PAPASOTIRIOU layout golden", () => {
  const committed = existsSync(GOLDEN_PATH)
    ? (JSON.parse(readFileSync(GOLDEN_PATH, "utf8")) as LayoutGolden)
    : null;

  it("committed golden exists with title/TOC page stats", () => {
    expect(committed).not.toBeNull();
    expect(committed!.pageCount).toBeGreaterThanOrEqual(4);
    expect(committed!.pages12).toBeDefined();
    expect(committed!.pages12!.perPageLines[0]).toBeGreaterThanOrEqual(28);
    expect(committed!.pages12!.perPageLines[1]).toBeGreaterThanOrEqual(24);
    expect(committed!.lineCount).toBeGreaterThan(100);
  });

  it("matches committed golden when PAPASOTIRIOU_DOCX is available", async () => {
    if (!committed) return;
    const envPath = process.env.PAPASOTIRIOU_DOCX;
    const docxPath = envPath && existsSync(envPath) ? envPath : DEFAULT_DOCX;
    if (!existsSync(docxPath)) return;

    const { doc, pageSetup, layout } = await importDocx(readFileSync(docxPath));
    const node = schema.nodeFromJSON(doc);
    const result = layoutDocument(node, { defaultSetup: pageSetup ?? layout?.page });
    const actual = layoutGoldenFrom(committed.fixture, result.lines, result.pagination);
    const check = assertGolden(actual, committed);
    expect(check.ok, check.diff).toBe(true);
    if (committed.pages12) {
      const p0 = result.pagination.placements.filter((p) => p.page === 0).length;
      const p1 = result.pagination.placements.filter((p) => p.page === 1).length;
      expect([p0, p1]).toEqual(committed.pages12.perPageLines);
    }
  });
});
