import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { importDocx } from "./import";
import { layoutDocument } from "@/lib/pm/word-layout/layoutDocument";
import { schema } from "@/lib/pm/schema";

const DEFAULT_PATH =
  "C:/Users/Steora/Downloads/MASTER-BATCH-2013/MASTER-BATCH-2013/TF-plus-data-v2-2013-FLAT-B2/PAPASOTIRIOU - MAY 10  2013 - ID 61861 - COMPLETE PROCEEDING - MARGARET OSADET - JG - RELEASED.docx";

describe("PAPASOTIRIOU pagination model", () => {
  it("paginates title, TOC, proceedings, and certification with margin-aware section breaks", async () => {
    const envPath = process.env.PAPASOTIRIOU_DOCX;
    const path = envPath && existsSync(envPath) ? envPath : DEFAULT_PATH;
    if (!existsSync(path)) return;

    const { doc, pageSetup, layout } = await importDocx(readFileSync(path));
    const root = schema.nodeFromJSON(doc);
    const setup = pageSetup ?? layout?.page!;
    const { pagination } = layoutDocument(root, { defaultSetup: setup });
    expect(pagination.pageCount).toBeGreaterThanOrEqual(4);
    expect(pagination.pages.some((p) => p.first)).toBe(true);
  });
});
