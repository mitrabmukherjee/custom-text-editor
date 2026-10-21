import { readFileSync, readdirSync, existsSync } from "fs";
import path from "path";
import { describe, expect, it } from "vitest";
import { importDocx } from "@/lib/docx/import";
import { schema } from "@/lib/pm/schema";
import { layoutDocument } from "./layoutDocument";
import { assertGolden, layoutGoldenFrom } from "./golden";
import type { LayoutGolden } from "./types";

const dirs = [
  path.join(process.cwd(), "public", "fixtures", "batch"),
  path.join(process.cwd(), "public", "fixtures"),
];

function fixturePairs(): { docx: string; golden: string; name: string }[] {
  const out: { docx: string; golden: string; name: string }[] = [];
  const seen = new Set<string>();
  for (const dir of dirs) {
    if (!existsSync(dir)) continue;
    for (const name of readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".docx"))) {
      if (seen.has(name)) continue;
      seen.add(name);
      const golden = path.join(dir, `${name}.layout-golden.json`);
      if (!existsSync(golden)) continue;
      out.push({ docx: path.join(dir, name), golden, name });
    }
  }
  return out;
}

describe("word layout goldens", () => {
  const pairs = fixturePairs();
  it("has at least one committed golden fixture", () => {
    expect(pairs.length).toBeGreaterThan(0);
  });

  for (const { docx, golden, name } of pairs) {
    it(`matches golden for ${name}`, async () => {
      const expected = JSON.parse(readFileSync(golden, "utf8")) as LayoutGolden;
      const { doc, pageSetup, layout } = await importDocx(readFileSync(docx));
      const node = schema.nodeFromJSON(doc);
      const result = layoutDocument(node, { defaultSetup: pageSetup ?? layout?.page });
      const actual = layoutGoldenFrom(name, result.lines, result.pagination);
      const check = assertGolden(actual, expected);
      expect(check.ok, check.diff).toBe(true);
    });
  }
});
