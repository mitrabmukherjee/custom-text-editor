/**
 * Write layout golden JSON for batch fixtures (run after engine changes).
 * Usage: node scripts/word-layout-golden.mjs
 */
import { readFileSync, writeFileSync, readdirSync } from "fs";
import path from "path";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const root = process.cwd();

async function main() {
  const { importDocx } = await import("../lib/docx/import.ts");
  const { schema } = await import("../lib/pm/schema.ts");
  const { layoutDocument } = await import("../lib/pm/word-layout/layoutDocument.ts");
  const { layoutGoldenFrom } = await import("../lib/pm/word-layout/golden.ts");

  const dirs = [
    path.join(root, "public", "fixtures", "batch"),
    path.join(root, "public", "fixtures"),
  ];
  const seen = new Set();
  for (const dir of dirs) {
    let files = [];
    try {
      files = readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".docx"));
    } catch {
      continue;
    }
    for (const name of files) {
      if (seen.has(name)) continue;
      seen.add(name);
      const full = path.join(dir, name);
      const { doc, pageSetup, layout } = await importDocx(readFileSync(full));
      const node = schema.nodeFromJSON(doc);
      const result = layoutDocument(node, {
        defaultSetup: pageSetup ?? layout?.page,
      });
      const golden = layoutGoldenFrom(name, result.lines, result.pagination);
      const out = path.join(dir, `${name}.layout-golden.json`);
      writeFileSync(out, JSON.stringify(golden, null, 2) + "\n");
      console.log("wrote", out);
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
