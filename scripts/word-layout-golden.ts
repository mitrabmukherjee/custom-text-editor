import { readFileSync, writeFileSync, readdirSync, existsSync } from "fs";
import path from "path";
import { importDocx } from "../lib/docx/import";
import { schema } from "../lib/pm/schema";
import { layoutDocument } from "../lib/pm/word-layout/layoutDocument";
import { layoutGoldenFrom } from "../lib/pm/word-layout/golden";

const root = process.cwd();
const dirs = [
  path.join(root, "public", "fixtures", "batch"),
  path.join(root, "public", "fixtures"),
];

async function main() {
  const seen = new Set<string>();
  for (const dir of dirs) {
    if (!existsSync(dir)) continue;
    for (const name of readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".docx"))) {
      if (seen.has(name)) continue;
      seen.add(name);
      const full = path.join(dir, name);
      const { doc, pageSetup, layout } = await importDocx(readFileSync(full));
      const node = schema.nodeFromJSON(doc);
      const result = layoutDocument(node, { defaultSetup: pageSetup ?? layout?.page });
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
