import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { collectComments, type CommentThread } from "@/lib/pm/comments";
import { PT_TO_PX, type DocumentLayout, type ParagraphFormat } from "@/lib/pm/format";
import type { DocJSON, NodeJSON } from "@/lib/pm/json";
import { LETTER_PAGE_SETUP } from "@/lib/pm/pagination/pageSetup";
import { collectChanges } from "@/lib/pm/track-changes/changes";
import { schema } from "@/lib/pm/schema";
import { docJsonToPlainText } from "@/lib/serializers/plainText";
import { exportDocx, importDocx } from "./index";

const DATE = "2026-10-13T09:30:00.000Z";
const PNG_1PX =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

function change(type: "insertion" | "deletion", id: string, name: string) {
  return { type, attrs: { id, authorId: name, authorName: name, date: DATE } };
}

const t = (text: string, marks?: NodeJSON["marks"]): NodeJSON =>
  marks ? { type: "text", text, marks } : { type: "text", text };

const SAMPLE: DocJSON = {
  type: "doc",
  content: [
    { type: "heading", attrs: { level: 1 }, content: [t("Deposition of Jane Doe")] },
    {
      type: "paragraph",
      content: [
        t("Q. Please state your "),
        t("full ", [change("insertion", "c1", "Alice")]),
        t("name", [{ type: "strong" }]),
        t(" for the record", [change("deletion", "c2", "Bob")]),
        t("."),
      ],
    },
    { type: "page_break" },
    {
      type: "paragraph",
      content: [t("A. Jane"), { type: "hard_break" }, t("Doe", [{ type: "em" }])],
    },
    {
      type: "bullet_list",
      content: [
        { type: "list_item", content: [{ type: "paragraph", content: [t("first")] }] },
        { type: "list_item", content: [{ type: "paragraph", content: [t("second")] }] },
      ],
    },
    {
      type: "ordered_list",
      attrs: { order: 1 },
      content: [
        { type: "list_item", content: [{ type: "paragraph", content: [t("one")] }] },
      ],
    },
  ],
};

async function documentXml(buffer: ArrayBuffer) {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file("word/document.xml")!.async("string")) ?? "";
}

const OOXML_SECTION_STARTS: Record<string, string> = {
  nextPage: "nextPage",
  continuous: "continuous",
  evenPage: "evenPage",
  oddPage: "oddPage",
};

function sectPrPageSizeKeyFromXml(sectPr: string): string {
  const pgSz = sectPr.match(/<w:pgSz[^/]*\/>/)?.[0] ?? "";
  const cols = sectPr.match(/<w:cols[^/]*\/>/)?.[0] ?? "";
  const w = pgSz.match(/w:w="([^"]+)"/)?.[1] ?? "";
  const h = pgSz.match(/w:h="([^"]+)"/)?.[1] ?? "";
  const colNum = cols.match(/w:num="(\d+)"/)?.[1] ?? "1";
  return `${w}|${h}|${colNum}`;
}

/** Inline `w:sectPr` start types in document order (matches `sectPrStartType` in import). */
function inlineSectPrStartTypes(xml: string): string[] {
  const body = xml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/)?.[1] ?? "";
  const starts: string[] = [];
  let prevKey: string | null = null;
  for (const para of body.matchAll(/<w:p[\s\S]*?<\/w:p>/g)) {
    const block = para[0];
    if (!block.includes("<w:sectPr")) continue;
    const sectPr = block.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0];
    if (!sectPr) continue;
    const key = sectPrPageSizeKeyFromXml(sectPr);
    const typeVal = sectPr.match(/<w:type w:val="([^"]+)"/)?.[1];
    let start: string;
    if (typeVal === "nextColumn") {
      const colNum = sectPr.match(/<w:cols[^>]*w:num="(\d+)"/)?.[1];
      start = !colNum || Number(colNum) <= 1 ? "continuous" : "nextPage";
    } else if (typeVal) {
      start = OOXML_SECTION_STARTS[typeVal] ?? "nextPage";
    } else if (prevKey === null || key === prevKey) {
      start = "continuous";
    } else {
      start = "nextPage";
    }
    starts.push(start);
    prevKey = key;
  }
  return starts;
}

function countImportedForcedBreaks(doc: DocJSON) {
  let pageBreaks = 0;
  let pageBreakBefore = 0;
  let sectionBreaks = 0;
  const sectionStarts: Record<string, number> = {};
  const walk = (nodes: NodeJSON[] | undefined) => {
    for (const node of nodes ?? []) {
      if (node.type === "page_break") pageBreaks += 1;
      if (node.type === "section_break") {
        sectionBreaks += 1;
        const start = (node.attrs?.start as string) ?? "nextPage";
        sectionStarts[start] = (sectionStarts[start] ?? 0) + 1;
      }
      if (node.type === "paragraph" || node.type === "heading") {
        const format = node.attrs?.format as ParagraphFormat | null | undefined;
        if (format?.pageBreakBefore) pageBreakBefore += 1;
      }
      walk(node.content);
    }
  };
  walk(doc.content);
  return { pageBreaks, pageBreakBefore, sectionBreaks, sectionStarts };
}

function sectionBreakStarts(doc: DocJSON): string[] {
  const out: string[] = [];
  const walk = (nodes: NodeJSON[] | undefined) => {
    for (const node of nodes ?? []) {
      if (node.type === "section_break") {
        out.push((node.attrs?.start as string) ?? "nextPage");
      }
      walk(node.content);
    }
  };
  walk(doc.content);
  return out;
}

describe("exportDocx", () => {
  it("writes Word revisions, page setup and a page-number footer", async () => {
    const buffer = await exportDocx(SAMPLE, LETTER_PAGE_SETUP, { trackRevisions: true });
    const xml = await documentXml(buffer);
    expect(xml).toMatch(/<w:ins [^>]*w:author="Alice"/);
    expect(xml).toMatch(/<w:del [^>]*w:author="Bob"/);
    expect(xml).toContain("<w:delText");
    expect(xml).toMatch(/<w:pgSz [^>]*w:w="12240"[^>]*w:h="15840"/);
    expect(xml).toMatch(/<w:pgMar [^>]*w:top="1440"/);
    expect(xml).toContain('w:type="page"');

    const zip = await JSZip.loadAsync(buffer);
    const footers = Object.keys(zip.files).filter((f) => /word\/footer\d*\.xml/.test(f));
    expect(footers.length).toBeGreaterThan(0);
    const footer = await zip.file(footers[0])!.async("string");
    expect(footer).toContain("PAGE");
    const settings = await zip.file("word/settings.xml")!.async("string");
    expect(settings).toContain("trackRevisions");
  });
});

describe("importDocx", () => {
  it("round-trips text, formatting, structure and tracked changes", async () => {
    const buffer = await exportDocx(SAMPLE, LETTER_PAGE_SETUP);
    const { doc, pageSetup } = await importDocx(buffer);

    expect(pageSetup).toMatchObject({
      width: 816,
      height: 1056,
      marginTop: 96,
      marginLeft: 96,
    });

    const node = schema.nodeFromJSON(doc);
    expect(node.child(0).type.name).toBe("heading");
    expect(node.child(2).type.name).toBe("page_break");
    expect(node.child(4).type.name).toBe("bullet_list");
    expect(node.child(4).childCount).toBe(2);
    expect(node.child(5).type.name).toBe("ordered_list");

    const changes = collectChanges(node);
    expect(changes.map((c) => [c.type, c.authorName, c.text])).toEqual([
      ["insertion", "Alice", "full "],
      ["deletion", "Bob", " for the record"],
    ]);
    expect(changes[0].date).toBe("2026-10-13T09:30:00Z");

    const bold = node.child(1).child(2);
    expect(bold.text).toBe("name");
    expect(bold.marks.map((m) => m.type.name)).toEqual(["strong"]);

    const accepted = docJsonToPlainText(doc, { mode: "accepted" });
    const original = docJsonToPlainText(SAMPLE, { mode: "accepted" });
    expect(accepted).toBe(original);
    expect(accepted).toContain("Q. Please state your full name.");
  });

  it("merges adjacent runs of one revision into a single change", async () => {
    const json: DocJSON = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            t("plain "),
            t("new ", [change("insertion", "x", "Alice")]),
            t("bold", [change("insertion", "x", "Alice"), { type: "strong" }]),
          ],
        },
      ],
    };
    const { doc } = await importDocx(await exportDocx(json, LETTER_PAGE_SETUP));
    const changes = collectChanges(schema.nodeFromJSON(doc));
    expect(changes).toHaveLength(1);
    expect(changes[0].text).toBe("new bold");
  });

  it("round-trips formatting revisions (w:rPrChange)", async () => {
    const formatChange = {
      type: "formatChange",
      attrs: {
        id: "f1",
        authorId: "QC",
        authorName: "QC",
        date: DATE,
        description: "Font: Italic",
        old: { bold: false, italic: false, fontFamily: null, fontSize: null },
      },
    };
    const json: DocJSON = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [t("plain "), t("sorry", [{ type: "em" }, formatChange])],
        },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    expect(await documentXml(buffer)).toMatch(/<w:rPrChange [^>]*w:author="QC"/);

    const { doc } = await importDocx(buffer);
    const changes = collectChanges(schema.nodeFromJSON(doc));
    expect(changes.map((c) => [c.type, c.authorName, c.text, c.description])).toEqual([
      ["format", "QC", "sorry", "Font: Italic"],
    ]);
  });

  it("round-trips underline, strikethrough, colour, highlight, caps and super/subscript", async () => {
    const json: DocJSON = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            t("under", [{ type: "underline", attrs: { style: "double" } }]),
            t("struck", [{ type: "strike", attrs: { style: "single" } }]),
            t("red", [{ type: "textColor", attrs: { color: "#ff0000" } }]),
            t("marked", [{ type: "highlight", attrs: { color: "#ffff00" } }]),
            t("shaded", [{ type: "highlight", attrs: { color: "#123456" } }]),
            t("caps", [{ type: "caps", attrs: { style: "small" } }]),
            t("2", [{ type: "vertAlign", attrs: { align: "superscript" } }]),
          ],
        },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    const xml = await documentXml(buffer);
    expect(xml).toMatch(/<w:u w:val="double"/);
    expect(xml).toContain("<w:strike/>");
    expect(xml).toMatch(/<w:color w:val="FF0000"/);
    expect(xml).toMatch(/<w:highlight w:val="yellow"/);
    expect(xml).toMatch(/<w:shd [^>]*w:fill="123456"/);
    expect(xml).toContain("<w:smallCaps/>");
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"/);

    const { doc } = await importDocx(buffer);
    const para = schema.nodeFromJSON(doc).child(0);
    const runs: [string, string, unknown][] = [];
    para.forEach((n) => {
      for (const m of n.marks) runs.push([n.text ?? "", m.type.name, m.attrs.style ?? m.attrs.color ?? m.attrs.align]);
    });
    expect(runs).toEqual([
      ["under", "underline", "double"],
      ["struck", "strike", "single"],
      ["red", "textColor", "#ff0000"],
      ["marked", "highlight", "#ffff00"],
      ["shaded", "highlight", "#123456"],
      ["caps", "caps", "small"],
      ["2", "vertAlign", "superscript"],
    ]);
  });

  it("round-trips formatting revisions of the new run properties", async () => {
    const json: DocJSON = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            t("now red", [
              { type: "textColor", attrs: { color: "#ff0000" } },
              {
                type: "formatChange",
                attrs: {
                  id: "f2",
                  authorId: "QC",
                  authorName: "QC",
                  date: DATE,
                  description: "Font: Font color: Red",
                  old: { bold: false, italic: false, fontFamily: null, fontSize: null, underline: "single" },
                },
              },
            ]),
          ],
        },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    expect(await documentXml(buffer)).toMatch(/<w:rPrChange [^>]*>.*<w:u w:val="single"/s);
    const { doc } = await importDocx(buffer);
    const [c] = collectChanges(schema.nodeFromJSON(doc));
    expect(c.description).toBe("Font: No underline, Font color: Red");
  });

  it("round-trips sections, paragraph revisions, contextual spacing and page-break-before", async () => {
    const firstSection = {
      page: { ...LETTER_PAGE_SETUP, marginTop: 144 },
      titlePage: true,
      headers: { default: [{ format: {}, runs: [{ kind: "text" as const, text: "Cover header" }] }], first: [] },
    };
    const layout: DocumentLayout = {
      page: LETTER_PAGE_SETUP,
      evenAndOddHeaders: true,
      headers: {
        default: [{ format: {}, runs: [{ kind: "text", text: "Odd header" }] }],
        even: [{ format: {}, runs: [{ kind: "text", text: "Even header" }] }],
      },
      pageNumberStart: 1,
    };
    const json: DocJSON = {
      type: "doc",
      attrs: { layout },
      content: [
        { type: "paragraph", content: [t("Cover")] },
        { type: "section_break", attrs: { start: "oddPage", section: firstSection } },
        {
          type: "paragraph",
          attrs: {
            format: { align: "center", contextualSpacing: true, pageBreakBefore: true },
            paraChange: {
              id: "p1",
              authorId: "QC",
              authorName: "QC",
              date: DATE,
              old: { contextualSpacing: true, pageBreakBefore: true },
              description: "Centered",
            },
          },
          content: [t("Body")],
        },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    const xml = await documentXml(buffer);
    expect(xml).toMatch(/<w:pPrChange [^>]*w:author="QC"/);
    expect(xml).toContain("<w:contextualSpacing/>");
    expect(xml).toContain("<w:pageBreakBefore/>");
    expect(xml.match(/<w:sectPr/g)).toHaveLength(2);
    expect(xml).toMatch(/<w:type w:val="oddPage"/);

    const { doc, layout: imported } = await importDocx(buffer);
    const node = schema.nodeFromJSON(doc);
    expect(node.child(1).type.name).toBe("section_break");
    expect(node.child(1).attrs.start).toBe("oddPage");
    const section = node.child(1).attrs.section;
    expect(section.page.marginTop).toBe(144);
    expect(section.titlePage).toBe(true);
    expect(section.headers.first).toEqual([]);
    expect(section.headers.default[0].runs).toEqual([{ kind: "text", text: "Cover header" }]);

    const body = node.child(2);
    expect(body.attrs.format).toMatchObject({ align: "center", contextualSpacing: true, pageBreakBefore: true });
    expect(collectChanges(node)).toMatchObject([
      { type: "paragraph", authorName: "QC", description: "Centered" },
    ]);

    expect(imported!.evenAndOddHeaders).toBe(true);
    expect(imported!.pageNumberStart).toBe(1);
    expect(imported!.headers!.even![0].runs).toEqual([{ kind: "text", text: "Even header" }]);
  });

  it("round-trips framed (floating) header paragraphs", async () => {
    const layout: DocumentLayout = {
      page: LETTER_PAGE_SETUP,
      headers: {
        default: [
          { format: {}, runs: [{ kind: "text", text: "Flow" }] },
          {
            format: {},
            runs: [{ kind: "field", field: "PAGE" }],
            floating: true,
            frame: { hAnchor: "page", vAnchor: "page", xAlign: "right", yAlign: "top", width: 96 },
          },
        ],
      },
    };
    const json: DocJSON = { type: "doc", attrs: { layout }, content: [{ type: "paragraph", content: [t("x")] }] };
    const { layout: imported } = await importDocx(await exportDocx(json, LETTER_PAGE_SETUP));
    const [flow, framed] = imported!.headers!.default!;
    expect(flow.frame).toBeUndefined();
    expect(framed).toMatchObject({
      floating: true,
      frame: { hAnchor: "page", vAnchor: "page", xAlign: "right", yAlign: "top", width: 96 },
    });
  });

  it("reads custom page sizes", async () => {
    const a4 = { ...LETTER_PAGE_SETUP, width: 793.73, height: 1122.53, marginLeft: 120 };
    const { pageSetup } = await importDocx(await exportDocx(SAMPLE, a4));
    expect(pageSetup!.width).toBeCloseTo(793.73, 0);
    expect(pageSetup!.height).toBeCloseTo(1122.53, 0);
    expect(pageSetup!.marginLeft).toBe(120);
  });

  it("round-trips Word paragraph formatting, tabs, fonts and section layout", async () => {
    const speaker: ParagraphFormat = {
      align: "justify",
      indentLeft: 168,
      indentFirstLine: -120,
      spaceAfter: 100 / 15,
      lineRule: "auto",
      lineValue: 1.5,
      fontFamily: "Arial",
      fontSize: 11.5,
      keepNext: true,
    };
    const layout: DocumentLayout = {
      page: { ...LETTER_PAGE_SETUP, marginTop: 72, marginLeft: 120, marginRight: 120, marginBottom: 120 },
      headerDistance: 48,
      footerDistance: 48,
      titlePage: true,
      lineNumbers: { countBy: 1, start: 1, restart: "newPage", distance: 24 },
      footers: {
        default: [
          {
            format: { fontFamily: "Arial", fontSize: 10, lineRule: "auto", lineValue: 1 },
            runs: [
              { kind: "field", field: "PAGE" },
              { kind: "text", text: " of " },
              { kind: "field", field: "NUMPAGES" },
            ],
          },
        ],
        first: [],
      },
    };
    const json: DocJSON = {
      type: "doc",
      attrs: { layout },
      content: [
        {
          type: "paragraph",
          attrs: { format: speaker },
          content: [
            t("OFFICER:"),
            { type: "tab" },
            t("Come in", [{ type: "textStyle", attrs: { fontFamily: "Courier New", fontSize: null } }]),
          ],
        },
      ],
    };
    const buffer = await exportDocx(json, layout.page!);
    const { doc, layout: imported } = await importDocx(buffer);

    const para = schema.nodeFromJSON(doc).child(0);
    const format = para.attrs.format as ParagraphFormat;
    expect(format).toMatchObject({
      align: "justify",
      indentLeft: 168,
      indentFirstLine: -120,
      lineRule: "auto",
      lineValue: 1.5,
      fontFamily: "Arial",
      fontSize: 11.5,
      keepNext: true,
    });
    expect(format.spaceAfter).toBeCloseTo(100 / 15);
    expect(para.child(1).type.name).toBe("tab");
    expect(para.child(2).marks.map((m) => [m.type.name, m.attrs.fontFamily])).toEqual([
      ["textStyle", "Courier New"],
    ]);
    expect(docJsonToPlainText(doc)).toBe("OFFICER:\tCome in");

    expect(imported).toMatchObject({
      headerDistance: 48,
      footerDistance: 48,
      titlePage: true,
      lineNumbers: { countBy: 1, start: 1, restart: "newPage", distance: 24 },
      justifyShrink: 0.25,
    });
    expect(imported!.footers!.first).toEqual([]);
    expect(imported!.footers!.default![0].runs).toEqual([
      { kind: "field", field: "PAGE" },
      { kind: "text", text: " of " },
      { kind: "field", field: "NUMPAGES" },
    ]);
  });

  it("round-trips styles and multilevel, restarted and continued lists", async () => {
    const item = (text: string, nested?: NodeJSON): NodeJSON => ({
      type: "list_item",
      content: [{ type: "paragraph", content: [t(text)] }, ...(nested ? [nested] : [])],
    });
    const doc: DocJSON = {
      type: "doc",
      attrs: {
        styles: [
          { id: "Quote", name: "Quote", type: "paragraph", format: { align: "center" }, run: { italic: true } },
        ],
      },
      content: [
        {
          type: "heading",
          attrs: { level: 2, format: { styleId: "Heading2", fontSize: 13, keepNext: true } },
          content: [t("Exhibits", [{ type: "strong" }])],
        },
        {
          type: "paragraph",
          attrs: { format: { styleId: "Quote", align: "center" } },
          content: [t("A quotation", [{ type: "em" }])],
        },
        {
          type: "ordered_list",
          attrs: { order: 1, numFormat: "legal" },
          content: [item("one", { type: "ordered_list", content: [item("one.one")] }), item("two")],
        },
        { type: "paragraph", content: [t("between")] },
        {
          type: "ordered_list",
          attrs: { order: 3, numFormat: "upperRoman" },
          content: [item("three")],
        },
        { type: "bullet_list", attrs: { bullet: "dash" }, content: [item("dash")] },
      ],
    };
    const buffer = await exportDocx(doc, LETTER_PAGE_SETUP);
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = await zip.file("word/styles.xml")!.async("string");
    expect(stylesXml).toContain('w:styleId="Quote"');
    expect(stylesXml).not.toContain("2E74B5");
    const xml = await documentXml(buffer);
    expect(xml).toContain('<w:pStyle w:val="Quote"/>');

    const { doc: back } = await importDocx(buffer);
    const node = schema.nodeFromJSON(back);
    expect(node.child(0).type.name).toBe("heading");
    expect(node.child(0).attrs.level).toBe(2);
    expect(node.child(1).attrs.format.styleId).toBe("Quote");
    const legal = node.child(2);
    expect(legal.attrs).toMatchObject({ numFormat: "legal", order: 1 });
    expect(legal.child(0).child(1).type.name).toBe("ordered_list");
    expect(node.child(4).attrs).toMatchObject({ numFormat: "upperRoman", order: 3 });
    expect(node.child(5).attrs.bullet).toBe("dash");
  });

  it("round-trips comments with replies, resolved state and multi-paragraph ranges", async () => {
    const mark = (id: string, authorId: string) => ({ type: "comment", attrs: { id, authorId } });
    const json: DocJSON = {
      type: "doc",
      attrs: {
        comments: {
          k1: {
            id: "k1",
            authorId: "Ann",
            authorName: "Ann Lee",
            date: DATE,
            text: "Check spelling\nand the date",
            resolved: true,
            replies: [{ id: "r1", authorId: "Bob", authorName: "Bob", date: DATE, text: "Done" }],
          },
          k2: { id: "k2", authorId: "Bob", authorName: "Bob", date: DATE, text: "Spans paragraphs" },
        },
      },
      content: [
        {
          type: "paragraph",
          content: [t("The "), t("witness", [mark("k1", "Ann")]), t(" said "), t("yes.", [mark("k2", "Bob")])],
        },
        { type: "paragraph", content: [t("Then left.", [mark("k2", "Bob")]), t(" End.")] },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await documentXml(buffer);
    expect(xml.match(/<w:commentRangeStart /g)).toHaveLength(3);
    expect(xml.match(/<w:commentReference /g)).toHaveLength(3);
    expect(await zip.file("word/comments.xml")!.async("string")).toContain('w:author="Ann Lee"');
    expect(zip.file("word/commentsExtended.xml")).toBeTruthy();

    const { doc } = await importDocx(buffer);
    const threads = doc.attrs?.comments as Record<string, CommentThread>;
    const list = Object.values(threads).sort((a, b) => a.text.localeCompare(b.text));
    expect(list.map((c) => [c.authorName, c.text, !!c.resolved])).toEqual([
      ["Ann Lee", "Check spelling\nand the date", true],
      ["Bob", "Spans paragraphs", false],
    ]);
    expect(list[0].replies?.map((r) => [r.authorName, r.text])).toEqual([["Bob", "Done"]]);
    const anchors = collectComments(schema.nodeFromJSON(doc));
    expect(anchors.map((c) => c.quote)).toEqual(["witness", "yes.Then left."]);
  });

  it("round-trips tables with merged cells, hyperlinks and pictures", async () => {
    const cell = (text: string, attrs: Record<string, unknown> = {}): NodeJSON => ({
      type: "table_cell",
      attrs: { colwidth: [200], ...attrs },
      content: [{ type: "paragraph", content: text ? [t(text)] : [] }],
    });
    const link = (href: string) => [{ type: "link", attrs: { href, title: null } }];
    const json: DocJSON = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [t("See "), t("the site", link("https://example.com/a?b=1")), t(" or "), t("intro", link("#intro"))],
        },
        {
          type: "table",
          attrs: { borders: true },
          content: [
            { type: "table_row", content: [cell("Name", { colspan: 2, colwidth: [200, 200], background: "#d9e2f3" }), cell("Age")] },
            { type: "table_row", content: [cell("Ann", { rowspan: 2 }), cell("Lee"), cell("40", { valign: "middle" })] },
            { type: "table_row", content: [cell("Ng"), cell("33")] },
          ],
        },
        { type: "table", attrs: { borders: false }, content: [{ type: "table_row", content: [cell("x"), cell("y")] }] },
        {
          type: "paragraph",
          content: [{ type: "image", attrs: { src: `data:image/png;base64,${PNG_1PX}`, alt: "dot", width: 40, height: 20 } }],
        },
      ],
    };
    const buffer = await exportDocx(json, LETTER_PAGE_SETUP);
    const xml = await documentXml(buffer);
    expect(xml).toContain("<w:hyperlink");
    expect(xml).toContain('w:anchor="intro"');
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain("<w:drawing>");
    const zip = await JSZip.loadAsync(buffer);
    expect(Object.keys(zip.files).some((f) => f.startsWith("word/media/"))).toBe(true);

    const { doc } = await importDocx(buffer);
    const node = schema.nodeFromJSON(doc);
    const links: string[] = [];
    node.descendants((n) => {
      const href = n.marks.find((m) => m.type.name === "link")?.attrs.href;
      if (href && n.isText) links.push(`${n.text}=${href}`);
    });
    expect(links).toEqual(["the site=https://example.com/a?b=1", "intro=#intro"]);

    const table = node.child(1);
    expect(table.type.name).toBe("table");
    expect(table.attrs.borders).toBe(true);
    expect(table.childCount).toBe(3);
    const head = table.child(0).child(0);
    expect(head.attrs).toMatchObject({ colspan: 2, colwidth: [200, 200], background: "#d9e2f3" });
    expect(table.child(1).child(0).attrs.rowspan).toBe(2);
    expect(table.child(1).child(2).attrs.valign).toBe("middle");
    expect(table.child(2).childCount).toBe(2);
    expect(node.child(2).attrs.borders).toBe(false);

    let image: Record<string, unknown> | null = null;
    node.descendants((n) => {
      if (n.type.name === "image") image = n.attrs;
    });
    expect(image).toMatchObject({ width: 40, height: 20, alt: "dot" });
    expect(String(image!["src"])).toMatch(/^data:image\/png;base64,/);
    expect(docJsonToPlainText(doc)).toContain("Name\tAge\n\nAnn\tLee\t40\n\nNg\t33\n\nx\ty");
  });

  it("imports HYPERLINK fields and vertically merged, shaded table cells", async () => {
    const base = await exportDocx({ type: "doc", content: [{ type: "paragraph" }] }, LETTER_PAGE_SETUP);
    const zip = await JSZip.loadAsync(base);
    const xml = await zip.file("word/document.xml")!.async("string");
    const body = [
      "<w:p><w:r><w:t xml:space=\"preserve\">Go </w:t></w:r>",
      "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
      "<w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"https://word.example/\" </w:instrText></w:r>",
      "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>",
      "<w:r><w:t>here</w:t></w:r>",
      "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
      "<w:r><w:t>.</w:t></w:r></w:p>",
      "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>",
      "<w:tblGrid><w:gridCol w:w=\"1500\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>",
      "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/><w:shd w:val=\"clear\" w:fill=\"FFFF00\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>",
      "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>",
      "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>",
      "<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>",
      "</w:tbl><w:p/>",
    ].join("");
    zip.file("word/document.xml", xml.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const node = schema.nodeFromJSON(doc);

    const para = node.child(0);
    expect(para.textContent).toBe("Go here.");
    const linked = para.child(1);
    expect(linked.text).toBe("here");
    expect(linked.marks.find((m) => m.type.name === "link")?.attrs.href).toBe("https://word.example/");

    const table = node.child(1);
    expect(table.type.name).toBe("table");
    expect(table.child(0).child(0).attrs).toMatchObject({ rowspan: 2, colwidth: [100], background: "#ffff00" });
    expect(table.child(1).childCount).toBe(1);
    expect(table.child(1).child(0).textContent).toBe("C");
  });

  it("rejects files that are not DOCX", async () => {
    const zip = new JSZip();
    zip.file("hello.txt", "hi");
    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    await expect(importDocx(buffer)).rejects.toThrow(/word\/document\.xml/);
  });

  it("imports inline sectPr w:type as section_break start (continuous, not next section)", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
        "<w:p>",
        "<w:pPr><w:sectPr>",
        '<w:type w:val="continuous"/>',
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr></w:pPr>",
        "<w:r><w:t>Section zero</w:t></w:r>",
        "</w:p>",
        "<w:p><w:r><w:t>Section one</w:t></w:r></w:p>",
        "<w:sectPr>",
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr>",
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const node = schema.nodeFromJSON(doc);
    expect(node.child(1).type.name).toBe("section_break");
    expect(node.child(1).attrs.start).toBe("continuous");
  });

  it("imports test.docx section break types without shifted starts", async () => {
    const buffer = await readFile(join(process.cwd(), "public", "test.docx"));
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    const xml = await documentXml(arrayBuffer);
    const expected = inlineSectPrStartTypes(xml);
    const { doc } = await importDocx(buffer);
    const imported = sectionBreakStarts(doc);
    expect(imported.length).toBe(expected.length);
    expect(imported).toEqual(expected);
    expect(countImportedForcedBreaks(doc).sectionBreaks).toBe(expected.length);
  });

  it("treats missing w:type as continuous when page geometry is unchanged", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    const pg =
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
        `<w:p><w:pPr><w:sectPr>${pg}</w:sectPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>`,
        `<w:p><w:pPr><w:sectPr>${pg}<w:pgNumType w:start="1"/></w:sectPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p>`,
        `<w:p><w:r><w:t>C</w:t></w:r></w:p>`,
        `<w:sectPr>${pg}</w:sectPr>`,
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    expect(sectionBreakStarts(doc)).toEqual(["continuous", "continuous"]);
  });

  it("imports sectPr in a table cell with continuous start", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
        "<w:tbl><w:tr><w:tc>",
        "<w:p>",
        "<w:pPr><w:sectPr>",
        '<w:type w:val="continuous"/>',
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr></w:pPr>",
        "<w:r><w:t>In table</w:t></w:r>",
        "</w:p>",
        "</w:tc></w:tr></w:tbl>",
        "<w:p><w:r><w:t>After table</w:t></w:r></w:p>",
        "<w:sectPr>",
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr>",
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const node = schema.nodeFromJSON(doc);
    const table = node.child(0);
    expect(table.type.name).toBe("table");
    expect(table.child(0).child(0).child(1).type.name).toBe("section_break");
    expect(table.child(0).child(0).child(1).attrs.start).toBe("continuous");
  });

  it("maps nextColumn to continuous for single-column sectPr", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
        "<w:p>",
        "<w:pPr><w:sectPr>",
        '<w:type w:val="nextColumn"/><w:cols w:num="1"/>',
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr></w:pPr>",
        "<w:r><w:t>One</w:t></w:r>",
        "</w:p>",
        "<w:p><w:r><w:t>Two</w:t></w:r></w:p>",
        "<w:sectPr>",
        '<w:pgSz w:w="12240" w:h="15840"/>',
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
        "</w:sectPr>",
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    expect(schema.nodeFromJSON(doc).child(1).attrs.start).toBe("continuous");
  });

  it("imports paragraph box borders and centered table width", async () => {
    const zip = await JSZip.loadAsync(await readFile(join(process.cwd(), "public", "test.docx")));
    const xml = await zip.file("word/document.xml")!.async("string");
    const body = [
      "<w:p><w:pPr>",
      '<w:pBdr><w:top w:val="dashed" w:sz="12" w:space="4" w:color="000000"/></w:pBdr>',
      '<w:jc w:val="center"/>',
      "</w:pPr><w:r><w:t>FORM 3</w:t></w:r></w:p>",
      "<w:tbl>",
      "<w:tblPr>",
      '<w:tblW w:w="3000" w:type="dxa"/>',
      '<w:jc w:val="center"/>',
      '<w:tblInd w:w="720" w:type="dxa"/>',
      "</w:tblPr>",
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>',
      "<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>",
      "</w:tbl>",
    ].join("");
    zip.file("word/document.xml", xml.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const node = schema.nodeFromJSON(doc);
    expect(node.child(0).attrs.format?.boxBorder?.top).toMatchObject({
      style: "dashed",
      widthPt: 1.5,
      paddingPx: (4 / 8) * PT_TO_PX,
    });
    const table = node.child(1);
    expect(table.attrs).toMatchObject({ align: "center", width: 200, indent: 48 });
  });

  it("imports per-cell tcBorders", async () => {
    const zip = await JSZip.loadAsync(await readFile(join(process.cwd(), "public", "test.docx")));
    const xml = await zip.file("word/document.xml")!.async("string");
    const body = [
      "<w:tbl>",
      "<w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>",
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>',
      "<w:tr><w:tc><w:tcPr>",
      "<w:tcBorders>",
      '<w:bottom w:val="dotted" w:sz="8" w:space="0" w:color="000000"/>',
      '<w:top w:val="nil"/>',
      "</w:tcBorders>",
      "</w:tcPr><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:tc></w:tr>",
      "</w:tbl>",
    ].join("");
    zip.file("word/document.xml", xml.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const cell = schema.nodeFromJSON(doc).child(0).child(0).child(0);
    expect(cell.attrs.edges?.bottom).toMatchObject({ style: "dotted", widthPt: 1 });
    expect(cell.attrs.edges?.top).toBe("none");
  });

  it("imports batch layout fixtures (form2 paragraph, form2 table, form3 table)", async () => {
    const dir = join(process.cwd(), "public", "fixtures", "batch");
    const { doc: p } = await importDocx(await readFile(join(dir, "form2-paragraph.docx")));
    const para = schema.nodeFromJSON(p);
    expect(para.textContent).toContain("FORM 2");
    expect(para.child(0).attrs.format?.align).toBe("center");

    const { doc: t } = await importDocx(await readFile(join(dir, "form2-table.docx")));
    const table = schema.nodeFromJSON(t).child(0);
    expect(table.type.name).toBe("table");
    expect(table.attrs.align).toBe("center");
    expect(table.attrs.width).toBe(200);
    expect(table.child(0).child(0).attrs.edges?.bottom).toBe("none");

    const { doc: f3 } = await importDocx(await readFile(join(dir, "form3-table.docx")));
    const f3table = schema.nodeFromJSON(f3).child(0);
    expect(f3table.textContent).toContain("FORM 3");
    expect(f3table.child(0).attrs.rowHeight).toBeGreaterThan(0);
    expect(f3table.child(0).child(0).attrs.cellMargin?.left).toBeGreaterThan(0);
    expect(f3table.child(0).child(0).attrs.edges?.bottom).toMatchObject({ style: "dotted" });
  });

  it("round-trips table layout attrs", async () => {
    const dir = join(process.cwd(), "public", "fixtures", "batch");
    const { doc } = await importDocx(await readFile(join(dir, "form3-table.docx")));
    const reimported = schema.nodeFromJSON((await importDocx(await exportDocx(doc, LETTER_PAGE_SETUP))).doc);
    const table = reimported.child(0);
    expect(table.attrs.align).toBe("center");
    expect(table.child(0).child(0).attrs.edges?.bottom).toMatchObject({ style: "dotted" });
  });

  it("round-trips paragraph box border w:space", async () => {
    const zip = await JSZip.loadAsync(await readFile(join(process.cwd(), "public", "test.docx")));
    const xml = await zip.file("word/document.xml")!.async("string");
    const body = [
      "<w:p><w:pPr>",
      '<w:pBdr><w:top w:val="dashed" w:sz="12" w:space="4" w:color="000000"/></w:pBdr>',
      "</w:pPr><w:r><w:t>Bordered</w:t></w:r></w:p>",
    ].join("");
    zip.file("word/document.xml", xml.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const exported = await exportDocx(doc, LETTER_PAGE_SETUP);
    const out = await documentXml(exported);
    expect(out).toMatch(/<w:top[^>]*w:space="4"/);
  });

  it("imports VML pict when imagedata uses a plain id attribute", async () => {
    const sectPr =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="png" ContentType="image/png"/>',
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
        "</Types>",
      ].join(""),
    );
    zip.file(
      "word/_rels/document.xml.rels",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/dot.png"/>',
        "</Relationships>",
      ].join(""),
    );
    zip.file("word/media/dot.png", Buffer.from(PNG_1PX, "base64"));
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">',
        "<w:body>",
        "<w:p><w:r><w:pict>",
        '<v:shape style="width:40pt;height:20pt"><v:imagedata id="rId5" o:title="dot"/></v:shape>',
        "</w:pict></w:r></w:p>",
        sectPr,
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const para = schema.nodeFromJSON(doc).child(0);
    expect(para.child(0).type.name).toBe("image");
    expect(para.child(0).attrs.alt).toBe("dot");
    expect(String(para.child(0).attrs.src)).toContain("image/png");
  });

  it("imports and exports table row trHeight hRule", async () => {
    const sectPr =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      "word/document.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
        "<w:tbl>",
        '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>',
        '<w:tr><w:trPr><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr>',
        "<w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr>",
        "</w:tbl>",
        sectPr,
        "</w:body>",
        "</w:document>",
      ].join(""),
    );
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const row = schema.nodeFromJSON(doc).child(0).child(0);
    expect(row.attrs.rowHeightRule).toBe("atLeast");
    expect(row.attrs.rowHeight).toBeGreaterThan(0);
    const exported = await documentXml(await exportDocx(doc, LETTER_PAGE_SETUP));
    expect(exported).toMatch(/w:trHeight[^>]*w:hRule="atLeast"/);
  });

  it("imports Zimmerman Form 3 (header table, Courier cs font, anchored VML signature)", async () => {
    const path = join(process.cwd(), "public", "fixtures", "zimmerman-form3.docx");
    if (!existsSync(path)) return;
    const { doc } = await importDocx(await readFile(path));
    const root = schema.nodeFromJSON(doc);
    let form3Format: ParagraphFormat | undefined;
    let anchoredOnUnderline = false;
    root.descendants((n) => {
      if (n.type.name === "paragraph" && n.textContent.trim() === "FORM 3") {
        form3Format = n.attrs.format as ParagraphFormat;
      }
      if (n.type.name === "paragraph" && /_{8,}/.test(n.textContent)) {
        const img = n.firstChild;
        if (img?.type.name === "image" && img.attrs.anchorX != null) anchoredOnUnderline = true;
      }
    });
    expect(form3Format?.fontFamily).toBe("Courier New");
    expect(anchoredOnUnderline).toBe(true);
    expect(root.textContent).toContain("Jennifer O'Hare");
  });

  it("imports body framePr on paragraphs", async () => {
    const zip = await JSZip.loadAsync(await readFile(join(process.cwd(), "public", "test.docx")));
    const xml = await zip.file("word/document.xml")!.async("string");
    const body =
      '<w:p><w:pPr><w:framePr w:hAnchor="margin" w:vAnchor="text" w:x="1440" w:y="-200"/></w:pPr>' +
      "<w:r><w:t>Signature</w:t></w:r></w:p>";
    zip.file("word/document.xml", xml.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
    const { doc } = await importDocx(await zip.generateAsync({ type: "arraybuffer" }));
    const node = schema.nodeFromJSON(doc).child(0);
    expect(node.attrs.floating).toBe(true);
    expect(node.attrs.frame?.hAnchor).toBe("margin");
  });
});
