import { describe, expect, it } from "vitest";
import { docJsonToPlainText, plainTextToDoc, splitPlainParagraphs } from "./plainText";

describe("splitPlainParagraphs", () => {
  it("splits on blank lines", () => {
    expect(splitPlainParagraphs("a\n\nb")).toEqual(["a", "b"]);
  });

  it("returns empty for whitespace only", () => {
    expect(splitPlainParagraphs("  \n  ")).toEqual([]);
  });
});

describe("plainTextToDoc / docJsonToPlainText", () => {
  it("round-trips paragraphs", () => {
    const plain = "Q. First question?\n\nA. First answer.";
    const doc = plainTextToDoc(plain);
    expect(docJsonToPlainText(doc)).toBe(plain);
  });

  it("handles empty input as single empty paragraph", () => {
    const doc = plainTextToDoc("");
    expect(doc.type).toBe("doc");
    expect(docJsonToPlainText(doc)).toBe("");
  });

  it("treats single newlines as within-paragraph hard breaks when re-exported", () => {
    const plain = "Line one\nLine two";
    const doc = plainTextToDoc(plain);
    expect(docJsonToPlainText(doc)).toBe("Line one\nLine two");
  });
});
