import { describe, expect, it } from "vitest";
import { isEditorMessage, isHostMessage } from "./protocol";

describe("isHostMessage", () => {
  it("accepts well-formed host messages", () => {
    expect(isHostMessage({ type: "te:load", plainText: "Q. Hi" })).toBe(true);
    expect(
      isHostMessage({ type: "te:load", plainText: "", editable: false }),
    ).toBe(true);
    expect(isHostMessage({ type: "te:get-plain", requestId: "1" })).toBe(true);
    expect(
      isHostMessage({ type: "te:get-plain", requestId: "1", mode: "original" }),
    ).toBe(true);
    expect(isHostMessage({ type: "te:set-plain", plainText: "x" })).toBe(true);
    expect(isHostMessage({ type: "te:set-editable", editable: true })).toBe(
      true,
    );
    expect(isHostMessage({ type: "te:focus" })).toBe(true);
    expect(
      isHostMessage({
        type: "te:load-template",
        styles: [{ id: "DepoQ", name: "Depo Q", type: "paragraph" }],
      }),
    ).toBe(true);
  });

  it("accepts DOCX, JSON and review messages", () => {
    const buffer = new ArrayBuffer(8);
    expect(isHostMessage({ type: "te:load-docx", buffer })).toBe(true);
    expect(isHostMessage({ type: "te:load-docx", buffer, editable: false })).toBe(true);
    expect(isHostMessage({ type: "te:get-docx", requestId: "a" })).toBe(true);
    expect(isHostMessage({ type: "te:get-json", requestId: "b" })).toBe(true);
    expect(isHostMessage({ type: "te:set-author", id: "u1", name: "Ann" })).toBe(true);
    expect(isHostMessage({ type: "te:set-track-changes", enabled: true })).toBe(true);
  });

  it("rejects malformed or unknown messages", () => {
    expect(isHostMessage(null)).toBe(false);
    expect(isHostMessage("te:focus")).toBe(false);
    expect(isHostMessage({ type: "te:load" })).toBe(false);
    expect(isHostMessage({ type: "te:load", plainText: "", editable: "no" })).toBe(
      false,
    );
    expect(isHostMessage({ type: "te:get-plain" })).toBe(false);
    expect(isHostMessage({ type: "te:get-plain", requestId: "1", mode: "all" })).toBe(false);
    expect(isHostMessage({ type: "te:load-docx", buffer: "PK.." })).toBe(false);
    expect(isHostMessage({ type: "te:get-docx" })).toBe(false);
    expect(isHostMessage({ type: "te:set-author", id: "", name: "x" })).toBe(false);
    expect(isHostMessage({ type: "te:set-track-changes", enabled: 1 })).toBe(false);
    expect(isHostMessage({ type: "te:ready" })).toBe(false);
  });
});

describe("isEditorMessage", () => {
  it("accepts well-formed editor messages", () => {
    expect(isEditorMessage({ type: "te:ready" })).toBe(true);
    expect(
      isEditorMessage({ type: "te:plain", requestId: "1", text: "abc" }),
    ).toBe(true);
    expect(
      isEditorMessage({ type: "te:json", requestId: "1", json: { type: "doc", content: [] } }),
    ).toBe(true);
    expect(
      isEditorMessage({ type: "te:docx", requestId: "1", buffer: new ArrayBuffer(4) }),
    ).toBe(true);
    expect(isEditorMessage({ type: "te:change", pageCount: 2, changeCount: 0 })).toBe(true);
    expect(isEditorMessage({ type: "te:layout", pageCount: 3 })).toBe(true);
    expect(isEditorMessage({ type: "te:action", name: "print" })).toBe(true);
    expect(isEditorMessage({ type: "te:error", message: "boom" })).toBe(true);
  });

  it("rejects malformed or unknown messages", () => {
    expect(isEditorMessage({ type: "te:plain", requestId: "1" })).toBe(false);
    expect(isEditorMessage({ type: "te:json", requestId: "1", json: {} })).toBe(false);
    expect(isEditorMessage({ type: "te:docx", requestId: "1", buffer: [] })).toBe(false);
    expect(isEditorMessage({ type: "te:change" })).toBe(false);
    expect(isEditorMessage({ type: "te:action", name: "delete" })).toBe(false);
    expect(isEditorMessage({ type: "te:load", plainText: "" })).toBe(false);
  });
});
