import { describe, expect, it } from "vitest";
import { DEFAULT_PARENT_ORIGINS, parseAllowedOrigins } from "./origins";

describe("parseAllowedOrigins", () => {
  it("falls back to local dev origins when unset or blank", () => {
    expect(parseAllowedOrigins(undefined)).toEqual(DEFAULT_PARENT_ORIGINS);
    expect(parseAllowedOrigins("  , ")).toEqual(DEFAULT_PARENT_ORIGINS);
  });

  it("splits, trims and strips trailing slashes", () => {
    expect(
      parseAllowedOrigins(" https://crm.example.com/ ,http://localhost:3000"),
    ).toEqual(["https://crm.example.com", "http://localhost:3000"]);
  });
});
