import { describe, expect, it } from "vitest";
import { normalizeAddress, normalizeDomain, normalizePrefix } from "./validation.js";

describe("request validation", () => {
  it("normalizes domains and prefixes", () => {
    expect(normalizeDomain(" Example.COM ")).toBe("example.com");
    expect(normalizePrefix(" Temp_Box ")).toBe("temp_box");
  });

  it("rejects unsafe or ambiguous prefixes", () => {
    expect(() => normalizePrefix("-temp")).toThrow();
    expect(() => normalizePrefix("temp ")).not.toThrow();
    expect(() => normalizePrefix("two words")).toThrow();
    expect(() => normalizePrefix("a".repeat(25))).toThrow();
  });

  it("normalizes alias addresses without accepting whitespace", () => {
    expect(normalizeAddress(" TEMP-000001@Example.com ")).toBe("temp-000001@example.com");
    expect(() => normalizeAddress("bad alias@example.com")).toThrow();
  });
});
