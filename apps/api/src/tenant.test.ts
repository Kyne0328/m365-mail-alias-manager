import { describe, expect, it, vi } from "vitest";
import { normalizeOrganization, verifyOrganization } from "./tenant.js";

describe("tenant organization verification", () => {
  it("requires the primary onmicrosoft.com shape", () => {
    expect(normalizeOrganization(" Contoso.onmicrosoft.com ")).toBe("contoso.onmicrosoft.com");
    expect(() => normalizeOrganization("contoso.com")).toThrow();
  });

  it("binds the supplied organization domain to the signed-in tenant id", async () => {
    const tenantId = "11111111-2222-4333-8444-555555555555";
    const fetcher = vi.fn(async () =>
      new Response(
        JSON.stringify({
          issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`
        }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      )
    ) as unknown as typeof fetch;

    await expect(
      verifyOrganization("tenant-a.onmicrosoft.com", tenantId, fetcher)
    ).resolves.toBeUndefined();

    await expect(
      verifyOrganization(
        "tenant-b.onmicrosoft.com",
        "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
        fetcher
      )
    ).rejects.toThrow("different Microsoft tenant");
  });
});
