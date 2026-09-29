const organizationPattern = /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.onmicrosoft\.com$/i;
const cache = new Map<string, { tenantId: string; expiresAt: number }>();
const CACHE_MS = 60 * 60 * 1000;

export function normalizeOrganization(value: unknown): string {
  const organization = String(value ?? "").trim().toLowerCase();
  if (!organizationPattern.test(organization)) {
    throw new Error("Use the tenant's primary domain, for example contoso.onmicrosoft.com.");
  }
  return organization;
}

export async function verifyOrganization(
  organization: string,
  expectedTenantId: string,
  fetcher: typeof fetch = fetch
): Promise<void> {
  const cached = cache.get(organization);
  if (cached && cached.expiresAt > Date.now()) {
    if (cached.tenantId !== expectedTenantId) {
      throw new Error("That organization domain belongs to a different Microsoft tenant.");
    }
    return;
  }

  const response = await fetcher(
    `https://login.microsoftonline.com/${encodeURIComponent(
      organization
    )}/v2.0/.well-known/openid-configuration`,
    { headers: { Accept: "application/json" } }
  );

  if (!response.ok) {
    throw new Error("Microsoft could not resolve that organization domain.");
  }

  const body = (await response.json()) as { issuer?: string };
  const match = body.issuer?.match(
    /^https:\/\/login\.microsoftonline\.com\/([0-9a-f-]{36})\/v2\.0$/i
  );

  if (!match) {
    throw new Error("Microsoft returned an unexpected tenant identity response.");
  }

  const tenantId = match[1].toLowerCase();
  cache.set(organization, { tenantId, expiresAt: Date.now() + CACHE_MS });

  if (tenantId !== expectedTenantId.toLowerCase()) {
    throw new Error("That organization domain belongs to a different Microsoft tenant.");
  }
}
