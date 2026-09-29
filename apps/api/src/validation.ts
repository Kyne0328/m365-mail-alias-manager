const domainPattern = /^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i;
const prefixPattern = /^[a-z0-9](?:[a-z0-9._-]{0,22}[a-z0-9])?$/i;

export function normalizeDomain(value: unknown): string {
  const domain = String(value ?? "").trim().toLowerCase();
  if (!domainPattern.test(domain)) {
    throw new Error("Choose a valid accepted email domain.");
  }
  return domain;
}

export function normalizePrefix(value: unknown): string {
  const prefix = String(value ?? "").trim().toLowerCase();
  if (!prefixPattern.test(prefix)) {
    throw new Error("Prefix must be 1–24 characters using letters, numbers, dots, underscores, or hyphens.");
  }
  return prefix;
}

export function normalizeAddress(value: unknown): string {
  const address = String(value ?? "").trim().toLowerCase();
  if (address.length > 320 || !address.includes("@") || /[\s\r\n]/.test(address)) {
    throw new Error("Choose a valid alias address.");
  }
  return address;
}
