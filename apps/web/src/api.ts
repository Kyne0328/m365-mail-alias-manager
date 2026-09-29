import { acquireApiToken } from "./auth";

const baseUrl = import.meta.env.VITE_API_BASE_URL?.replace(/\/$/, "");

export function warmApi(): void {
  if (!baseUrl) return;

  void fetch(`${baseUrl}/health`, {
    method: "GET",
    mode: "cors",
    cache: "no-store"
  }).catch(() => {
    // Best-effort wake-up for scale-to-zero hosts. Normal API calls report real failures.
  });
}

export class ApiError extends Error {
  readonly status: number;
  readonly code?: string;

  constructor(message: string, status: number, code?: string) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
  }
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  if (!baseUrl) {
    throw new ApiError("The API URL is not configured.", 500, "CLIENT_CONFIG");
  }

  const auth = await acquireApiToken();
  const response = await fetch(`${baseUrl}${path}`, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${auth.accessToken}`,
      ...(init?.headers ?? {})
    }
  });

  const payload = (await response.json().catch(() => null)) as
    | { message?: string; code?: string }
    | T
    | null;

  if (!response.ok) {
    const problem = payload as { message?: string; code?: string } | null;
    throw new ApiError(
      problem?.message ?? "The server could not complete the request.",
      response.status,
      problem?.code
    );
  }

  return payload as T;
}

export type Session = {
  tenantId: string;
  userId: string;
  username?: string;
  name?: string;
};

export type DomainInfo = {
  domain: string;
  isDefault: boolean;
  type: string;
};

export type AliasInfo = {
  address: string;
  sequence: number;
  prefix: string;
  domain: string;
};

export type AliasSet = {
  mailbox: string;
  primaryAddress: string;
  aliases: AliasInfo[];
  count: number;
  limit: number;
  nextSequence: number;
};

export type WorkspaceSnapshot = {
  organization: string;
  domains: DomainInfo[];
  aliasSet: AliasSet;
};

export type OnboardingInfo = {
  adminConsentUrl: string;
  setupScript: string;
};

export const api = {
  session: () => request<Session>("/api/session"),
  bootstrap: () => request<WorkspaceSnapshot>("/api/bootstrap"),
  onboarding: () => request<OnboardingInfo>("/api/onboarding"),
  domains: (organization: string) =>
    request<{ domains: DomainInfo[] }>(
      `/api/domains?organization=${encodeURIComponent(organization)}`
    ),
  aliases: (organization: string) =>
    request<AliasSet>(
      `/api/aliases?organization=${encodeURIComponent(organization)}`
    ),
  createAlias: (organization: string, domain: string) =>
    request<AliasSet>("/api/aliases", {
      method: "POST",
      body: JSON.stringify({ organization, domain })
    }),
  deleteAlias: (organization: string, address: string) =>
    request<AliasSet>("/api/aliases", {
      method: "DELETE",
      body: JSON.stringify({ organization, address })
    })
};
