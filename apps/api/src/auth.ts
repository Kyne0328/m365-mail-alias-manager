import type { FastifyReply, FastifyRequest } from "fastify";
import { createRemoteJWKSet, jwtVerify, type JWTPayload } from "jose";
import { config } from "./config.js";

const jwks = createRemoteJWKSet(
  new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys")
);

const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export type AuthContext = {
  tenantId: string;
  userId: string;
  username?: string;
  name?: string;
};

declare module "fastify" {
  interface FastifyRequest {
    auth?: AuthContext;
  }
}

function claim(payload: JWTPayload, name: string): string | undefined {
  const value = payload[name];
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

export async function authenticate(request: FastifyRequest, reply: FastifyReply) {
  const header = request.headers.authorization;
  if (!header?.startsWith("Bearer ")) {
    return reply.code(401).send({ code: "UNAUTHORIZED", message: "A Microsoft access token is required." });
  }

  try {
    const token = header.slice("Bearer ".length);
    const result = await jwtVerify(token, jwks, {
      audience: config.apiAudience
    });

    const tenantId = claim(result.payload, "tid");
    const userId = claim(result.payload, "oid");
    const username =
      claim(result.payload, "preferred_username") ??
      claim(result.payload, "upn") ??
      claim(result.payload, "email");
    const scopes = new Set((claim(result.payload, "scp") ?? "").split(" ").filter(Boolean));
    const authorizedParty = claim(result.payload, "azp");

    if (!tenantId || !guidPattern.test(tenantId)) {
      throw new Error("Token has no valid tenant ID.");
    }

    if (claim(result.payload, "ver") !== "2.0") {
      throw new Error("Only Microsoft identity platform v2 access tokens are accepted.");
    }

    const expectedIssuer = `https://login.microsoftonline.com/${tenantId}/v2.0`;
    if (result.payload.iss !== expectedIssuer) {
      throw new Error("Token issuer does not match its tenant.");
    }

    if (authorizedParty !== config.clientId) {
      throw new Error("Token was acquired by an unauthorized client application.");
    }

    if (!userId || !guidPattern.test(userId)) {
      throw new Error("Token does not identify a Microsoft Entra user with an immutable object ID.");
    }

    if (!scopes.has(config.apiScope)) {
      throw new Error("Required API scope is missing.");
    }

    request.auth = {
      tenantId,
      userId,
      username: username?.toLowerCase(),
      name: claim(result.payload, "name")
    };
  } catch (error) {
    request.log.warn({ err: error }, "Rejected access token");
    return reply.code(401).send({ code: "UNAUTHORIZED", message: "Your Microsoft session is not valid for this API." });
  }
}

export function requireAuth(request: FastifyRequest): AuthContext {
  if (!request.auth) {
    throw new Error("Authentication context is unavailable.");
  }
  return request.auth;
}
