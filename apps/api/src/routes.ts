import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { authenticate, requireAuth } from "./auth.js";
import { config } from "./config.js";
import {
  type AliasSet,
  type ExchangeDomain,
  type WorkspaceSnapshot,
  ExchangeError,
  runExchange
} from "./exchange.js";
import { buildAdminConsentUrl, buildSetupScript } from "./onboarding.js";
import { normalizeOrganization, verifyOrganization } from "./tenant.js";
import { normalizeAddress, normalizeDomain } from "./validation.js";
import { withKeyedLock } from "./mutex.js";

function sendInputError(reply: FastifyReply, error: unknown) {
  return reply.code(400).send({
    code: "INVALID_INPUT",
    message: error instanceof Error ? error.message : "The request is invalid."
  });
}

function sendExchangeError(reply: FastifyReply, error: unknown) {
  if (!(error instanceof ExchangeError)) {
    return reply.code(500).send({
      code: "SERVER_ERROR",
      message: "The server could not complete the Exchange request."
    });
  }

  if (error.category === "authorization" || error.category === "connection") {
    return reply.code(403).send({
      code: "EXCHANGE_NOT_READY",
      message:
        "This Microsoft 365 tenant has not finished the Mail Alias Manager setup, or a required permission assignment is still propagating."
    });
  }

  if (error.category === "mailbox") {
    return reply.code(404).send({
      code: "MAILBOX_NOT_FOUND",
      message: "The signed-in account does not have an Exchange Online mailbox that this app can manage."
    });
  }

  if (error.category === "domain") {
    return reply.code(400).send({
      code: "DOMAIN_NOT_ACCEPTED",
      message: error.message
    });
  }

  return reply.code(502).send({
    code: "EXCHANGE_ERROR",
    message: error.message || "Exchange Online could not complete the request."
  });
}

async function verifiedOrganization(
  request: FastifyRequest,
  raw: unknown
): Promise<string> {
  const auth = requireAuth(request);
  const organization = normalizeOrganization(raw);
  await verifyOrganization(organization, auth.tenantId);
  return organization;
}

export async function registerRoutes(app: FastifyInstance) {
  app.get("/api/session", { preHandler: authenticate }, async (request) => {
    const auth = requireAuth(request);
    return {
      tenantId: auth.tenantId,
      userId: auth.userId,
      username: auth.username,
      name: auth.name
    };
  });

  app.get("/api/bootstrap", { preHandler: authenticate }, async (request, reply) => {
    try {
      const auth = requireAuth(request);
      return await runExchange<WorkspaceSnapshot>({
        action: "bootstrap",
        tenantId: auth.tenantId,
        userId: auth.userId,
        username: auth.username,
        limit: config.aliasLimit
      });
    } catch (error) {
      if (error instanceof ExchangeError) return sendExchangeError(reply, error);
      return sendInputError(reply, error);
    }
  });

  app.get("/api/onboarding", { preHandler: authenticate }, async (request) => {
    const auth = requireAuth(request);
    return {
      adminConsentUrl: buildAdminConsentUrl(auth.tenantId),
      setupScript: buildSetupScript()
    };
  });

  // Backward-compatible read endpoints. The PWA now uses /api/bootstrap so
  // accepted domains and aliases are fetched in one Exchange connection.
  app.get("/api/domains", { preHandler: authenticate }, async (request, reply) => {
    try {
      const query = request.query as { organization?: string };
      const organization = await verifiedOrganization(request, query.organization);
      const domains = await runExchange<ExchangeDomain[]>({
        action: "domains",
        organization
      });
      return { domains };
    } catch (error) {
      if (error instanceof ExchangeError) return sendExchangeError(reply, error);
      return sendInputError(reply, error);
    }
  });

  app.get("/api/aliases", { preHandler: authenticate }, async (request, reply) => {
    try {
      const auth = requireAuth(request);
      const query = request.query as { organization?: string };
      const organization = await verifiedOrganization(request, query.organization);
      return await runExchange<AliasSet>({
        action: "aliases",
        organization,
        userId: auth.userId,
        limit: config.aliasLimit
      });
    } catch (error) {
      if (error instanceof ExchangeError) return sendExchangeError(reply, error);
      return sendInputError(reply, error);
    }
  });

  app.post("/api/aliases", { preHandler: authenticate }, async (request, reply) => {
    try {
      const auth = requireAuth(request);
      const body = (request.body ?? {}) as {
        organization?: string;
        domain?: string;
      };
      const organization = await verifiedOrganization(request, body.organization);
      const domain = normalizeDomain(body.domain);
      const lockKey = `${auth.tenantId}:${auth.userId}`;

      return await withKeyedLock(lockKey, () =>
        runExchange<AliasSet>({
          action: "create",
          organization,
          userId: auth.userId,
          domain,
          limit: config.aliasLimit
        })
      );
    } catch (error) {
      if (error instanceof ExchangeError) return sendExchangeError(reply, error);
      return sendInputError(reply, error);
    }
  });

  app.delete("/api/aliases", { preHandler: authenticate }, async (request, reply) => {
    try {
      const auth = requireAuth(request);
      const body = (request.body ?? {}) as {
        organization?: string;
        address?: string;
      };
      const organization = await verifiedOrganization(request, body.organization);
      const address = normalizeAddress(body.address);
      const lockKey = `${auth.tenantId}:${auth.userId}`;

      return await withKeyedLock(lockKey, () =>
        runExchange<AliasSet>({
          action: "delete",
          organization,
          userId: auth.userId,
          limit: config.aliasLimit,
          address
        })
      );
    } catch (error) {
      if (error instanceof ExchangeError) return sendExchangeError(reply, error);
      return sendInputError(reply, error);
    }
  });
}
