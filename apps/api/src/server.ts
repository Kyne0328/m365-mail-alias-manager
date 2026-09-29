import Fastify from "fastify";
import cors from "@fastify/cors";
import helmet from "@fastify/helmet";
import rateLimit from "@fastify/rate-limit";
import { config } from "./config.js";
import { registerRoutes } from "./routes.js";

const app = Fastify({
  logger: true,
  trustProxy: true,
  bodyLimit: 16 * 1024
});

await app.register(helmet, {
  contentSecurityPolicy: false
});

await app.register(cors, {
  credentials: false,
  allowedHeaders: ["Authorization", "Content-Type"],
  methods: ["GET", "POST", "DELETE", "OPTIONS"],
  origin(origin, callback) {
    if (!origin) {
      callback(null, true);
      return;
    }

    const normalized = origin.replace(/\/$/, "");
    callback(null, config.allowedOrigins.includes(normalized));
  }
});

await app.register(rateLimit, {
  max: 120,
  timeWindow: "1 minute"
});

app.get("/health", async () => ({
  ok: true,
  service: "m365-mail-alias-manager-api"
}));

await registerRoutes(app);

app.setErrorHandler((error, request, reply) => {
  request.log.error({ err: error }, "Unhandled API error");
  void reply.code(500).send({
    code: "SERVER_ERROR",
    message: "The server encountered an unexpected error."
  });
});

await app.listen({
  port: config.port,
  host: "0.0.0.0"
});
