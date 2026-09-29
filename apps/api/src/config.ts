function required(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function optional(name: string, fallback = ""): string {
  return process.env[name]?.trim() || fallback;
}

export const config = {
  port: Number(optional("PORT", "8787")),
  clientId: required("ENTRA_CLIENT_ID"),
  apiAudience: required("API_AUDIENCE"),
  apiScope: optional("API_SCOPE", "access_as_user"),
  allowedOrigins: optional("ALLOWED_ORIGINS", "http://localhost:5173")
    .split(",")
    .map((value) => value.trim().replace(/\/$/, ""))
    .filter(Boolean),
  publicAppUrl: optional("PUBLIC_APP_URL", "http://localhost:5173").replace(/\/$/, ""),
  aliasLimit: 30
};

if (!Number.isInteger(config.port) || config.port < 1 || config.port > 65535) {
  throw new Error("PORT must be a valid TCP port.");
}
