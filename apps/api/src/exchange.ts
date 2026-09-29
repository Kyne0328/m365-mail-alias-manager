import { spawn } from "node:child_process";
import path from "node:path";

const marker = "__ALIAS_MANAGER_JSON__";

export type ExchangeDomain = {
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

type ExchangeAction =
  | { action: "domains"; organization: string }
  | {
      action: "aliases";
      organization: string;
      userId: string;
      limit: number;
    }
  | {
      action: "create";
      organization: string;
      userId: string;
      domain: string;
      prefix: string;
      limit: number;
    }
  | {
      action: "delete";
      organization: string;
      userId: string;
      limit: number;
      address: string;
    };

type ExchangeEnvelope<T> =
  | { ok: true; data: T }
  | { ok: false; error: string; category?: string };

export class ExchangeError extends Error {
  readonly category?: string;

  constructor(message: string, category?: string) {
    super(message);
    this.name = "ExchangeError";
    this.category = category;
  }
}

export async function runExchange<T>(input: ExchangeAction): Promise<T> {
  const scriptPath =
    process.env.EXCHANGE_SCRIPT_PATH ??
    path.resolve(process.cwd(), "scripts", "exchange.ps1");

  return new Promise<T>((resolve, reject) => {
    const child = spawn(
      process.env.PWSH_PATH ?? "pwsh",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-File", scriptPath],
      {
        env: process.env,
        stdio: ["pipe", "pipe", "pipe"],
        windowsHide: true
      }
    );

    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new ExchangeError("Exchange operation timed out.", "timeout"));
    }, 90_000);

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
      if (stdout.length > 2_000_000) stdout = stdout.slice(-2_000_000);
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
      if (stderr.length > 200_000) stderr = stderr.slice(-200_000);
    });

    child.on("error", (error) => {
      clearTimeout(timer);
      reject(new ExchangeError(`Unable to start PowerShell: ${error.message}`, "runtime"));
    });

    child.on("close", () => {
      clearTimeout(timer);
      const line = stdout
        .split(/\r?\n/)
        .reverse()
        .find((entry) => entry.startsWith(marker));

      if (!line) {
        reject(
          new ExchangeError(
            stderr.trim() || "Exchange returned no structured response.",
            "runtime"
          )
        );
        return;
      }

      try {
        const envelope = JSON.parse(line.slice(marker.length)) as ExchangeEnvelope<T>;
        if (!envelope.ok) {
          reject(new ExchangeError(envelope.error, envelope.category));
          return;
        }
        resolve(envelope.data);
      } catch (error) {
        reject(
          new ExchangeError(
            `Unable to parse the Exchange response: ${error instanceof Error ? error.message : String(error)}`,
            "runtime"
          )
        );
      }
    });

    child.stdin.end(JSON.stringify(input));
  });
}
