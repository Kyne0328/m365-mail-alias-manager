import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { randomUUID } from "node:crypto";
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

export type WorkspaceSnapshot = {
  organization: string;
  domains: ExchangeDomain[];
  aliasSet: AliasSet;
};

type ExchangeAction =
  | {
      action: "bootstrap";
      tenantId: string;
      userId: string;
      username?: string;
      limit: number;
    }
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
  | { id: string; ok: true; data: T }
  | { id: string; ok: false; error: string; category?: string };

type PendingRequest = {
  resolve(value: unknown): void;
  reject(error: Error): void;
  timer: NodeJS.Timeout;
};

export class ExchangeError extends Error {
  readonly category?: string;

  constructor(message: string, category?: string) {
    super(message);
    this.name = "ExchangeError";
    this.category = category;
  }
}

let worker: ChildProcessWithoutNullStreams | null = null;
let stdoutBuffer = "";
let stderrTail = "";
const pending = new Map<string, PendingRequest>();

function rejectAll(error: Error) {
  for (const request of pending.values()) {
    clearTimeout(request.timer);
    request.reject(error);
  }
  pending.clear();
}

function consumeStdout(chunk: string) {
  stdoutBuffer += chunk;

  let newline = stdoutBuffer.indexOf("\n");
  while (newline >= 0) {
    const line = stdoutBuffer.slice(0, newline).replace(/\r$/, "");
    stdoutBuffer = stdoutBuffer.slice(newline + 1);

    if (line.startsWith(marker)) {
      try {
        const envelope = JSON.parse(line.slice(marker.length)) as ExchangeEnvelope<unknown>;
        const request = pending.get(envelope.id);
        if (request) {
          clearTimeout(request.timer);
          pending.delete(envelope.id);

          if (envelope.ok) {
            request.resolve(envelope.data);
          } else {
            request.reject(new ExchangeError(envelope.error, envelope.category));
          }
        }
      } catch {
        // Ignore non-protocol output. The worker can emit module warnings on stdout.
      }
    }

    newline = stdoutBuffer.indexOf("\n");
  }

  if (stdoutBuffer.length > 2_000_000) {
    stdoutBuffer = stdoutBuffer.slice(-2_000_000);
  }
}

function startWorker(): ChildProcessWithoutNullStreams {
  if (worker && !worker.killed && worker.exitCode === null) {
    return worker;
  }

  const scriptPath =
    process.env.EXCHANGE_SCRIPT_PATH ??
    path.resolve(process.cwd(), "scripts", "exchange.ps1");

  stdoutBuffer = "";
  stderrTail = "";

  const child = spawn(
    process.env.PWSH_PATH ?? "pwsh",
    ["-NoLogo", "-NoProfile", "-NonInteractive", "-File", scriptPath],
    {
      env: process.env,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true
    }
  );

  worker = child;
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");

  child.stdout.on("data", consumeStdout);
  child.stderr.on("data", (chunk: string) => {
    stderrTail += chunk;
    if (stderrTail.length > 200_000) {
      stderrTail = stderrTail.slice(-200_000);
    }
  });

  child.on("error", (error) => {
    if (worker === child) worker = null;
    rejectAll(new ExchangeError(`Unable to start PowerShell: ${error.message}`, "runtime"));
  });

  child.on("close", () => {
    if (worker === child) worker = null;
    const detail = stderrTail.trim();
    rejectAll(
      new ExchangeError(
        detail || "The Exchange worker stopped unexpectedly.",
        "runtime"
      )
    );
  });

  return child;
}

export function warmExchangeRuntime(): void {
  startWorker();
}

export async function runExchange<T>(input: ExchangeAction): Promise<T> {
  const child = startWorker();
  const id = randomUUID();

  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id);
      if (worker === child) {
        child.kill();
        worker = null;
      }
      reject(new ExchangeError("Exchange operation timed out.", "timeout"));
    }, 90_000);

    pending.set(id, {
      resolve: (value) => resolve(value as T),
      reject,
      timer
    });

    const payload = JSON.stringify({ id, input }) + "\n";
    child.stdin.write(payload, "utf8", (error) => {
      if (!error) return;
      const request = pending.get(id);
      if (!request) return;
      clearTimeout(request.timer);
      pending.delete(id);
      request.reject(
        new ExchangeError(`Unable to send request to Exchange worker: ${error.message}`, "runtime")
      );
    });
  });
}
