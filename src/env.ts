export interface Env {
  DB: D1Database;

  // Configuration (vars)
  SITE_URL: string;
  LIBRARY?: string;
  ROOT_PATH?: string;
  ZIP_MAX_FILES?: string;
  ZIP_MAX_BYTES?: string;
  SYNC_TIME_BUDGET_MS?: string;
  CORS_ORIGINS?: string;

  // Credentials (secrets)
  TENANT_ID?: string;
  CLIENT_ID?: string;
  CLIENT_SECRET?: string;
  CERT_PRIVATE_KEY?: string;
  CERT_THUMBPRINT?: string;

  // Overrides, mainly for tests and sovereign clouds
  GRAPH_BASE_URL?: string;
  LOGIN_BASE_URL?: string;

  /** "true" serves a built-in fake SharePoint library instead of calling Graph. */
  MOCK_GRAPH?: string;
}

export interface Config {
  siteUrl: URL;
  /** Server-relative site path without trailing slash, e.g. "/sites/Projects" ("" for the root site). */
  sitePath: string;
  library: string;
  /** Drive-relative root folder path: "" or "/Clients/Acme". */
  rootPath: string;
  zipMaxFiles: number;
  zipMaxBytes: number;
  syncTimeBudgetMs: number;
  corsOrigins: string[];
  graphBaseUrl: string;
  loginBaseUrl: string;
  mock: boolean;
}

function int(value: string | undefined, fallback: number): number {
  const n = Number.parseInt(value ?? "", 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export function normalizeFolderPath(input: string | undefined): string {
  const segments = (input ?? "")
    .normalize("NFC")
    .replace(/\\/g, "/")
    .split("/")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return segments.length ? `/${segments.join("/")}` : "";
}

let cached: { key: Env; config: Config } | undefined;

export function getConfig(env: Env): Config {
  if (cached?.key === env) return cached.config;

  if (!env.SITE_URL) throw new Error("SITE_URL is not configured");
  const siteUrl = new URL(env.SITE_URL);
  const sitePath = decodeURIComponent(siteUrl.pathname).replace(/\/+$/, "");

  const config: Config = {
    siteUrl,
    sitePath,
    library: (env.LIBRARY ?? "").trim(),
    rootPath: normalizeFolderPath(env.ROOT_PATH),
    zipMaxFiles: int(env.ZIP_MAX_FILES, 500),
    zipMaxBytes: int(env.ZIP_MAX_BYTES, 2 * 1024 ** 3),
    syncTimeBudgetMs: int(env.SYNC_TIME_BUDGET_MS, 60_000),
    corsOrigins: (env.CORS_ORIGINS ?? "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean),
    graphBaseUrl: (env.GRAPH_BASE_URL || "https://graph.microsoft.com/v1.0").replace(/\/+$/, ""),
    loginBaseUrl: (env.LOGIN_BASE_URL || "https://login.microsoftonline.com").replace(/\/+$/, ""),
    mock: env.MOCK_GRAPH === "true",
  };
  cached = { key: env, config };
  return config;
}

/** Changing any of these invalidates the mirrored index. */
export function configFingerprint(config: Config): string {
  return JSON.stringify([config.siteUrl.origin, config.sitePath.toLowerCase(), config.library.toLowerCase()]);
}
