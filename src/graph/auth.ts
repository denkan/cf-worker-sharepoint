import type { Config, Env } from "../env";
import { ApiError } from "../errors";

type Fetcher = typeof fetch;

interface CachedToken {
  key: string;
  accessToken: string;
  expiresAt: number;
}

/** Refresh this long before the token actually expires. */
const REFRESH_MARGIN_MS = 5 * 60 * 1000;

let cachedToken: CachedToken | undefined;
let inflight: { key: string; promise: Promise<CachedToken> } | undefined;

/**
 * App-only (client credentials) token for Microsoft Graph, cached per isolate.
 * Supports a client secret or a certificate (signed client assertion).
 */
export async function getGraphToken(env: Env, config: Config, fetcher: Fetcher): Promise<string> {
  if (config.mock && !env.CLIENT_SECRET && !env.CERT_PRIVATE_KEY) {
    env = { ...env, TENANT_ID: "mock-tenant", CLIENT_ID: "mock-client", CLIENT_SECRET: "mock-secret" };
  }
  const tenantId = env.TENANT_ID;
  const clientId = env.CLIENT_ID;
  if (!tenantId || !clientId) {
    throw new ApiError(500, "INTERNAL_ERROR", "TENANT_ID and CLIENT_ID must be configured");
  }
  const key = `${tenantId}:${clientId}`;
  const now = Date.now();

  if (cachedToken?.key === key && cachedToken.expiresAt - REFRESH_MARGIN_MS > now) {
    return cachedToken.accessToken;
  }
  if (inflight?.key !== key) {
    const promise = requestToken(env, config, fetcher, tenantId, clientId).then(
      (token) => {
        cachedToken = { ...token, key };
        inflight = undefined;
        return cachedToken;
      },
      (err) => {
        inflight = undefined;
        throw err;
      },
    );
    inflight = { key, promise };
  }
  return (await inflight.promise).accessToken;
}

export function clearTokenCache() {
  cachedToken = undefined;
  inflight = undefined;
}

async function requestToken(
  env: Env,
  config: Config,
  fetcher: Fetcher,
  tenantId: string,
  clientId: string,
): Promise<Omit<CachedToken, "key">> {
  const tokenUrl = `${config.loginBaseUrl}/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const form = new URLSearchParams({
    client_id: clientId,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  if (env.CERT_PRIVATE_KEY && env.CERT_THUMBPRINT) {
    form.set("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
    form.set(
      "client_assertion",
      await createClientAssertion(tokenUrl, clientId, env.CERT_PRIVATE_KEY, env.CERT_THUMBPRINT),
    );
  } else if (env.CLIENT_SECRET) {
    form.set("client_secret", env.CLIENT_SECRET);
  } else {
    throw new ApiError(
      500,
      "INTERNAL_ERROR",
      "Configure either CLIENT_SECRET or CERT_PRIVATE_KEY + CERT_THUMBPRINT",
    );
  }

  const started = Date.now();
  const res = await fetcher(tokenUrl, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: form,
  });
  const body = (await res.json().catch(() => ({}))) as {
    access_token?: string;
    expires_in?: number;
    error?: string;
    error_description?: string;
  };
  if (!res.ok || !body.access_token) {
    throw new ApiError(
      502,
      "UPSTREAM_AUTH_ERROR",
      `Could not obtain a Graph token: ${body.error ?? res.status} ${body.error_description?.split("\r\n")[0] ?? ""}`.trim(),
    );
  }
  return {
    accessToken: body.access_token,
    expiresAt: started + (body.expires_in ?? 3600) * 1000,
  };
}

// --- Certificate-based client assertion (RFC 7523) -------------------------

const b64url = (bytes: Uint8Array) =>
  btoa(String.fromCharCode(...bytes))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

const b64urlJson = (value: unknown) => b64url(new TextEncoder().encode(JSON.stringify(value)));

function hexToBytes(hex: string): Uint8Array {
  const clean = hex.replace(/[^0-9a-f]/gi, "");
  const out = new Uint8Array(clean.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = Number.parseInt(clean.slice(i * 2, i * 2 + 2), 16);
  return out;
}

function pemToDer(pem: string): ArrayBuffer {
  if (/BEGIN RSA PRIVATE KEY/.test(pem)) {
    throw new ApiError(
      500,
      "INTERNAL_ERROR",
      "CERT_PRIVATE_KEY must be PKCS#8 (BEGIN PRIVATE KEY). Convert with: openssl pkcs8 -topk8 -nocrypt -in key.pem",
    );
  }
  const base64 = pem
    .replace(/\\n/g, "\n")
    .replace(/-----(BEGIN|END) [A-Z ]+-----/g, "")
    .replace(/\s+/g, "");
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

let cachedKey: { pem: string; key: CryptoKey } | undefined;

async function importKey(pem: string): Promise<CryptoKey> {
  if (cachedKey?.pem === pem) return cachedKey.key;
  const key = await crypto.subtle.importKey(
    "pkcs8",
    pemToDer(pem),
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );
  cachedKey = { pem, key };
  return key;
}

export async function createClientAssertion(
  audience: string,
  clientId: string,
  privateKeyPem: string,
  thumbprintHex: string,
): Promise<string> {
  const thumbprint = hexToBytes(thumbprintHex);
  // Entra accepts the SHA-1 thumbprint as x5t, or a SHA-256 one as x5t#S256.
  const header =
    thumbprint.length === 32
      ? { alg: "RS256", typ: "JWT", "x5t#S256": b64url(thumbprint) }
      : { alg: "RS256", typ: "JWT", x5t: b64url(thumbprint) };
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    aud: audience,
    iss: clientId,
    sub: clientId,
    jti: crypto.randomUUID(),
    nbf: now - 30,
    iat: now,
    exp: now + 600,
  };
  const signingInput = `${b64urlJson(header)}.${b64urlJson(payload)}`;
  const signature = await crypto.subtle.sign(
    "RSASSA-PKCS1-v1_5",
    await importKey(privateKeyPem),
    new TextEncoder().encode(signingInput),
  );
  return `${signingInput}.${b64url(new Uint8Array(signature))}`;
}
