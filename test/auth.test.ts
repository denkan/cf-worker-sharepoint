import { describe, expect, it } from "vitest";
import { createClientAssertion, getGraphToken } from "../src/graph/auth";
import { getConfig } from "../src/env";
import { env, mock } from "./helpers";

const b64urlDecode = (s: string) => {
  const b64 = s.replace(/-/g, "+").replace(/_/g, "/") + "=".repeat((4 - (s.length % 4)) % 4);
  return Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
};

async function makeKey() {
  const pair = (await crypto.subtle.generateKey(
    { name: "RSASSA-PKCS1-v1_5", modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: "SHA-256" },
    true,
    ["sign", "verify"],
  )) as CryptoKeyPair;
  const der = new Uint8Array((await crypto.subtle.exportKey("pkcs8", pair.privateKey)) as ArrayBuffer);
  const b64 = btoa(String.fromCharCode(...der)).replace(/(.{64})/g, "$1\n");
  return { pair, pem: `-----BEGIN PRIVATE KEY-----\n${b64}\n-----END PRIVATE KEY-----\n` };
}

describe("certificate client assertion", () => {
  it("produces a verifiable RS256 JWT with the right claims", async () => {
    const { pair, pem } = await makeKey();
    const sha1 = "A1B2C3D4E5F60718293A4B5C6D7E8F9012345678";
    const jwt = await createClientAssertion("https://login.example/tenant/oauth2/v2.0/token", "client-123", pem, sha1);
    const [h, p, s] = jwt.split(".");
    const header = JSON.parse(new TextDecoder().decode(b64urlDecode(h!)));
    const payload = JSON.parse(new TextDecoder().decode(b64urlDecode(p!)));
    expect(header).toMatchObject({ alg: "RS256", typ: "JWT" });
    expect(Array.from(b64urlDecode(header.x5t)).map((b) => b.toString(16).padStart(2, "0")).join("")).toBe(sha1.toLowerCase());
    expect(payload).toMatchObject({ aud: "https://login.example/tenant/oauth2/v2.0/token", iss: "client-123", sub: "client-123" });
    expect(payload.exp - payload.iat).toBe(600);
    const ok = await crypto.subtle.verify(
      "RSASSA-PKCS1-v1_5",
      pair.publicKey,
      b64urlDecode(s!),
      new TextEncoder().encode(`${h}.${p}`),
    );
    expect(ok).toBe(true);
  });

  it("uses x5t#S256 for SHA-256 thumbprints and accepts escaped newlines", async () => {
    const { pem } = await makeKey();
    const jwt = await createClientAssertion("aud", "id", pem.replace(/\n/g, "\\n"), "ab".repeat(32));
    const header = JSON.parse(new TextDecoder().decode(b64urlDecode(jwt.split(".")[0]!)));
    expect(header["x5t#S256"]).toBeDefined();
    expect(header.x5t).toBeUndefined();
  });

  it("is sent instead of a client secret when a certificate is configured", async () => {
    const { pem } = await makeKey();
    const certEnv = { ...env, TENANT_ID: "t", CLIENT_ID: "c", CERT_PRIVATE_KEY: pem, CERT_THUMBPRINT: "ab".repeat(20) };
    let sentAssertion = false;
    const spy: typeof fetch = async (input, init) => {
      const form = new URLSearchParams(String(init?.body));
      sentAssertion = Boolean(form.get("client_assertion")) && !form.get("client_secret");
      return mock.handle(input, init);
    };
    expect(await getGraphToken(certEnv, getConfig(certEnv), spy)).toBe("mock-access-token");
    expect(sentAssertion).toBe(true);
  });
});
