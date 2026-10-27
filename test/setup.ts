import { applyD1Migrations, reset, type D1Migration } from "cloudflare:test";
import { env as testEnv } from "cloudflare:workers";
import type { Env } from "../src/env";
import { beforeEach } from "vitest";
import { mockSharePoint } from "../src/mock/graph-mock";
import { clearDriveContextCache } from "../src/sharepoint/drive";
import { clearTokenCache } from "../src/graph/auth";

beforeEach(async () => {
  await reset();
  const env = testEnv as unknown as Env & { TEST_MIGRATIONS: D1Migration[] };
  await applyD1Migrations(env.DB, env.TEST_MIGRATIONS);
  mockSharePoint.reset();
  clearDriveContextCache();
  clearTokenCache();
});
