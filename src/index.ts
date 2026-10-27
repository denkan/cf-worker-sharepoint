import { app } from "./app";
import type { Env } from "./env";
import { deleteExpiredManifests } from "./api/zip";
import { runSync } from "./sync/delta";

export default {
  fetch: app.fetch,

  async scheduled(_controller: ScheduledController, env: Env, ctx: ExecutionContext) {
    const result = await runSync(env);
    if (result.status === "failed") console.error("Sync failed", result);
    else console.log("Sync", result);
    ctx.waitUntil(deleteExpiredManifests(env.DB));
  },
} satisfies ExportedHandler<Env>;
