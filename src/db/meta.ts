export async function getMeta(db: D1Database, key: string): Promise<string | null> {
  const row = await db.prepare("SELECT value FROM meta WHERE key = ?").bind(key).first<{ value: string | null }>();
  return row?.value ?? null;
}

export async function getMetaMany<K extends string>(db: D1Database, keys: readonly K[]): Promise<Record<K, string | null>> {
  const placeholders = keys.map(() => "?").join(",");
  const { results } = await db
    .prepare(`SELECT key, value FROM meta WHERE key IN (${placeholders})`)
    .bind(...keys)
    .all<{ key: string; value: string | null }>();
  const out = Object.fromEntries(keys.map((k) => [k, null])) as Record<K, string | null>;
  for (const row of results) out[row.key as K] = row.value;
  return out;
}

export function setMetaStmt(db: D1Database, key: string, value: string | null): D1PreparedStatement {
  return db
    .prepare("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT (key) DO UPDATE SET value = excluded.value")
    .bind(key, value);
}

export async function setMeta(db: D1Database, entries: Record<string, string | null>): Promise<void> {
  const stmts = Object.entries(entries).map(([k, v]) => setMetaStmt(db, k, v));
  if (stmts.length) await db.batch(stmts);
}
