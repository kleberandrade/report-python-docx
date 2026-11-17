import { neon } from "@neondatabase/serverless";
import { createHash, timingSafeEqual } from "node:crypto";
import type { Evaluation, Stage } from "./evaluations";
import { parseResponse } from "./evaluations";

function database() {
  if (!process.env.DATABASE_URL) throw new Error("Database not configured");
  return neon(process.env.DATABASE_URL);
}
export async function saveResponse(item: Evaluation) {
  const sql = database();
  await sql`INSERT INTO class_evaluations (id, stage, payload)
    VALUES (${item.id}, ${item.stage}, ${JSON.stringify(item)}::jsonb)
    ON CONFLICT (id) DO NOTHING`;
}
export async function readResponses(stage: Stage): Promise<Evaluation[]> {
  const sql = database();
  const rows =
    await sql`SELECT payload FROM class_evaluations WHERE stage = ${stage}
    ORDER BY created_at, id LIMIT 10001`;
  if (rows.length > 10000) throw new Error("Export limit reached");
  return rows.map((row) => parseResponse(row.payload));
}
export function authorized(secret: string | null): boolean {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected || expected.length < 24 || !secret) return false;
  return timingSafeEqual(
    createHash("sha256").update(secret).digest(),
    createHash("sha256").update(expected).digest(),
  );
}
