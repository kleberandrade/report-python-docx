import { neon } from "@neondatabase/serverless";
import { readFile } from "node:fs/promises";
if (!process.env.DATABASE_URL)
  throw new Error("Set DATABASE_URL in the environment first");
const sql = neon(process.env.DATABASE_URL);
const source = await readFile(new URL("./schema.sql", import.meta.url), "utf8");
await sql.transaction(
  source
    .split(";")
    .map((statement) => statement.trim())
    .filter(Boolean)
    .map((statement) => sql.query(statement)),
);
console.log("Academic evaluation schema ready");
