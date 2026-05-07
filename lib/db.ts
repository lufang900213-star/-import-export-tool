import { Pool } from "pg";
import type { PersistedRecord } from "./types";

const TABLE = process.env.DB_TABLE || "waybill_records";

let pool: Pool | null = null;

function getPool(): Pool {
  if (pool) return pool;

  const connectionString = process.env.DATABASE_URL;
  if (!connectionString) {
    throw new Error("数据库未配置，请在 Vercel 环境变量中设置 DATABASE_URL");
  }

  pool = new Pool({
    connectionString,
    ssl: connectionString.includes("sslmode=require") ? { rejectUnauthorized: false } : undefined,
  });

  return pool;
}

export async function listRecords(): Promise<PersistedRecord[]> {
  const client = await getPool().connect();
  try {
    const result = await client.query<PersistedRecord>(
      `select * from ${TABLE} order by submittedAt desc`,
    );
    return result.rows;
  } finally {
    client.release();
  }
}

export async function insertRecords(rows: PersistedRecord[]): Promise<PersistedRecord[]> {
  const client = await getPool().connect();
  try {
    await client.query("begin");
    const query = `
      insert into ${TABLE} (
        id,
        templateId,
        templateName,
        submittedAt,
        externalCode,
        receiverName,
        receiverPhone,
        senderName,
        senderPhone,
        senderAddress,
        receiverAddress,
        weight,
        quantity,
        temperature,
        remark
      ) values (
        $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15
      )
      on conflict (id) do update set
        templateId = excluded.templateId,
        templateName = excluded.templateName,
        submittedAt = excluded.submittedAt,
        externalCode = excluded.externalCode,
        receiverName = excluded.receiverName,
        receiverPhone = excluded.receiverPhone,
        senderName = excluded.senderName,
        senderPhone = excluded.senderPhone,
        senderAddress = excluded.senderAddress,
        receiverAddress = excluded.receiverAddress,
        weight = excluded.weight,
        quantity = excluded.quantity,
        temperature = excluded.temperature,
        remark = excluded.remark
    `;

    for (const row of rows) {
      await client.query(query, [
        row.id,
        row.templateId,
        row.templateName,
        row.submittedAt,
        row.externalCode,
        row.receiverName,
        row.receiverPhone,
        row.senderName,
        row.senderPhone,
        row.senderAddress,
        row.receiverAddress,
        row.weight,
        row.quantity,
        row.temperature,
        row.remark,
      ]);
    }

    await client.query("commit");
    const result = await client.query<PersistedRecord>(`select * from ${TABLE} order by submittedAt desc`);
    return result.rows;
  } catch (error) {
    await client.query("rollback");
    throw error;
  } finally {
    client.release();
  }
}
