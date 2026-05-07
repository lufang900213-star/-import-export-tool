import { NextResponse } from "next/server";
import { insertRecords } from "../../../lib/db";
import type { PersistedRecord } from "../../../lib/types";

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as { rows?: PersistedRecord[] };
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const data = await insertRecords(rows);

    return NextResponse.json({
      data,
      successCount: rows.length,
      failedCount: 0,
    });
  } catch (error) {
    return NextResponse.json(
      {
        data: [],
        successCount: 0,
        failedCount: 0,
        error: error instanceof Error ? error.message : "提交失败",
      },
      { status: 500 },
    );
  }
}
